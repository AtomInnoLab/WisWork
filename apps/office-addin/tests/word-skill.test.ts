// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from 'vitest'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
import {
  BrowserWordAdapter,
  type WordAdapter,
  type WordOoxmlResult,
} from '../src/skills/word/browser-word-adapter.js'
import { createWordSkill } from '../src/skills/word/word-skill.js'

function adapter(overrides: Partial<WordAdapter> = {}): WordAdapter {
  return {
    getDocumentSnapshot: vi.fn().mockResolvedValue({ text: 'Title\nFirst', fingerprint: 'before' }),
    getDocumentText: vi.fn().mockResolvedValue({
      totalParagraphs: 2,
      totalParagraphsExact: true,
      hasMore: false,
      showing: { start: 0, end: 2 },
      paragraphs: [
        { index: 0, text: 'Title', style: 'Heading 1', alignment: 'Centered' },
        { index: 1, text: 'First', style: 'List Paragraph', listLevel: 0, listString: '1.' },
      ],
    }),
    getDocumentStructure: vi.fn().mockResolvedValue({
      paragraphCount: 2,
      sectionCount: 1,
      tableCount: 0,
      contentControlCount: 0,
      truncated: { paragraphs: false, sections: false, tables: false, contentControls: false },
      headings: [{ text: 'Title', level: 1, paragraphIndex: 0 }],
      tables: [],
      contentControls: [],
    }),
    getOoxml: vi.fn().mockResolvedValue({
      xml: '<w:body><w:p><w:r><w:t>Title</w:t></w:r></w:p></w:body>',
      children: [{ index: 0, type: 'p', line: 1, paragraphIndex: 0, text: 'Title' }],
    } satisfies WordOoxmlResult),
    screenshotDocument: vi.fn().mockRejectedValue(new Error('office_api_unsupported')),
    fingerprint: vi.fn().mockResolvedValue('before'),
    executeOperations: vi.fn().mockResolvedValue(undefined),
    verifyOperations: vi.fn().mockResolvedValue(true),
    executeDocumentWrite: vi.fn().mockResolvedValue(undefined),
    verifyDocumentWrite: vi.fn().mockResolvedValue(true),
    ...overrides,
  }
}

function call(name: string, input: Record<string, unknown> = {}) {
  return { id: 'call-1', name, input }
}

describe('Word compatibility skill', () => {
  it('exposes the exact host inventory without shared or other-host tools', () => {
    const skill = createWordSkill({
      adapter: adapter(),
      vfs: new InMemoryVfs(),
      proposals: createStructuredProposalController(),
    })
    expect(skill.tools.map((tool) => tool.name)).toEqual([
      'get_document_text',
      'get_document_structure',
      'get_ooxml',
      'screenshot_document',
      'write_document',
      'execute_office_js',
    ])
    expect(skill.tools.map((tool) => ({ name: tool.name, schema: tool.inputSchema }))).toEqual([
      {
        name: 'get_document_text',
        schema: {
          type: 'object',
          properties: {
            startParagraph: { type: 'integer', minimum: 0, maximum: 1_000_000 },
            endParagraph: { type: 'integer', minimum: 0, maximum: 1_000_000 },
            includeFormatting: { type: 'boolean' },
          },
          required: [],
          additionalProperties: false,
        },
      },
      {
        name: 'get_document_structure',
        schema: { type: 'object', properties: {}, required: [], additionalProperties: false },
      },
      {
        name: 'get_ooxml',
        schema: {
          type: 'object',
          properties: {
            startChild: { type: 'integer', minimum: 0, maximum: 1_000_000 },
            endChild: { type: 'integer', minimum: 0, maximum: 1_000_000 },
          },
          required: [],
          additionalProperties: false,
        },
      },
      {
        name: 'screenshot_document',
        schema: {
          type: 'object',
          properties: {
            page: { type: 'integer', minimum: 1, maximum: 100_000 },
            explanation: { type: 'string', maxLength: 50 },
          },
          required: [],
          additionalProperties: false,
        },
      },
      {
        name: 'write_document',
        schema: {
          type: 'object',
          properties: {
            text: { type: 'string', minLength: 1, maxLength: 12_000 },
            mode: { type: 'string', enum: ['replace', 'append', 'prepend'] },
            explanation: { type: 'string', maxLength: 100 },
            format: { type: 'string', enum: ['markdown', 'plain_text'] },
          },
          required: ['text', 'mode'],
          additionalProperties: false,
        },
      },
      {
        name: 'execute_office_js',
        schema: {
          type: 'object',
          properties: {
            code: { type: 'string', minLength: 1, maxLength: 32 * 1024 },
            explanation: { type: 'string', maxLength: 100 },
          },
          required: ['code'],
          additionalProperties: false,
        },
      },
    ])
  })

  it('returns bounded normalized text and rejects unknown or invalid fields', async () => {
    const fake = adapter()
    const skill = createWordSkill({
      adapter: fake,
      vfs: new InMemoryVfs(),
      proposals: createStructuredProposalController(),
    })
    await expect(
      skill.executeTool(
        call('get_document_text', {
          startParagraph: 0,
          endParagraph: 2,
          includeFormatting: true,
        }),
      ),
    ).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('"listString":"1."'),
    })
    expect(fake.getDocumentText).toHaveBeenCalledWith(
      { startParagraph: 0, endParagraph: 2, includeFormatting: true },
      undefined,
    )
    await expect(
      skill.executeTool(call('get_document_text', { startParagraph: -1 })),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true })
    await expect(
      skill.executeTool(call('get_document_text', { unknown: true })),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true })
  })

  it('writes bounded OOXML to the shared VFS and returns only its mapping summary', async () => {
    const vfs = new InMemoryVfs()
    const skill = createWordSkill({
      adapter: adapter(),
      vfs,
      proposals: createStructuredProposalController(),
    })
    const result = await skill.executeTool(call('get_ooxml', { startChild: 0, endChild: 0 }))
    expect(result).toMatchObject({ mutated: false })
    expect(JSON.parse(result.output)).toMatchObject({
      file: '/home/user/ooxml/body-0-0.xml',
      children: [{ index: 0, type: 'p', paragraphIndex: 0 }],
    })
    expect(vfs.readText('/home/user/ooxml/body-0-0.xml')).toContain('<w:t>Title</w:t>')
  })

  it('maps unsupported APIs and Office failures to stable safe errors', async () => {
    const unsupported = createWordSkill({
      adapter: adapter({
        getDocumentStructure: vi.fn().mockRejectedValue(new Error('office_api_unsupported')),
      }),
      vfs: new InMemoryVfs(),
      proposals: createStructuredProposalController(),
    })
    await expect(unsupported.executeTool(call('get_document_structure'))).resolves.toMatchObject({
      output: 'office_api_unsupported',
      isError: true,
    })

    const failed = createWordSkill({
      adapter: adapter({ getDocumentText: vi.fn().mockRejectedValue(new Error('secret body')) }),
      vfs: new InMemoryVfs(),
      proposals: createStructuredProposalController(),
    })
    await expect(failed.executeTool(call('get_document_text'))).resolves.toMatchObject({
      output: 'office_read_failed',
      isError: true,
    })
  })

  it('rejects an adapter result that exceeds the model-output bound', async () => {
    const skill = createWordSkill({
      adapter: adapter({
        getDocumentText: vi.fn().mockResolvedValue({
          totalParagraphs: 1,
          showing: { start: 0, end: 1 },
          paragraphs: [{ index: 0, text: 'x'.repeat(300_000) }],
        }),
      }),
      vfs: new InMemoryVfs(),
      proposals: createStructuredProposalController(),
    })
    await expect(skill.executeTool(call('get_document_text'))).resolves.toMatchObject({
      output: 'office_read_failed',
      isError: true,
    })
  })

  it('fails screenshots closed and propagates cancellation', async () => {
    const controller = new AbortController()
    controller.abort()
    const fake = adapter()
    const skill = createWordSkill({
      adapter: fake,
      vfs: new InMemoryVfs(),
      proposals: createStructuredProposalController(),
    })
    await expect(
      skill.executeTool(call('screenshot_document', { page: 1 }), controller.signal),
    ).resolves.toMatchObject({ output: 'cancelled', isError: true })
    expect(fake.screenshotDocument).not.toHaveBeenCalled()
    await expect(
      skill.executeTool(call('screenshot_document', { page: 1 })),
    ).resolves.toMatchObject({ output: 'office_api_unsupported', isError: true })
  })

  it('bounds and validates a successful screenshot result', async () => {
    const skill = createWordSkill({
      adapter: adapter({
        screenshotDocument: vi.fn().mockResolvedValue({
          base64: 'not base64!',
          mime: 'image/png',
        }),
      }),
      vfs: new InMemoryVfs(),
      proposals: createStructuredProposalController(),
    })
    await expect(skill.executeTool(call('screenshot_document'))).resolves.toMatchObject({
      output: 'office_read_failed',
      isError: true,
    })
  })

  it('returns a valid Word screenshot as model-visible image data', async () => {
    const base64 = 'iVBORw0KGgoAAAA='
    const skill = createWordSkill({
      adapter: adapter({
        screenshotDocument: vi.fn().mockResolvedValue({ base64, mime: 'image/png' }),
      }),
      vfs: new InMemoryVfs(),
      proposals: createStructuredProposalController(),
    })
    await expect(
      skill.executeTool(call('screenshot_document', { page: 1 })),
    ).resolves.toMatchObject({
      output: '{"mime":"image/png"}',
      modelContent: [{ type: 'image', image: { mime: 'image/png', base64 } }],
      display: { kind: 'images', items: [{ url: `data:image/png;base64,${base64}` }] },
      mutated: false,
    })
  })

  it('does not write OOXML when the success summary exceeds its bound', async () => {
    const vfs = new InMemoryVfs()
    const skill = createWordSkill({
      adapter: adapter({
        getOoxml: vi.fn().mockResolvedValue({
          xml: '<w:body/>',
          children: Array.from({ length: 5_000 }, (_, index) => ({
            index,
            type: 'p',
            line: index,
            text: 'x'.repeat(100),
          })),
        }),
      }),
      vfs,
      proposals: createStructuredProposalController(),
    })
    await expect(skill.executeTool(call('get_ooxml'))).resolves.toMatchObject({
      output: 'office_read_failed',
      isError: true,
    })
    expect(() => vfs.readText('/home/user/ooxml/body.xml')).toThrow('vfs_not_found')
  })

  it('turns an allowlisted declarative Word program into a confirmation-gated mutation', async () => {
    const proposals = createStructuredProposalController()
    const fake = adapter()
    const skill = createWordSkill({ adapter: fake, vfs: new InMemoryVfs(), proposals })
    const code = JSON.stringify({
      version: 1,
      operations: [{ op: 'insert_text', location: 'end', text: 'hello' }],
    })
    await expect(skill.executeTool(call('execute_office_js', { code }))).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('proposalId'),
    })
    const proposal = proposals.pending()!
    expect(proposal).toMatchObject({ code, impact: { host: 'word', count: 1 } })
    expect(fake.executeOperations).not.toHaveBeenCalled()
    await proposals.confirm(proposal.id)
    expect(fake.executeOperations).toHaveBeenCalledWith(
      [{ op: 'insert_text', location: 'end', text: 'hello' }],
      expect.any(AbortSignal),
    )
  })

  it('offers a direct structured writing tool and treats a pending proposal as success', async () => {
    const proposals = createStructuredProposalController()
    const fake = adapter()
    const skill = createWordSkill({ adapter: fake, vfs: new InMemoryVfs(), proposals })

    const first = await skill.executeTool(
      call('write_document', {
        text: 'Large language models generate text from learned patterns.',
        mode: 'replace',
        explanation: 'Write an introduction to LLMs',
      }),
    )
    expect(JSON.parse(first.output)).toMatchObject({
      status: 'awaiting_user_confirmation',
      mutated: false,
    })
    expect(first).toMatchObject({ mutated: false, summary: 'Awaiting confirmation' })
    const pending = proposals.pending()!
    expect(pending).toMatchObject({
      toolName: 'write_document',
      title: 'Write an introduction to LLMs',
      before: 'Title\nFirst',
      after: 'Large language models generate text from learned patterns.',
      preview: {
        mode: 'replace',
      },
    })

    const duplicate = await skill.executeTool(
      call('write_document', { text: 'A duplicate attempt', mode: 'append' }),
    )
    expect(JSON.parse(duplicate.output)).toEqual({
      proposalId: pending.id,
      status: 'awaiting_user_confirmation',
      mutated: false,
    })
    expect(proposals.pending()?.id).toBe(pending.id)
    expect(fake.fingerprint).not.toHaveBeenCalled()
    expect(fake.getDocumentSnapshot).toHaveBeenCalledWith(undefined)

    await proposals.confirm(pending.id)
    expect(fake.fingerprint).toHaveBeenCalledOnce()
    expect(fake.executeDocumentWrite).toHaveBeenCalledWith(
      expect.objectContaining({
        mode: 'replace',
        semanticText: 'Large language models generate text from learned patterns.',
        blocks: [
          {
            type: 'paragraph',
            spans: [{ text: 'Large language models generate text from learned patterns.' }],
          },
        ],
      }),
      expect.any(AbortSignal),
    )
  })

  it('keeps long-document comparisons bounded while preserving the edited edge', async () => {
    const proposals = createStructuredProposalController()
    const skill = createWordSkill({
      adapter: adapter({
        getDocumentSnapshot: vi
          .fn()
          .mockResolvedValue({ text: `START${'x'.repeat(100_000)}END`, fingerprint: 'before' }),
      }),
      vfs: new InMemoryVfs(),
      proposals,
    })

    const result = await skill.executeTool(
      call('write_document', { text: 'APPENDED', mode: 'append' }),
    )
    expect(result).not.toHaveProperty('isError')
    expect(proposals.pending()).toMatchObject({
      before: expect.stringMatching(/^…\n.*END$/s),
      after: expect.stringMatching(/^…\n.*END\nAPPENDED$/s),
    })
    expect(String(proposals.pending()?.before).length).toBeLessThanOrEqual(8_002)
  })

  it('accepts the maximum non-ASCII draft without exceeding the proposal byte budget', async () => {
    const proposals = createStructuredProposalController()
    const skill = createWordSkill({
      adapter: adapter({
        getDocumentSnapshot: vi.fn().mockResolvedValue({
          text: '旧'.repeat(20_000),
          fingerprint: 'before',
        }),
      }),
      vfs: new InMemoryVfs(),
      proposals,
    })
    const draft = '新'.repeat(12_000)

    const result = await skill.executeTool(call('write_document', { text: draft, mode: 'replace' }))

    expect(result).not.toHaveProperty('isError')
    expect(proposals.pending()?.after).toBe(draft)
  })

  it('rejects JavaScript and unallowlisted declarative operations without a proposal', async () => {
    const proposals = createStructuredProposalController()
    const skill = createWordSkill({ adapter: adapter(), vfs: new InMemoryVfs(), proposals })
    for (const code of [
      'context.document.body.insertText("x", "End")',
      '{"version":1,"operations":[{"op":"fetch","url":"https://example.com"}]}',
    ]) {
      await expect(skill.executeTool(call('execute_office_js', { code }))).resolves.toMatchObject({
        output: 'invalid_tool_input',
        isError: true,
        mutated: false,
      })
    }
    expect(proposals.pending()).toBeUndefined()
  })

  it('refuses a stale declarative Word proposal before execution', async () => {
    const fake = adapter({
      fingerprint: vi.fn().mockResolvedValueOnce('before').mockResolvedValueOnce('changed'),
    })
    const proposals = createStructuredProposalController()
    const skill = createWordSkill({ adapter: fake, vfs: new InMemoryVfs(), proposals })
    await skill.executeTool(
      call('execute_office_js', {
        code: '{"version":1,"operations":[{"op":"insert_text","location":"end","text":"x"}]}',
      }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(fake.executeOperations).not.toHaveBeenCalled()
  })
})

describe('browser Word adapter', () => {
  const originals = { Office: globalThis.Office, Word: globalThis.Word }

  afterEach(() => {
    Object.assign(globalThis, originals)
  })

  it('checks the active host and Word API set before entering Word.run', async () => {
    const run = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'Excel',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: { run },
    })
    await expect(new BrowserWordAdapter().getDocumentText({}, undefined)).rejects.toThrow(
      'office_api_unsupported',
    )
    expect(run).not.toHaveBeenCalled()

    ;(globalThis.Office.context.requirements.isSetSupported as ReturnType<typeof vi.fn>) = vi
      .fn()
      .mockReturnValue(false)
    globalThis.Office.context.host = 'Word' as unknown as Office.HostType
    await expect(new BrowserWordAdapter().getDocumentText({}, undefined)).rejects.toThrow(
      'office_api_unsupported',
    )
    expect(run).not.toHaveBeenCalled()
  })

  it('exports bounded PDF slices and renders the requested page to PNG', async () => {
    const closeAsync = vi.fn((callback: () => void) => callback())
    const getSliceAsync = vi.fn((index: number, callback: (result: unknown) => void) =>
      callback({ status: 'succeeded', value: { data: index === 0 ? [37, 80] : [68, 70] } }),
    )
    const getFileAsync = vi.fn(
      (_type: unknown, _options: unknown, callback: (result: unknown) => void) =>
        callback({
          status: 'succeeded',
          value: { size: 4, sliceCount: 2, getSliceAsync, closeAsync },
        }),
    )
    Object.assign(globalThis, {
      Office: {
        FileType: { Pdf: 'pdf' },
        context: {
          host: 'Word',
          document: { getFileAsync },
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: { run: vi.fn() },
    })
    const renderPage = vi.fn().mockResolvedValue('iVBORw0KGgoAAA=')
    await expect(new BrowserWordAdapter({ renderPage }).screenshotDocument(2)).resolves.toEqual({
      base64: 'iVBORw0KGgoAAA=',
      mime: 'image/png',
    })
    expect(renderPage).toHaveBeenCalledWith(Uint8Array.from([37, 80, 68, 70]), 2, undefined)
    expect(closeAsync).toHaveBeenCalledOnce()
  })

  it('normalizes paragraph style and list metadata through real Word.run request contexts', async () => {
    const sync = vi.fn().mockResolvedValue(undefined)
    const paragraphs = {
      items: [
        {
          text: 'Item',
          style: 'List Paragraph',
          alignment: 'Left',
          outlineLevel: 0,
          load: vi.fn(),
          listItemOrNullObject: { isNullObject: false, level: 1, listString: 'a.', load: vi.fn() },
        },
      ],
      load: vi.fn(),
    }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'Word',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ document: { body: { paragraphs } }, sync }),
      },
    })
    await expect(new BrowserWordAdapter().getDocumentText({}, undefined)).resolves.toEqual({
      totalParagraphs: 1,
      totalParagraphsExact: true,
      hasMore: false,
      showing: { start: 0, end: 1 },
      paragraphs: [
        {
          index: 0,
          text: 'Item',
          style: 'List Paragraph',
          alignment: 'Left',
          listLevel: 1,
          listString: 'a.',
        },
      ],
    })
    expect(paragraphs.load).toHaveBeenCalledWith({ $skip: 0, $top: 501 })
    expect(sync).toHaveBeenCalledTimes(2)
  })

  it('keeps Word fingerprints stable across volatile OOXML metadata but detects real edits', async () => {
    const flatOpc = (paragraph: string, modified: string) =>
      `<?xml version="1.0"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/docProps/core.xml" pkg:contentType="application/xml"><pkg:xmlData><cp:coreProperties xmlns:cp="urn:core"><dcterms:modified xmlns:dcterms="urn:terms">${modified}</dcterms:modified></cp:coreProperties></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${paragraph}<w:sectPr w:rsidR="12345678" w:rsidSect="23456789"/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`
    let current = flatOpc(
      '<w:p w:rsidR="00112233" w:rsidRDefault="33445566" w14:paraId="11111111" w14:textId="22222222"><w:r><w:t>Hello</w:t></w:r><w:proofErr w:type="spellStart"></w:proofErr><w:lastRenderedPageBreak></w:lastRenderedPageBreak></w:p>',
      '2026-08-22T10:00:00Z',
    )
    const body = { getOoxml: vi.fn(() => ({ value: current })) }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'Word',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ document: { body }, sync: vi.fn().mockResolvedValue(undefined) }),
      },
    })
    const subject = new BrowserWordAdapter()
    const before = await subject.fingerprint()

    current = flatOpc(
      '<w:p w14:textId="BBBBBBBB" w14:paraId="AAAAAAAA" w:rsidRDefault="88776655" w:rsidR="99887766"><w:r><w:t>Hello</w:t></w:r></w:p>',
      '2026-08-22T10:01:00Z',
    )
    const metadataOnly = await subject.fingerprint()
    expect(metadataOnly).toBe(before)

    current = flatOpc(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r></w:p>',
      '2026-08-22T10:02:00Z',
    )
    await expect(subject.fingerprint()).resolves.not.toBe(before)
    current = flatOpc('<w:p><w:r><w:t>Hello changed</w:t></w:r></w:p>', '2026-08-22T10:03:00Z')
    await expect(subject.fingerprint()).resolves.not.toBe(before)
  })

  it('marks bounded paragraph pages as incomplete and rejects out-of-range starts', async () => {
    const makeParagraph = (index: number) => ({ text: `p${index}`, load: vi.fn() })
    const paragraphs = {
      items: Array.from({ length: 501 }, (_, index) => makeParagraph(index)),
      load: vi.fn(),
    }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'Word',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ document: { body: { paragraphs } }, sync: vi.fn() }),
      },
    })
    await expect(
      new BrowserWordAdapter().getDocumentText({ includeFormatting: false }),
    ).resolves.toMatchObject({
      totalParagraphs: 500,
      totalParagraphsExact: false,
      hasMore: true,
      showing: { start: 0, end: 500 },
      paragraphs: { length: 500 },
    })

    paragraphs.items = []
    await expect(
      new BrowserWordAdapter().getDocumentText({ startParagraph: 999 }, undefined),
    ).rejects.toThrow('invalid_tool_input')
    expect(paragraphs.load).toHaveBeenLastCalledWith({ $skip: 998, $top: 1 })
  })

  it('preflights every declarative operation before queuing any Word mutation', async () => {
    const insertText = vi.fn()
    const replacement = vi.fn()
    const sync = vi.fn().mockResolvedValue(undefined)
    const oversizedResults = {
      items: Array.from({ length: 1_001 }, () => ({ insertText: replacement })),
      load: vi.fn(),
    }
    const body = {
      insertText,
      getOoxml: vi.fn(() => ({
        value:
          '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p></w:body></w:document>',
      })),
      search: vi.fn(() => oversizedResults),
    }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'Word',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: {
        run: (callback: (context: unknown) => unknown) => callback({ document: { body }, sync }),
      },
    })

    await expect(
      new BrowserWordAdapter().executeOperations([
        { op: 'insert_text', location: 'end', text: 'queued-too-early' },
        { op: 'replace_all', search: 'before', replacement: 'after', matchCase: true },
      ]),
    ).rejects.toThrow('office_write_failed')
    expect(sync).toHaveBeenCalledOnce()
    expect(insertText).not.toHaveBeenCalled()
    expect(replacement).not.toHaveBeenCalled()
  })

  it('verifies the exact full bounded Word text after a declarative write', async () => {
    const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
    let current = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${Array.from(
      { length: 501 },
      (_, index) => paragraph(`p${index}`),
    ).join('')}</w:body></w:document>`
    let pending = ''
    const body = {
      insertText: vi.fn((value: string) => {
        pending += value
      }),
      getOoxml: vi.fn(() => ({ value: current })),
      search: vi.fn(() => ({ items: [], load: vi.fn() })),
    }
    const sync = vi.fn(async () => {
      if (!pending) return
      current = current.replace(
        '</w:t></w:r></w:p></w:body>',
        `${pending}</w:t></w:r></w:p></w:body>`,
      )
      pending = ''
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'Word',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: {
        run: (callback: (context: unknown) => unknown) => callback({ document: { body }, sync }),
      },
    })
    const subject = new BrowserWordAdapter()
    const operation = { op: 'insert_text', location: 'end', text: 'done' } as const

    await subject.executeOperations([operation])
    await expect(subject.verifyOperations([operation])).resolves.toBe(true)

    await subject.executeOperations([operation])
    current = current.replace('<w:t>p0</w:t>', '<w:t>changed</w:t>')
    await expect(subject.verifyOperations([operation])).resolves.toBe(false)
  })

  it('writes native Word paragraphs and tables and verifies their exact post-write snapshot', async () => {
    let current =
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:body></w:document>'
    let pending: string | undefined
    let queuedSnapshots: Array<{ value: string }> = []
    const paragraph = {
      styleBuiltIn: '',
      insertText: vi.fn(() => ({ font: {} })),
    }
    const insertParagraph = vi.fn(() => paragraph)
    const insertTable = vi.fn(() => ({ styleBuiltIn: '', headerRowCount: 0 }))
    const body = {
      clear: vi.fn(),
      paragraphs: { getFirst: vi.fn(() => paragraph) },
      insertParagraph,
      insertTable,
      insertOoxml: vi.fn((xml: string) => {
        pending = xml
      }),
      getOoxml: vi.fn(() => {
        const result = { value: current }
        queuedSnapshots.push(result)
        return result
      }),
    }
    const sync = vi.fn(async () => {
      if (pending) current = pending
      pending = undefined
      for (const snapshot of queuedSnapshots) snapshot.value = current
      queuedSnapshots = []
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'Word',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: {
        run: (callback: (context: unknown) => unknown) => callback({ document: { body }, sync }),
      },
    })
    const subject = new BrowserWordAdapter()
    const write = {
      mode: 'replace' as const,
      blocks: [
        { type: 'heading' as const, level: 1, spans: [{ text: 'Title' }] },
        { type: 'table' as const, rows: [['A']], headerRows: 1 as const },
      ],
      semanticText: 'Title\nA',
      structure: { headings: 1, lists: 0, tables: 1 },
    }

    await subject.executeDocumentWrite(write)

    expect(body.insertOoxml).toHaveBeenCalledOnce()
    expect(body.clear).not.toHaveBeenCalled()
    expect(insertParagraph).not.toHaveBeenCalled()
    expect(insertTable).not.toHaveBeenCalled()
    await expect(subject.verifyDocumentWrite(write)).resolves.toBe(true)
  })

  it('accepts semantically identical Word OOXML normalized after the write sync', async () => {
    let current =
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:body></w:document>'
    let pending: string | undefined
    let queuedSnapshots: Array<{ value: string }> = []
    const body = {
      insertOoxml: vi.fn((xml: string) => {
        pending = xml
      }),
      getOoxml: vi.fn(() => {
        const result = { value: current }
        queuedSnapshots.push(result)
        return result
      }),
    }
    const sync = vi.fn(async () => {
      if (pending) current = pending
      pending = undefined
      for (const snapshot of queuedSnapshots) snapshot.value = current
      queuedSnapshots = []
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'Word',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: {
        run: (callback: (context: unknown) => unknown) => callback({ document: { body }, sync }),
      },
    })
    const subject = new BrowserWordAdapter()
    const write = {
      mode: 'replace' as const,
      blocks: [
        { type: 'heading' as const, level: 1, spans: [{ text: 'Title' }] },
        { type: 'table' as const, rows: [['A']], headerRows: 1 as const },
      ],
      semanticText: 'Title\nA',
      structure: { headings: 1, lists: 0, tables: 1 },
    }

    await subject.executeDocumentWrite(write)
    current = current.replace('<w:tblPr>', '<w:tblPr><w:tblW w:w="0" w:type="auto"/>')

    await expect(subject.verifyDocumentWrite(write)).resolves.toBe(true)
  })

  it('accepts multiple transient empty paragraphs Word appends after a final table', async () => {
    let current =
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:body></w:document>'
    let pending: string | undefined
    let queuedSnapshots: Array<{ value: string }> = []
    const body = {
      insertOoxml: vi.fn((xml: string) => {
        pending = xml
      }),
      getOoxml: vi.fn(() => {
        const result = { value: current }
        queuedSnapshots.push(result)
        return result
      }),
    }
    const sync = vi.fn(async () => {
      if (pending)
        current = pending.replace(
          '</w:tbl>',
          '</w:tbl><w:p/><w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>',
        )
      pending = undefined
      for (const snapshot of queuedSnapshots) snapshot.value = current
      queuedSnapshots = []
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'Word',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: {
        run: (callback: (context: unknown) => unknown) => callback({ document: { body }, sync }),
      },
    })
    const subject = new BrowserWordAdapter()
    const write = {
      mode: 'replace' as const,
      blocks: [
        { type: 'paragraph' as const, spans: [{ text: 'Text' }] },
        { type: 'table' as const, rows: [['A']], headerRows: 1 as const },
      ],
      semanticText: 'Text\nA',
      structure: { headings: 0, lists: 0, tables: 1 },
    }

    await subject.executeDocumentWrite(write)
    await expect(subject.verifyDocumentWrite(write)).resolves.toBe(true)
  })

  it('accepts localized built-in heading style IDs normalized during the write sync', async () => {
    let current =
      '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Old</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="1"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/header1.xml"><pkg:xmlData><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr></pkg:xmlData></pkg:part></pkg:package>'
    let pending: string | undefined
    let queuedSnapshots: Array<{ value: string }> = []
    const body = {
      insertOoxml: vi.fn((xml: string) => {
        pending = xml
      }),
      getOoxml: vi.fn(() => {
        const result = { value: current }
        queuedSnapshots.push(result)
        return result
      }),
    }
    const sync = vi.fn(async () => {
      if (pending)
        current = pending
          .replace('w:val="Heading1"', 'w:val="1"')
          .replace('<w:outlineLvl w:val="0"/>', '')
          .replace('</w:sectPr>', '<w:cols w:space="708"/></w:sectPr>')
      pending = undefined
      for (const snapshot of queuedSnapshots) snapshot.value = current
      queuedSnapshots = []
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'Word',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: {
        run: (callback: (context: unknown) => unknown) => callback({ document: { body }, sync }),
      },
    })
    const subject = new BrowserWordAdapter()
    const write = {
      mode: 'replace' as const,
      blocks: [{ type: 'heading' as const, level: 1, spans: [{ text: 'Title' }] }],
      semanticText: 'Title',
      structure: { headings: 1, lists: 0, tables: 0 },
    }

    await subject.executeDocumentWrite(write)

    expect(body.insertOoxml).toHaveBeenCalledWith(expect.stringContaining('w:val="1"'), 'Replace')
    await expect(subject.verifyDocumentWrite(write)).resolves.toBe(true)
  })

  it('reads the post-write Word snapshot in a separate sync batch', async () => {
    const original =
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:body></w:document>'
    let current = original
    let pending: string | undefined
    const body = {
      insertOoxml: vi.fn((xml: string) => {
        pending = xml
      }),
      getOoxml: vi.fn(() => ({ value: current })),
    }
    const sync = vi.fn(async () => {
      if (pending) current = pending
      pending = undefined
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'Word',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: {
        run: (callback: (context: unknown) => unknown) => callback({ document: { body }, sync }),
      },
    })
    const subject = new BrowserWordAdapter()
    const write = {
      mode: 'replace' as const,
      blocks: [{ type: 'paragraph' as const, spans: [{ text: 'New' }] }],
      semanticText: 'New',
      structure: { headings: 0, lists: 0, tables: 0 },
    }

    await subject.executeDocumentWrite(write)

    await expect(subject.verifyDocumentWrite(write)).resolves.toBe(true)
  })

  it('reports a content-free verification stage when Word changes inserted text', async () => {
    let current =
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:body></w:document>'
    let pending: string | undefined
    let queuedSnapshots: Array<{ value: string }> = []
    const body = {
      insertOoxml: vi.fn((xml: string) => {
        pending = xml
      }),
      getOoxml: vi.fn(() => {
        const result = { value: current }
        queuedSnapshots.push(result)
        return result
      }),
    }
    const sync = vi.fn(async () => {
      if (pending) current = pending.replace('>New<', '>Changed<')
      pending = undefined
      for (const snapshot of queuedSnapshots) snapshot.value = current
      queuedSnapshots = []
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'Word',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: {
        run: (callback: (context: unknown) => unknown) => callback({ document: { body }, sync }),
      },
    })
    const subject = new BrowserWordAdapter()
    const write = {
      mode: 'replace' as const,
      blocks: [{ type: 'paragraph' as const, spans: [{ text: 'New' }] }],
      semanticText: 'New',
      structure: { headings: 0, lists: 0, tables: 0 },
    }

    await expect(subject.executeDocumentWrite(write)).rejects.toThrow(
      'office_recovery_failed:word_text',
    )
  })

  it('restores the original Word body when cancellation arrives during the write sync', async () => {
    const original =
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Original</w:t></w:r></w:p></w:body></w:document>'
    let current = original
    let syncCount = 0
    let queuedSnapshots: Array<{ value: string }> = []
    const controller = new AbortController()
    const insertOoxml = vi.fn((xml: string) => {
      current = xml
    })
    const body = {
      clear: vi.fn(),
      paragraphs: {
        getFirst: vi.fn(() => ({
          insertText: vi.fn(() => ({ font: {} })),
        })),
      },
      insertParagraph: vi.fn(() => ({ insertText: vi.fn(() => ({ font: {} })) })),
      insertTable: vi.fn(),
      insertOoxml,
      getOoxml: vi.fn(() => {
        const result = { value: current }
        queuedSnapshots.push(result)
        return result
      }),
    }
    const sync = vi.fn(async () => {
      syncCount += 1
      if (syncCount === 2) {
        current = current.replace('Original', 'Changed')
        controller.abort()
      }
      for (const snapshot of queuedSnapshots) snapshot.value = current
      queuedSnapshots = []
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'Word',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      Word: {
        run: (callback: (context: unknown) => unknown) => callback({ document: { body }, sync }),
      },
    })

    await expect(
      new BrowserWordAdapter().executeDocumentWrite(
        {
          mode: 'replace',
          blocks: [{ type: 'paragraph', spans: [{ text: 'Changed' }] }],
          semanticText: 'Changed',
          structure: { headings: 0, lists: 0, tables: 0 },
        },
        controller.signal,
      ),
    ).rejects.toThrow('cancelled')
    expect(insertOoxml).toHaveBeenCalledWith(original, 'Replace')
    expect(current).toBe(original)
  })
})
