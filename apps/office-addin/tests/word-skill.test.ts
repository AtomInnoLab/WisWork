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
})
