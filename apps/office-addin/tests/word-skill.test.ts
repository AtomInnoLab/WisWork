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

  it('keeps raw Office.js visible but disabled without an audited hardened evaluator', async () => {
    const proposals = createStructuredProposalController()
    const skill = createWordSkill({ adapter: adapter(), vfs: new InMemoryVfs(), proposals })
    await expect(
      skill.executeTool(
        call('execute_office_js', { code: 'context.document.body.insertText("x", "End")' }),
      ),
    ).resolves.toMatchObject({ output: 'office_api_unsupported', isError: true, mutated: false })
    expect(proposals.pending()).toBeUndefined()
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
})
