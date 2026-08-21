import JSZip from 'jszip'
import { describe, expect, it, vi } from 'vitest'
import { convertDocument, type ConversionLimits } from '../src/skills/shared/conversion-engine.js'
import {
  ConversionWorkerRuntime,
  type ConversionWorkerLike,
} from '../src/skills/shared/conversion-runtime.js'
import { createSharedBrowserSkill } from '../src/skills/shared/shared-skill.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'

const limits: ConversionLimits = {
  maxInputBytes: 256 * 1024,
  maxEntries: 32,
  maxEntryBytes: 64 * 1024,
  maxArchiveBytes: 128 * 1024,
  maxCompressionRatio: 20,
  maxPages: 4,
  maxSheets: 3,
  maxRows: 20,
  maxColumns: 10,
  maxCells: 40,
  maxPagePixels: 1_000_000,
  maxTotalPixels: 2_000_000,
  maxOutputBytes: 128 * 1024,
}

async function zip(entries: Record<string, string>): Promise<Uint8Array> {
  const archive = new JSZip()
  for (const [path, value] of Object.entries(entries)) archive.file(path, value)
  return archive.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

async function docx(text = 'Hello world'): Promise<Uint8Array> {
  return zip({
    '[Content_Types].xml': '<Types/>',
    'word/document.xml': `<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`,
  })
}

async function xlsx(cell = 'B2'): Promise<Uint8Array> {
  return zip({
    '[Content_Types].xml': '<Types/>',
    'xl/workbook.xml':
      '<workbook xmlns:r="r"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels':
      '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
    'xl/sharedStrings.xml': '<sst><si><t>Hello</t></si></sst>',
    'xl/worksheets/sheet1.xml': `<worksheet><sheetData><row r="2"><c r="${cell}" t="s"><v>0</v></c></row></sheetData></worksheet>`,
  })
}

describe('bounded conversion engine', () => {
  it('converts DOCX text and XLSX sheets to deterministic outputs', async () => {
    await expect(
      convertDocument(
        { kind: 'docx-to-text', inputName: 'paper.docx', bytes: await docx() },
        limits,
      ),
    ).resolves.toEqual([{ path: 'paper.txt', bytes: new TextEncoder().encode('Hello world') }])

    const converted = await convertDocument(
      { kind: 'xlsx-to-csv', inputName: 'book.xlsx', bytes: await xlsx() },
      limits,
    )
    expect(converted).toHaveLength(1)
    expect(converted[0].path).toBe('book-Data.csv')
    expect(new TextDecoder().decode(converted[0].bytes)).toBe(',\n,Hello')
  })

  it('extracts text from a bounded PDF fixture', async () => {
    const bytes = new TextEncoder().encode(
      '%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 200 200]/Resources<</Font<</F1 4 0 R>>>>/Contents 5 0 R>>endobj\n4 0 obj<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>endobj\n5 0 obj<</Length 38>>stream\nBT /F1 12 Tf 20 100 Td (PDF hello) Tj ET\nendstream endobj\nxref\n0 6\n0000000000 65535 f \ntrailer<</Root 1 0 R/Size 6>>\nstartxref\n0\n%%EOF',
    )
    const output = await convertDocument(
      { kind: 'pdf-to-text', inputName: 'note.pdf', bytes },
      limits,
    )
    expect(new TextDecoder().decode(output[0].bytes)).toContain('PDF hello')
  })

  it.each([
    ['archive traversal', () => zip({ '../escape': 'x', 'word/document.xml': '<w:t>x</w:t>' })],
    ['malformed XML', () => docx('<broken')],
    ['extreme A1 coordinate', () => xlsx('XFD1048576')],
  ])('rejects %s without outputs', async (_name, fixture) => {
    await expect(
      convertDocument(
        {
          kind: _name === 'extreme A1 coordinate' ? 'xlsx-to-csv' : 'docx-to-text',
          inputName: 'bad.zip',
          bytes: await fixture(),
        },
        limits,
      ),
    ).rejects.toThrow(/conversion_(?:archive_unsafe|invalid_document|limit)/)
  })

  it('rejects ZIP bomb metadata before decompression', async () => {
    const bomb = await zip({ 'word/document.xml': 'x'.repeat(60_000) })
    await expect(
      convertDocument(
        { kind: 'docx-to-text', inputName: 'bomb.docx', bytes: bomb },
        { ...limits, maxCompressionRatio: 2 },
      ),
    ).rejects.toThrow('conversion_archive_unsafe')
  })

  it('enforces archive entry and conversion output byte limits', async () => {
    const tooMany = await zip({
      'word/document.xml': '<w:document><w:t>x</w:t></w:document>',
      a: '1',
      b: '2',
    })
    await expect(
      convertDocument(
        { kind: 'docx-to-text', inputName: 'many.docx', bytes: tooMany },
        { ...limits, maxEntries: 2 },
      ),
    ).rejects.toThrow('conversion_limit')
    await expect(
      convertDocument(
        { kind: 'docx-to-text', inputName: 'large.docx', bytes: await docx('long output') },
        { ...limits, maxOutputBytes: 4 },
      ),
    ).rejects.toThrow('conversion_limit')
  })

  it('enforces PDF page count before reading pages', async () => {
    const getPage = vi.fn()
    await expect(
      convertDocument(
        { kind: 'pdf-to-text', inputName: 'long.pdf', bytes: Uint8Array.of(1) },
        limits,
        { loadPdf: async () => ({ numPages: 5, getPage, destroy: vi.fn() }) },
      ),
    ).rejects.toThrow('conversion_limit')
    expect(getPage).not.toHaveBeenCalled()
  })

  it('rejects oversized PDF page pixels before rendering', async () => {
    const pdf = {
      numPages: 1,
      getPage: async () => ({
        getViewport: () => ({ width: 2_000, height: 2_000 }),
        render: vi.fn(),
      }),
      destroy: vi.fn(),
    }
    await expect(
      convertDocument(
        { kind: 'pdf-to-images', inputName: 'huge.pdf', bytes: Uint8Array.of(1) },
        limits,
        { loadPdf: async () => pdf, renderPdfPage: vi.fn() },
      ),
    ).rejects.toThrow('conversion_limit')
    expect(pdf.getPage).toBeDefined()
  })
})

class FakeWorker implements ConversionWorkerLike {
  onmessage: ((event: MessageEvent) => void) | null = null
  onerror: ((event: ErrorEvent) => void) | null = null
  terminated = false
  constructor(private readonly respond: boolean) {}
  postMessage(message: { id: string }): void {
    if (this.respond)
      queueMicrotask(() =>
        this.onmessage?.(
          new MessageEvent('message', {
            data: {
              id: message.id,
              ok: true,
              outputs: [{ path: 'result.txt', bytes: Uint8Array.of(111, 107) }],
            },
          }),
        ),
      )
  }
  terminate(): void {
    this.terminated = true
  }
}

describe('conversion worker runtime', () => {
  it('atomically mounts worker outputs only after a successful response', async () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/input.docx', Uint8Array.of(1))
    const runtime = new ConversionWorkerRuntime(vfs, { workerFactory: () => new FakeWorker(true) })
    await expect(runtime.run('docx-to-text', '/home/user/input.docx')).resolves.toEqual([
      '/home/user/result.txt',
    ])
    expect(vfs.readText('/home/user/result.txt')).toBe('ok')
  })

  it('rolls back every returned output when VFS quota rejects the batch', async () => {
    const vfs = new InMemoryVfs({ maxFiles: 2 })
    vfs.writeFile('/home/user/input.docx', Uint8Array.of(1))
    const worker = new FakeWorker(true)
    worker.postMessage = function (message: { id: string }): void {
      queueMicrotask(() =>
        this.onmessage?.(
          new MessageEvent('message', {
            data: {
              id: message.id,
              ok: true,
              outputs: [
                { path: 'one.txt', bytes: Uint8Array.of(1) },
                { path: 'two.txt', bytes: Uint8Array.of(2) },
              ],
            },
          }),
        ),
      )
    }
    const runtime = new ConversionWorkerRuntime(vfs, { workerFactory: () => worker })
    await expect(runtime.run('docx-to-text', '/home/user/input.docx')).rejects.toThrow('vfs_limit')
    expect(vfs.list('/home/user')).toEqual(['/home/user/input.docx'])
  })

  it('registers all real conversion commands in the shared Office skill', async () => {
    const run = vi.fn().mockResolvedValue(['/home/user/out.txt'])
    const skill = createSharedBrowserSkill({
      vfs: new InMemoryVfs(),
      conversionRuntime: { run },
    })
    const bash = skill.tools.find((tool) => tool.name === 'bash')
    expect(bash?.description).toContain('pdf-to-text')
    expect(bash?.description).toContain('pdf-to-images')
    expect(bash?.description).toContain('docx-to-text')
    expect(bash?.description).toContain('xlsx-to-csv')
    await expect(
      skill.executeTool({
        id: 'convert',
        name: 'bash',
        input: { command: 'docx-to-text /home/user/paper.docx' },
      }),
    ).resolves.toMatchObject({ output: '/home/user/out.txt', isError: false })
    expect(run).toHaveBeenCalledWith(
      'docx-to-text',
      '/home/user/paper.docx',
      expect.any(AbortSignal),
    )
  })

  it('terminates a hanging worker on timeout and mounts no partial output', async () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/input.docx', Uint8Array.of(1))
    const worker = new FakeWorker(false)
    const runtime = new ConversionWorkerRuntime(vfs, {
      timeoutMs: 5,
      workerFactory: () => worker,
    })
    await expect(runtime.run('docx-to-text', '/home/user/input.docx')).rejects.toThrow(
      'command_timeout',
    )
    expect(worker.terminated).toBe(true)
    expect(vfs.list('/home/user')).toEqual(['/home/user/input.docx'])
  })

  it('terminates on cancellation and leaves the VFS unchanged', async () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/input.docx', Uint8Array.of(1))
    const worker = new FakeWorker(false)
    const runtime = new ConversionWorkerRuntime(vfs, { workerFactory: () => worker })
    const controller = new AbortController()
    const pending = runtime.run('docx-to-text', '/home/user/input.docx', controller.signal)
    controller.abort()
    await expect(pending).rejects.toThrow('cancelled')
    expect(worker.terminated).toBe(true)
    expect(vfs.list('/home/user')).toEqual(['/home/user/input.docx'])
  })
})
