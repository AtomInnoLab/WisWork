import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  exportPowerPointDocument,
  supportsPowerPointDocumentExport,
} from '../src/skills/powerpoint/presentation-document-export.js'

type Callback = (result: unknown) => void
const success = (value?: unknown) => ({ status: 'succeeded', value })
function fixture(bytes = [80, 75, 3, 4, 1]) {
  const file = {
    size: bytes.length,
    sliceCount: 1,
    getSliceAsync: vi.fn(function (this: unknown, index: number, callback: Callback) {
      expect(this).toBe(file)
      callback(success({ index, size: bytes.length, data: bytes }))
    }),
    closeAsync: vi.fn(function (this: unknown, callback: Callback) {
      expect(this).toBe(file)
      callback(success())
    }),
  }
  const document = {
    getFileAsync: vi.fn(function (
      this: unknown,
      _type: unknown,
      _options: unknown,
      callback: Callback,
    ) {
      expect(this).toBe(document)
      callback(success(file))
    }),
  }
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', document, requirements: { isSetSupported: () => true } },
    FileType: { Compressed: 'compressed', Pdf: 'pdf' },
    HostType: { Word: 'Word' },
  })
  return { file, document }
}
afterEach(() => {
  vi.unstubAllGlobals()
  vi.useRealTimers()
})

describe('current PowerPoint document export', () => {
  it('reads real compressed document bytes and closes the file', async () => {
    const { file, document } = fixture()
    expect(supportsPowerPointDocumentExport()).toBe(true)
    expect(await exportPowerPointDocument('pptx')).toEqual(Uint8Array.from([80, 75, 3, 4, 1]))
    expect(document.getFileAsync).toHaveBeenCalledWith(
      'compressed',
      { sliceSize: 65536 },
      expect.any(Function),
    )
    expect(file.closeAsync).toHaveBeenCalledOnce()
  })
  it('exports actual PDF bytes', async () => {
    const { document } = fixture([37, 80, 68, 70, 45, 49])
    await exportPowerPointDocument('pdf')
    expect(document.getFileAsync.mock.calls[0]?.[0]).toBe('pdf')
  })
  it('joins full slices and a shorter final slice', async () => {
    const bytes = new Array<number>(65539).fill(7)
    bytes.splice(0, 4, 80, 75, 3, 4)
    const { file } = fixture(bytes)
    file.sliceCount = 2
    file.getSliceAsync.mockImplementation((index, callback) => {
      const data = bytes.slice(index * 65536, (index + 1) * 65536)
      callback(success({ index, size: data.length, data }))
    })
    expect(await exportPowerPointDocument('pptx')).toEqual(Uint8Array.from(bytes))
    expect(file.getSliceAsync).toHaveBeenCalledTimes(2)
    expect(file.closeAsync).toHaveBeenCalledOnce()
  })
  it.each(['pptx', 'pdf'] as const)(
    'rejects oversized %s before reading slices',
    async (format) => {
      const { file } = fixture()
      file.size = (format === 'pptx' ? 20 : 10) * 1024 * 1024 + 1
      await expect(exportPowerPointDocument(format)).rejects.toThrow()
      expect(file.getSliceAsync).not.toHaveBeenCalled()
      expect(file.closeAsync).toHaveBeenCalledOnce()
    },
  )
  it('reports missing requirement support', () => {
    fixture()
    vi.spyOn(Office.context.requirements, 'isSetSupported').mockReturnValue(false)
    expect(supportsPowerPointDocumentExport('pdf')).toBe(false)
  })
  it('rejects other hosts and unsupported requirement sets', async () => {
    fixture()
    Office.context.host = Office.HostType.Word
    expect(supportsPowerPointDocumentExport()).toBe(false)
    await expect(exportPowerPointDocument('pptx')).rejects.toThrow(
      'office_document_export_unavailable',
    )
  })
  it.each(['size', 'sliceCount'])('rejects invalid file %s and closes', async (key) => {
    const { file } = fixture()
    Object.assign(file, { [key]: 0 })
    await expect(exportPowerPointDocument('pptx')).rejects.toThrow()
    expect(file.closeAsync).toHaveBeenCalledOnce()
  })
  it.each([
    { index: 1, size: 5, data: [80, 75, 3, 4, 1] },
    { index: 0, size: 4, data: [80, 75, 3, 4, 1] },
    { index: 0, size: 5, data: [80, 75, 3, 4, 256] },
    { index: 0, size: 5, data: [0, 0, 0, 0, 0] },
  ])('rejects malformed slice/signature %#', async (slice) => {
    const { file } = fixture()
    file.getSliceAsync.mockImplementation((_index, callback) => callback(success(slice)))
    await expect(exportPowerPointDocument('pptx')).rejects.toThrow()
    expect(file.closeAsync).toHaveBeenCalledOnce()
  })
  it('closes a file arriving after abort', async () => {
    const { file, document } = fixture()
    let deliver: Callback = () => {}
    document.getFileAsync.mockImplementation((_type, _options, callback) => {
      deliver = callback
    })
    const controller = new AbortController()
    const pending = exportPowerPointDocument('pptx', controller.signal)
    controller.abort()
    await expect(pending).rejects.toThrow('office_document_export_cancelled')
    deliver(success(file))
    expect(file.closeAsync).toHaveBeenCalledOnce()
  })
  it('closes acquired files when a slice callback fails or throws', async () => {
    for (const throws of [false, true]) {
      const { file } = fixture()
      file.getSliceAsync.mockImplementation((_index, callback) => {
        if (throws) throw new Error('host failure')
        callback({ status: 'failed' })
      })
      await expect(exportPowerPointDocument('pptx')).rejects.toThrow(
        'office_document_export_failed',
      )
      expect(file.closeAsync).toHaveBeenCalledOnce()
    }
  })
  it('bounds slice and close callbacks', async () => {
    vi.useFakeTimers()
    const { file } = fixture()
    file.getSliceAsync.mockImplementation(() => {})
    file.closeAsync.mockImplementation(() => {})
    const pending = expect(exportPowerPointDocument('pptx')).rejects.toThrow(
      'office_document_export_timeout',
    )
    await vi.advanceTimersByTimeAsync(30_000)
    await pending
    expect(file.closeAsync).toHaveBeenCalledOnce()
  })
  it('closes on abort while reading an acquired file', async () => {
    const { file } = fixture()
    const controller = new AbortController()
    file.getSliceAsync.mockImplementation(() => controller.abort())
    await expect(exportPowerPointDocument('pptx', controller.signal)).rejects.toThrow(
      'office_document_export_cancelled',
    )
    expect(file.closeAsync).toHaveBeenCalledOnce()
  })
  it('times out callbacks and closes a late acquired file', async () => {
    vi.useFakeTimers()
    const { file, document } = fixture()
    let deliver: Callback = () => {}
    document.getFileAsync.mockImplementation((_type, _options, callback) => {
      deliver = callback
    })
    const pending = expect(exportPowerPointDocument('pptx')).rejects.toThrow(
      'office_document_export_timeout',
    )
    await vi.advanceTimersByTimeAsync(30_000)
    await pending
    deliver(success(file))
    expect(file.closeAsync).toHaveBeenCalledOnce()
  })
  it.each([undefined, null, {}, { status: 123 }])(
    'rejects malformed async acquisition callbacks: %j',
    async (result) => {
      const { document } = fixture()
      let deliver: Callback = () => {}
      document.getFileAsync.mockImplementation((_type, _options, callback) => {
        deliver = callback
      })
      const pending = expect(exportPowerPointDocument('pptx')).rejects.toThrow(
        'office_document_export_invalid',
      )
      expect(() => deliver(result)).not.toThrow()
      await pending
    },
  )
  it('does not close duplicate acquired files while a slice is pending', async () => {
    const { file, document } = fixture()
    let deliverFile: Callback = () => {}
    let deliverSlice: Callback = () => {}
    document.getFileAsync.mockImplementation((_type, _options, callback) => {
      deliverFile = callback
    })
    file.getSliceAsync.mockImplementation((_index, callback) => {
      deliverSlice = callback
    })
    const pending = exportPowerPointDocument('pptx')
    deliverFile(success(file))
    await Promise.resolve()
    deliverFile(success(file))
    expect(file.closeAsync).not.toHaveBeenCalled()
    deliverSlice(success({ index: 0, size: 5, data: [80, 75, 3, 4, 1] }))
    await pending
    expect(file.closeAsync).toHaveBeenCalledOnce()
  })
  it('normalizes non-Error SDK faults and still closes the file', async () => {
    const { file } = fixture()
    Object.defineProperty(file, 'size', {
      get: () => {
        throw { sdk: 'broken' }
      },
    })
    await expect(exportPowerPointDocument('pptx')).rejects.toThrow('office_document_export_failed')
    expect(file.closeAsync).toHaveBeenCalledOnce()
  })
})
