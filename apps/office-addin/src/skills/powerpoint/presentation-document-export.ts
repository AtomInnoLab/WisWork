const SLICE_BYTES = 64 * 1024
const CALLBACK_TIMEOUT_MS = 10_000
type Format = 'pptx' | 'pdf'

function officeApi(): typeof Office | undefined {
  return typeof Office === 'undefined' ? undefined : Office
}

export function supportsPowerPointDocumentExport(format: Format = 'pptx'): boolean {
  try {
    const api = officeApi()
    return (
      !!api &&
      String(api.context?.host) === 'PowerPoint' &&
      typeof api.context.document?.getFileAsync === 'function' &&
      (format === 'pptx' || format === 'pdf') &&
      api.FileType?.[format === 'pptx' ? 'Compressed' : 'Pdf'] !== undefined &&
      api.context.requirements.isSetSupported(
        format === 'pptx' ? 'CompressedFile' : 'PdfFile',
        '1.1',
      )
    )
  } catch {
    return false
  }
}

function waitFor<T>(
  invoke: (callback: (result: Office.AsyncResult<T>) => void) => void,
  signal?: AbortSignal,
  late?: (value: T) => void,
): Promise<T> {
  return new Promise((resolve, reject) => {
    let settled = false
    let delivered: T | undefined
    const finish = (error?: Error, value?: T) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      signal?.removeEventListener('abort', abort)
      if (error) reject(error)
      else {
        delivered = value
        resolve(value as T)
      }
    }
    const abort = () => finish(new Error('office_document_export_cancelled'))
    const timer = setTimeout(
      () => finish(new Error('office_document_export_timeout')),
      CALLBACK_TIMEOUT_MS,
    )
    signal?.addEventListener('abort', abort, { once: true })
    if (signal?.aborted) {
      abort()
      return
    }
    try {
      invoke((result) => {
        try {
          if (
            !result ||
            typeof result !== 'object' ||
            ((result.status as unknown) !== 'succeeded' && (result.status as unknown) !== 'failed')
          ) {
            finish(new Error('office_document_export_invalid'))
            return
          }
          if (settled) {
            if (
              (result.status as unknown) === 'succeeded' &&
              result.value &&
              result.value !== delivered
            )
              late?.(result.value)
            return
          }
          if ((result.status as unknown) !== 'succeeded')
            finish(new Error('office_document_export_failed'))
          else finish(undefined, result.value)
        } catch {
          finish(new Error('office_document_export_invalid'))
        }
      })
    } catch {
      finish(new Error('office_document_export_failed'))
    }
  })
}

export async function exportPowerPointDocument(
  format: Format,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  if (signal?.aborted) throw new Error('office_document_export_cancelled')
  if (!supportsPowerPointDocumentExport(format))
    throw new Error('office_document_export_unavailable')
  const api = officeApi()!
  const closed = new WeakSet<object>()
  const close = async (file: Office.File) => {
    if (!file || typeof file !== 'object' || closed.has(file)) return
    closed.add(file)
    try {
      await waitFor<void>((callback) => file.closeAsync(callback))
    } catch {
      /* Close remains bounded even if the host never responds. */
    }
  }
  const file = await waitFor<Office.File>(
    (callback) =>
      api.context.document.getFileAsync(
        format === 'pptx' ? api.FileType.Compressed : api.FileType.Pdf,
        { sliceSize: SLICE_BYTES },
        callback,
      ),
    signal,
    (lateFile) => {
      void close(lateFile)
    },
  )
  try {
    const limit = (format === 'pptx' ? 20 : 10) * 1024 * 1024
    if (
      !file ||
      !Number.isSafeInteger(file.size) ||
      file.size < 1 ||
      file.size > limit ||
      !Number.isSafeInteger(file.sliceCount) ||
      file.sliceCount !== Math.ceil(file.size / SLICE_BYTES) ||
      typeof file.getSliceAsync !== 'function' ||
      typeof file.closeAsync !== 'function'
    )
      throw new Error('office_document_export_invalid')
    const bytes = new Uint8Array(file.size)
    let offset = 0
    for (let index = 0; index < file.sliceCount; index++) {
      const slice = await waitFor<Office.Slice>(
        (callback) => file.getSliceAsync(index, callback),
        signal,
      )
      if (signal?.aborted) throw new Error('office_document_export_cancelled')
      const expected = Math.min(SLICE_BYTES, file.size - offset)
      if (
        !slice ||
        slice.index !== index ||
        slice.size !== expected ||
        !Array.isArray(slice.data) ||
        slice.data.length !== expected
      )
        throw new Error('office_document_export_invalid')
      for (let i = 0; i < expected; i++) {
        const value: unknown = slice.data[i]
        if (typeof value !== 'number' || !Number.isInteger(value) || value < 0 || value > 255)
          throw new Error('office_document_export_invalid')
        bytes[offset + i] = value
      }
      offset += expected
    }
    const signature = format === 'pptx' ? [80, 75, 3, 4] : [37, 80, 68, 70, 45]
    if (offset !== file.size || signature.some((value, index) => bytes[index] !== value))
      throw new Error('office_document_export_invalid')
    return bytes
  } catch (error) {
    if (
      error instanceof Error &&
      /^office_document_export_(invalid|failed|timeout|cancelled)$/.test(error.message)
    )
      throw error
    throw new Error('office_document_export_failed', { cause: error })
  } finally {
    await close(file)
  }
}
