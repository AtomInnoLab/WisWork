const MAX_PDF_BYTES = 16 * 1024 * 1024
const MAX_PAGE_PIXELS = 4_000_000

function cancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new Error('cancelled')
}

export async function renderPdfPageToPng(
  source: Uint8Array,
  pageNumber: number,
  signal?: AbortSignal,
): Promise<string> {
  cancelled(signal)
  if (source.byteLength < 5 || source.byteLength > MAX_PDF_BYTES || pageNumber < 1)
    throw new Error('office_read_failed')
  GlobalWorkerOptions.workerSrc = workerUrl
  cancelled(signal)
  const loading = getDocument({
    data: source.slice(),
    useWorkerFetch: false,
    useWasm: false,
    maxImageSize: MAX_PAGE_PIXELS,
    stopAtErrors: true,
  })
  let abortDestroy: Promise<void> | undefined
  const abortLoading = () => {
    abortDestroy ??= loading.destroy()
  }
  signal?.addEventListener('abort', abortLoading, { once: true })
  try {
    const pdf = await loading.promise
    if (pageNumber > pdf.numPages) throw new Error('invalid_tool_input')
    const page = await pdf.getPage(pageNumber)
    const original = page.getViewport({ scale: 1 })
    if (!Number.isFinite(original.width) || !Number.isFinite(original.height))
      throw new Error('office_read_failed')
    const scale = Math.min(2, Math.sqrt(MAX_PAGE_PIXELS / (original.width * original.height)))
    const viewport = page.getViewport({ scale })
    const canvasWidth = Math.ceil(viewport.width)
    const canvasHeight = Math.ceil(viewport.height)
    if (canvasWidth < 1 || canvasHeight < 1 || canvasWidth * canvasHeight > MAX_PAGE_PIXELS)
      throw new Error('office_read_failed')
    const canvas = document.createElement('canvas')
    canvas.width = canvasWidth
    canvas.height = canvasHeight
    const context = canvas.getContext('2d', { alpha: false })
    if (!context) throw new Error('office_read_failed')
    const render = page.render({ canvas, canvasContext: context, viewport })
    const abort = () => render.cancel()
    signal?.addEventListener('abort', abort, { once: true })
    try {
      await render.promise
    } finally {
      signal?.removeEventListener('abort', abort)
    }
    cancelled(signal)
    const url = canvas.toDataURL('image/png')
    if (!url.startsWith('data:image/png;base64,')) throw new Error('office_read_failed')
    return url.slice('data:image/png;base64,'.length)
  } catch (error) {
    if (signal?.aborted) throw new Error('cancelled', { cause: error })
    if (
      error instanceof Error &&
      ['invalid_tool_input', 'office_read_failed'].includes(error.message)
    )
      throw error
    throw new Error('office_read_failed', { cause: error })
  } finally {
    signal?.removeEventListener('abort', abortLoading)
    await (abortDestroy ?? loading.destroy())
  }
}
import { getDocument, GlobalWorkerOptions } from 'pdfjs-dist/legacy/build/pdf.mjs'
import workerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
