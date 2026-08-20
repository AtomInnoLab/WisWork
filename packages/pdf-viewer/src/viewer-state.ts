import type {
  CancellableRenderTask,
  DestroyableDocument,
  ViewerDocumentState,
  ViewerState,
} from './types.js'

const MIN_SCALE = 0.5
const MAX_SCALE = 4
export const MAX_PDF_PAGES = 2_000
export const MAX_CANVAS_EDGE = 8_192
export const MAX_CANVAS_PIXELS = 8_000_000
export const MAX_VIEWPORT_CSS_EDGE = 16_384
export const MAX_VIEWPORT_CSS_PIXELS = 64_000_000

export function assertDocumentPageBudget(pageCount: number): void {
  if (!Number.isSafeInteger(pageCount) || pageCount < 1 || pageCount > MAX_PDF_PAGES) {
    throw new Error(`PDF page limit exceeded (${MAX_PDF_PAGES})`)
  }
}

export function getPageWindow(page: number, pageCount: number, radius = 1): number[] {
  const current = Math.min(pageCount, Math.max(1, Math.trunc(page)))
  const start = Math.max(1, current - radius)
  const end = Math.min(pageCount, current + radius)
  return Array.from({ length: Math.max(0, end - start + 1) }, (_, index) => start + index)
}

export function getDocumentPages(pageCount: number): number[] {
  if (!Number.isSafeInteger(pageCount) || pageCount < 1) return []
  return Array.from({ length: pageCount }, (_, index) => index + 1)
}

export function nearestPageToViewport(
  pages: readonly { page: number; top: number; bottom: number }[],
): number | null {
  let nearest: { page: number; distance: number } | null = null
  for (const candidate of pages) {
    if (!Number.isSafeInteger(candidate.page) || candidate.page < 1) continue
    const distance = Math.abs(candidate.top)
    if (!nearest || distance < nearest.distance) nearest = { page: candidate.page, distance }
  }
  return nearest?.page ?? null
}

export function computeCanvasBudget(width: number, height: number, desiredScale: number) {
  if (
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    !Number.isFinite(desiredScale) ||
    width <= 0 ||
    height <= 0 ||
    desiredScale <= 0 ||
    width > MAX_VIEWPORT_CSS_EDGE ||
    height > MAX_VIEWPORT_CSS_EDGE ||
    width * height > MAX_VIEWPORT_CSS_PIXELS
  )
    throw new Error('PDF viewport exceeds the rendering budget')
  const allocationScale = Math.min(
    desiredScale,
    MAX_CANVAS_EDGE / width,
    MAX_CANVAS_EDGE / height,
    Math.sqrt(MAX_CANVAS_PIXELS / (width * height)),
  )
  const result = {
    width: Math.max(1, Math.floor(width * allocationScale)),
    height: Math.max(1, Math.floor(height * allocationScale)),
    allocationScale,
  }
  if (
    result.width > MAX_CANVAS_EDGE ||
    result.height > MAX_CANVAS_EDGE ||
    result.width * result.height > MAX_CANVAS_PIXELS
  ) {
    throw new Error('PDF canvas exceeds the rendering budget')
  }
  return result
}

export function resetViewerPageForSource(_currentPage: number): number {
  return 1
}

export function clampViewerPageToDocument(page: number, pageCount: number): number {
  return Math.min(Math.max(1, pageCount), Math.max(1, Math.trunc(page)))
}

export async function runPageRenderSafely(
  render: () => Promise<void>,
  onError: (message: string) => void,
): Promise<boolean> {
  try {
    await render()
    return true
  } catch (reason) {
    onError(reason instanceof Error ? reason.message : String(reason))
    return false
  }
}

export function createPageResourceScope() {
  let render: { cancel(): void } | null = null
  let text: { cancel(): void } | null = null
  let page: { cleanup(): unknown } | null = null
  let disposed = false
  return {
    setRenderTask(value: { cancel(): void }) {
      render = value
    },
    setTextLayer(value: { cancel(): void }) {
      text = value
    },
    setPage(value: { cleanup(): unknown }) {
      page = value
    },
    dispose() {
      if (disposed) return
      disposed = true
      render?.cancel()
      text?.cancel()
      page?.cleanup()
    },
  }
}

export function createViewerState(): ViewerState {
  return { sourceKey: null, loadToken: 0, document: null, page: 1, scale: 1 }
}

export function beginDocumentLoad(
  state: ViewerState,
  sourceKey: string,
): { state: ViewerState; token: number } {
  const token = state.loadToken + 1
  return {
    token,
    state: { ...state, sourceKey, loadToken: token, document: null, page: 1 },
  }
}

export function acceptDocumentLoad(
  state: ViewerState,
  token: number,
  document: ViewerDocumentState,
): ViewerState {
  if (token !== state.loadToken) return state
  return { ...state, document, page: Math.min(state.page, Math.max(1, document.pageCount)) }
}

export function setViewerPage(state: ViewerState, page: number): ViewerState {
  const max = Math.max(1, state.document?.pageCount ?? 1)
  return { ...state, page: Math.min(max, Math.max(1, Math.trunc(page))) }
}

export function setViewerScale(state: ViewerState, scale: number): ViewerState {
  if (!Number.isFinite(scale)) return state
  return { ...state, scale: Math.min(MAX_SCALE, Math.max(MIN_SCALE, scale)) }
}

export function createViewerResourceScope(
  options: {
    revokeObjectUrl?: (url: string) => void
  } = {},
) {
  const renders = new Set<CancellableRenderTask>()
  const documents = new Set<DestroyableDocument>()
  const objectUrls = new Set<string>()
  const revokeObjectUrl = options.revokeObjectUrl ?? URL.revokeObjectURL.bind(URL)
  let disposed = false

  return {
    trackRenderTask(task: CancellableRenderTask) {
      if (disposed) task.cancel()
      else renders.add(task)
      return () => renders.delete(task)
    },
    trackDocument(document: DestroyableDocument) {
      if (disposed) void document.destroy()
      else documents.add(document)
      return () => documents.delete(document)
    },
    trackObjectUrl(url: string) {
      if (disposed) revokeObjectUrl(url)
      else objectUrls.add(url)
      return () => objectUrls.delete(url)
    },
    async dispose() {
      if (disposed) return
      disposed = true
      for (const render of renders) render.cancel()
      renders.clear()
      const pending = [...documents].map(async (document) => document.destroy())
      documents.clear()
      for (const url of objectUrls) revokeObjectUrl(url)
      objectUrls.clear()
      await Promise.allSettled(pending)
    },
  }
}
