import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { CSSProperties, MouseEvent as ReactMouseEvent } from 'react'
import { GlobalWorkerOptions, TextLayer, getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy, RenderTask } from 'pdfjs-dist'
import workerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import type { PdfPoint, PdfSource, ViewerLocation } from './types.js'
import { pdfPointToViewportCss, viewportClientToPdfPoint } from './coordinates.js'
import {
  assertDocumentPageBudget,
  clampViewerPageToDocument,
  computeCanvasBudget,
  createPageResourceScope,
  getDocumentPages,
  getPageWindow,
  nearestPageToViewport,
  resetViewerPageForSource,
  runPageRenderSafely,
} from './viewer-state.js'

GlobalWorkerOptions.workerSrc = workerUrl

export interface ReadonlyPdfPageProps {
  document: PDFDocumentProxy
  page: number
  scale: number
  visible?: boolean
  rotation?: number
  className?: string
  onClick?: (point: PdfPoint) => void
  onDoubleClick?: (point: PdfPoint) => void
  location?: ViewerLocation | null
  onViewportSize?: (page: number, width: number, height: number) => void
}

export function ReadonlyPdfPage({
  document: pdfDocument,
  page,
  scale,
  visible = true,
  rotation = 0,
  className,
  onClick,
  onDoubleClick,
  location,
  onViewportSize,
}: ReadonlyPdfPageProps) {
  const holderRef = useRef<HTMLDivElement>(null)
  const viewportRef = useRef<ReturnType<
    Awaited<ReturnType<PDFDocumentProxy['getPage']>>['getViewport']
  > | null>(null)

  useEffect(() => {
    const holder = holderRef.current
    if (!visible || !holder) return
    let disposed = false
    const resources = createPageResourceScope()
    void runPageRenderSafely(
      async () => {
        const pdfPage = await pdfDocument.getPage(page)
        if (disposed) return
        resources.setPage(pdfPage)
        const viewport = pdfPage.getViewport({
          scale,
          rotation: (pdfPage.rotate + rotation) % 360,
        })
        onViewportSize?.(page, viewport.width, viewport.height)
        viewportRef.current = viewport
        const dpr = Math.min(window.devicePixelRatio || 1, 2)
        const budget = computeCanvasBudget(viewport.width, viewport.height, dpr)
        const canvas = window.document.createElement('canvas')
        canvas.width = budget.width
        canvas.height = budget.height
        canvas.style.width = `${Math.floor(viewport.width)}px`
        canvas.style.height = `${Math.floor(viewport.height)}px`
        const renderTask: RenderTask = pdfPage.render({
          canvas,
          viewport,
          transform:
            budget.allocationScale === 1
              ? undefined
              : [budget.allocationScale, 0, 0, budget.allocationScale, 0, 0],
        })
        resources.setRenderTask(renderTask)
        try {
          await renderTask.promise
        } catch {
          return
        }
        if (disposed) return
        const textLayerElement = window.document.createElement('div')
        textLayerElement.className = 'textLayer'
        const textLayer = new TextLayer({
          textContentSource: pdfPage.streamTextContent(),
          container: textLayerElement,
          viewport,
        })
        resources.setTextLayer(textLayer)
        const marker =
          location && location.page === page ? window.document.createElement('span') : null
        if (marker && location) {
          const point = pdfPointToViewportCss(viewport, location.x, location.y)
          marker.className = 'pdf-sync-marker'
          marker.style.left = `${point.left}px`
          marker.style.top = `${point.top}px`
        }
        holder.replaceChildren(canvas, textLayerElement, ...(marker ? [marker] : []))
        try {
          await textLayer.render()
        } catch {
          // A page leaving the viewport is an expected cancellation path.
        }
      },
      (message) => {
        resources.dispose()
        viewportRef.current = null
        if (disposed) return
        const alert = window.document.createElement('div')
        alert.setAttribute('role', 'alert')
        alert.className = 'pdf-page-error'
        alert.textContent = message
        holder.replaceChildren(alert)
      },
    )
    return () => {
      disposed = true
      resources.dispose()
      viewportRef.current = null
      holder.replaceChildren()
    }
  }, [pdfDocument, location, onViewportSize, page, rotation, scale, visible])

  const emitPoint = (
    event: ReactMouseEvent<HTMLDivElement>,
    callback?: (point: PdfPoint) => void,
  ) => {
    const viewport = viewportRef.current
    if (!callback || !viewport) return
    const rect = event.currentTarget.getBoundingClientRect()
    const point = viewportClientToPdfPoint(rect, viewport, event.clientX, event.clientY)
    if (point) callback({ page, ...point })
  }

  return (
    <div
      ref={holderRef}
      className={className ?? 'pdf-page-content'}
      data-pdf-page={page}
      onClick={(event) => emitPoint(event, onClick)}
      onDoubleClick={(event) => emitPoint(event, onDoubleClick)}
    />
  )
}

export interface ReadonlyPdfViewerProps {
  source: PdfSource | null
  sourceKey?: string | number
  assetBaseUrl?: string
  className?: string
  location?: ViewerLocation | null
  onPageChange?: (page: number) => void
  onScaleChange?: (scale: number) => void
  onPageClick?: (point: PdfPoint) => void
  onPageDoubleClick?: (point: PdfPoint) => void
}

function controlledDocumentSource(source: PdfSource) {
  if (source.kind === 'bytes') return { data: source.data }
  const parsed = new URL(source.url)
  if (parsed.protocol !== 'wiswork-latex-pdf:' && parsed.protocol !== 'blob:') {
    throw new Error(`Unsupported PDF URL protocol: ${parsed.protocol}`)
  }
  return { url: source.url }
}

export function ReadonlyPdfViewer({
  source,
  sourceKey,
  assetBaseUrl = new URL('pdfjs/', document.baseURI).href,
  className,
  location,
  onPageChange,
  onScaleChange,
  onPageClick,
  onPageDoubleClick,
}: ReadonlyPdfViewerProps) {
  const rootRef = useRef<HTMLDivElement>(null)
  const loadToken = useRef(0)
  const [pdfDocument, setPdfDocument] = useState<PDFDocumentProxy | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [scale, setScale] = useState(1)
  const [currentPage, setCurrentPage] = useState(1)
  const [pageSizes, setPageSizes] = useState<Record<number, { width: number; height: number }>>({})
  const scrollFrame = useRef<number | null>(null)
  const sourceIdentity =
    sourceKey ?? (source?.kind === 'url' ? source.url : source?.data.byteLength)

  useEffect(() => {
    const token = ++loadToken.current
    setError(null)
    setPdfDocument(null)
    setPageSizes({})
    setCurrentPage((page) => resetViewerPageForSource(page))
    if (!source) return
    let loadingTask: ReturnType<typeof getDocument> | null = null
    try {
      loadingTask = getDocument({
        ...controlledDocumentSource(source),
        cMapUrl: `${assetBaseUrl}cmaps/`,
        cMapPacked: true,
        standardFontDataUrl: `${assetBaseUrl}standard_fonts/`,
        wasmUrl: `${assetBaseUrl}wasm/`,
      })
      void loadingTask.promise
        .then((document) => {
          if (token !== loadToken.current) return
          try {
            assertDocumentPageBudget(document.numPages)
          } catch (reason) {
            void document.destroy()
            throw reason
          }
          setCurrentPage((page) => clampViewerPageToDocument(page, document.numPages))
          setPdfDocument(document)
        })
        .catch((reason: unknown) => {
          if (token === loadToken.current) {
            setError(reason instanceof Error ? reason.message : String(reason))
          }
        })
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    }
    return () => {
      loadToken.current += 1
      void loadingTask?.destroy()
    }
  }, [assetBaseUrl, source, sourceIdentity])

  const pages = useMemo(
    () => getPageWindow(currentPage, pdfDocument?.numPages ?? 0),
    [currentPage, pdfDocument],
  )
  const documentPages = useMemo(() => getDocumentPages(pdfDocument?.numPages ?? 0), [pdfDocument])
  const renderedPages = useMemo(() => new Set(pages), [pages])

  const recordPageSize = useCallback((page: number, width: number, height: number) => {
    setPageSizes((current) => {
      const existing = current[page]
      if (existing?.width === width && existing.height === height) return current
      return { ...current, [page]: { width, height } }
    })
  }, [])

  useEffect(
    () => () => {
      if (scrollFrame.current !== null) window.cancelAnimationFrame(scrollFrame.current)
    },
    [],
  )

  const scrollToPage = (page: number) => {
    setCurrentPage(page)
    onPageChange?.(page)
    window.requestAnimationFrame(() => {
      rootRef.current
        ?.querySelector<HTMLElement>(`[data-page="${page}"]`)
        ?.scrollIntoView({ block: 'start' })
    })
  }

  const updateCurrentPageFromScroll = () => {
    if (scrollFrame.current !== null) return
    scrollFrame.current = window.requestAnimationFrame(() => {
      scrollFrame.current = null
      const root = rootRef.current
      if (!root) return
      const rootTop = root.getBoundingClientRect().top
      const page = nearestPageToViewport(
        Array.from(root.querySelectorAll<HTMLElement>('[data-page]'), (element) => {
          const rect = element.getBoundingClientRect()
          return {
            page: Number(element.dataset.page),
            top: rect.top - rootTop,
            bottom: rect.bottom - rootTop,
          }
        }),
      )
      if (page && page !== currentPage) {
        setCurrentPage(page)
        onPageChange?.(page)
      }
    })
  }

  useEffect(() => {
    if (!location || !pdfDocument || location.page > pdfDocument.numPages) return
    setCurrentPage(location.page)
    onPageChange?.(location.page)
  }, [location, onPageChange, pdfDocument])

  useEffect(() => {
    if (!location || location.page !== currentPage) return
    rootRef.current
      ?.querySelector<HTMLElement>(`[data-page="${location.page}"]`)
      ?.scrollIntoView({ block: 'center' })
  }, [currentPage, location, pages])

  const updateScale = (next: number) => {
    const clamped = Math.min(4, Math.max(0.5, next))
    setScale(clamped)
    setPageSizes({})
    onScaleChange?.(clamped)
  }

  const pageStyle = (page: number): CSSProperties | undefined =>
    location?.page === page
      ? { outline: '2px solid var(--accent, #5b8def)', outlineOffset: '3px' }
      : undefined

  return (
    <section className={className ?? 'readonly-pdf-viewer'}>
      <div className="readonly-pdf-toolbar">
        <button
          type="button"
          disabled={currentPage <= 1}
          onClick={() => scrollToPage(Math.max(1, currentPage - 1))}
        >
          ‹
        </button>
        <output>
          {currentPage}/{pdfDocument?.numPages ?? 0}
        </output>
        <button
          type="button"
          disabled={!pdfDocument || currentPage >= pdfDocument.numPages}
          onClick={() => scrollToPage(Math.min(pdfDocument?.numPages ?? 1, currentPage + 1))}
        >
          ›
        </button>
        <button type="button" onClick={() => updateScale(scale - 0.1)} aria-label="Zoom out">
          −
        </button>
        <output>{Math.round(scale * 100)}%</output>
        <button type="button" onClick={() => updateScale(scale + 0.1)} aria-label="Zoom in">
          +
        </button>
      </div>
      <div ref={rootRef} className="readonly-pdf-scroll" onScroll={updateCurrentPageFromScroll}>
        {error && <div role="alert">{error}</div>}
        {!source && <div className="readonly-pdf-empty">No PDF</div>}
        {documentPages.map((page) => {
          const measured = pageSizes[page]
          const fallback = pageSizes[currentPage] ?? Object.values(pageSizes)[0]
          const size = measured ?? fallback ?? { width: 612 * scale, height: 792 * scale }
          return (
            <div
              key={page}
              data-page={page}
              className="readonly-pdf-page"
              style={{ ...pageStyle(page), width: size.width, minHeight: size.height }}
            >
              {pdfDocument && renderedPages.has(page) && (
                <ReadonlyPdfPage
                  document={pdfDocument}
                  page={page}
                  scale={scale}
                  visible
                  location={location}
                  onClick={(point) => {
                    setCurrentPage(page)
                    onPageChange?.(page)
                    onPageClick?.(point)
                  }}
                  onDoubleClick={onPageDoubleClick}
                  onViewportSize={recordPageSize}
                />
              )}
            </div>
          )
        })}
      </div>
    </section>
  )
}
