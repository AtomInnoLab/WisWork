import type { PdfViewportLike } from './types.js'

export function viewportClientToPdfPoint(
  rect: { left: number; top: number; width: number; height: number },
  viewport: PdfViewportLike,
  clientX: number,
  clientY: number,
): { x: number; y: number } | null {
  const x = clientX - rect.left
  const y = clientY - rect.top
  if (x < 0 || y < 0 || x > rect.width || y > rect.height) return null
  const [pdfX, pdfY] = viewport.convertToPdfPoint(x, y)
  if (pdfX === undefined || pdfY === undefined) return null
  return { x: pdfX, y: pdfY }
}

export function pdfPointToViewportCss(
  viewport: PdfViewportLike,
  x: number,
  y: number,
): { left: number; top: number } {
  const [left, top] = viewport.convertToViewportPoint(x, y)
  if (left === undefined || top === undefined) throw new Error('Invalid PDF viewport transform')
  return { left, top }
}
