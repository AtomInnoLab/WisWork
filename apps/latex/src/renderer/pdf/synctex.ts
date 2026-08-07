export interface PageRect {
  left: number
  top: number
  width: number
  height: number
}

export function clientToPdfPoint(
  rect: PageRect,
  scale: number,
  clientX: number,
  clientY: number,
): { x: number; y: number } | null {
  if (!Number.isFinite(scale) || scale <= 0) return null
  const cssX = clientX - rect.left
  const cssY = clientY - rect.top
  if (cssX < 0 || cssY < 0 || cssX > rect.width || cssY > rect.height) return null
  return { x: cssX / scale, y: cssY / scale }
}

export function pdfPointToClient(
  rect: PageRect,
  scale: number,
  x: number,
  y: number,
): { clientX: number; clientY: number } | null {
  if (!Number.isFinite(scale) || scale <= 0 || x < 0 || y < 0) return null
  const clientX = rect.left + x * scale
  const clientY = rect.top + y * scale
  if (clientX > rect.left + rect.width || clientY > rect.top + rect.height) return null
  return { clientX, clientY }
}
