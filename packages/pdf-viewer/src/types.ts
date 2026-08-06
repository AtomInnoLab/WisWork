export interface ViewerDocumentState {
  pageCount: number
}

export interface ViewerState {
  sourceKey: string | null
  loadToken: number
  document: ViewerDocumentState | null
  page: number
  scale: number
}

export type PdfSource =
  { kind: 'bytes'; data: Uint8Array | ArrayBuffer } | { kind: 'url'; url: string }

export interface PdfPoint {
  page: number
  x: number
  y: number
}

export type ViewerLocation = PdfPoint

export interface CancellableRenderTask {
  cancel(): void
}

export interface DestroyableDocument {
  destroy(): void | Promise<void>
}

export interface PdfViewportLike {
  width: number
  height: number
  convertToPdfPoint(x: number, y: number): number[]
  convertToViewportPoint(x: number, y: number): number[]
}
