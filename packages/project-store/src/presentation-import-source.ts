/** Preserve the original import receipt serialization, including property order. */
export function presentationImportContent(artifact: {
  documentId: string
  projectId: string
  requestId: string
  planRevision?: number
  pptxBase64: string
  pagePptxBase64?: string[]
  pages?: { id: string; title: string; sourceSlideId: string }[]
}): string {
  return artifact.pagePptxBase64 === undefined
    ? artifact.pptxBase64
    : JSON.stringify({
        documentId: artifact.documentId,
        projectId: artifact.projectId,
        requestId: artifact.requestId,
        planRevision: artifact.planRevision,
        pages: artifact.pages,
        pagePptxBase64: artifact.pagePptxBase64,
      })
}
export interface PresentationImportSource {
  version: 1
  documentId: string
  projectId: string
  requestId: string
  source: 'compiled' | 'production'
  planRevision?: number
  artifactDigest: string
  pages: { id: string; title: string; sourceSlideId?: string }[]
}
export function parsePresentationImportSource(value: unknown): PresentationImportSource {
  const p = value as PresentationImportSource | undefined
  const id = (value: unknown) => typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
  if (
    !p ||
    typeof p !== 'object' ||
    Array.isArray(p) ||
    Object.keys(p).sort().join(',') !==
      (p.planRevision === undefined
        ? 'artifactDigest,documentId,pages,projectId,requestId,source,version'
        : 'artifactDigest,documentId,pages,planRevision,projectId,requestId,source,version') ||
    p.version !== 1 ||
    typeof p.documentId !== 'string' ||
    !p.documentId ||
    p.documentId.length > 2048 ||
    !id(p.projectId) ||
    !id(p.requestId) ||
    !['compiled', 'production'].includes(p.source) ||
    typeof p.artifactDigest !== 'string' ||
    !/^[a-f0-9]{64}$/.test(p.artifactDigest) ||
    (p.planRevision !== undefined &&
      (!Number.isSafeInteger(p.planRevision) || p.planRevision < 1)) ||
    (p.source === 'production' && p.planRevision === undefined) ||
    !Array.isArray(p.pages) ||
    !p.pages.length ||
    p.pages.length > 32 ||
    p.pages.some(
      (page) =>
        !page ||
        Object.keys(page).sort().join(',') !==
          (page.sourceSlideId === undefined ? 'id,title' : 'id,sourceSlideId,title') ||
        !id(page.id) ||
        page.id.length > 80 ||
        typeof page.title !== 'string' ||
        !page.title ||
        page.title.length > 300 ||
        (page.sourceSlideId !== undefined &&
          (!/^[1-9]\d{0,9}#$/.test(page.sourceSlideId) ||
            Number(page.sourceSlideId.slice(0, -1)) < 256 ||
            Number(page.sourceSlideId.slice(0, -1)) > 0xffffffff)) ||
        (p.source === 'production' && page.sourceSlideId === undefined),
    ) ||
    new Set(p.pages.map((page) => page.id)).size !== p.pages.length ||
    (p.source === 'compiled' &&
      p.pages.some((page) => page.sourceSlideId !== undefined) &&
      (p.pages.some((page) => page.sourceSlideId === undefined) ||
        new Set(p.pages.map((page) => page.sourceSlideId)).size !== p.pages.length))
  )
    throw new Error('presentation_import_source_invalid')
  return structuredClone(p)
}
