import { createHash } from 'node:crypto'
import type { PresentationStore } from '@wiswork/project-store'
import {
  presentationImportContent,
  parsePresentationImportSource,
  type PresentationImportSource,
} from '@wiswork/project-store/presentation-import-source'
import { parsePresentationDeck } from '@wiswork/pptx-engine/presentation'

function binary(base64: unknown): string {
  if (
    typeof base64 !== 'string' ||
    !base64 ||
    base64.length > Math.ceil((10 * 1024 * 1024) / 3) * 4 ||
    Buffer.from(base64, 'base64').toString('base64') !== base64
  )
    throw new Error('invalid_state')
  return base64
}
export function readPresentationImportSource(
  store: PresentationStore,
  projectId: string,
  documentId: string,
  requestId: string,
  source: 'compiled' | 'production',
): PresentationImportSource {
  let content: string, pages: PresentationImportSource['pages'], planRevision: number | undefined
  if (source === 'production') {
    const record = store.production(projectId, documentId, requestId)
    if (!record) throw new Error('not_found')
    if (record.pages.some((page) => page.state !== 'compiled')) throw new Error('page_not_ready')
    const deck = parsePresentationDeck(record.deck)
    planRevision = record.plan.revision
    const mapped = deck.slides.map((slide, index) => ({
      id: slide.id,
      title: slide.title,
      sourceSlideId: record.pages[index]!.result!.sourceSlideId,
    }))
    const pagePptxBase64 = record.pages.map((page) => binary(page.result!.pptxBase64))
    if (
      pagePptxBase64.reduce((bytes, page) => bytes + Buffer.byteLength(page, 'base64'), 0) >
      10 * 1024 * 1024
    )
      throw new Error('output_too_large')
    pages = mapped
    content = presentationImportContent({
      documentId,
      projectId,
      requestId,
      planRevision,
      pages: mapped,
      pptxBase64: '',
      pagePptxBase64,
    })
  } else {
    const record = store.request(projectId, documentId, requestId)
    if (!record) throw new Error('not_found')
    if (record.status !== 'compiled') throw new Error('page_not_ready')
    const deck = parsePresentationDeck(record.deck)
    const result = record.result as {
      pptxBase64?: unknown
      projectId?: string
      requestId?: string
      pages?: PresentationImportSource['pages']
    }
    if (
      (result.projectId !== undefined && result.projectId !== projectId) ||
      (result.requestId !== undefined && result.requestId !== requestId)
    )
      throw new Error('invalid_state')
    pages = result.pages ?? deck.slides.map(({ id, title }) => ({ id, title }))
    if (
      !Array.isArray(pages) ||
      pages.length !== deck.slides.length ||
      pages.some(
        (page, index) =>
          page?.id !== deck.slides[index]!.id || page?.title !== deck.slides[index]!.title,
      )
    )
      throw new Error('invalid_state')
    planRevision = record.plan?.revision
    content = binary(result.pptxBase64)
  }
  try {
    return parsePresentationImportSource({
      version: 1,
      documentId,
      projectId,
      requestId,
      source,
      ...(planRevision === undefined ? {} : { planRevision }),
      artifactDigest: createHash('sha256').update(content).digest('hex'),
      pages,
    })
  } catch {
    throw new Error('invalid_state')
  }
}
