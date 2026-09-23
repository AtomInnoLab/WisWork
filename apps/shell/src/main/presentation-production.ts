import type {
  PresentationStore,
  PresentationProductionRecord,
  PresentationProductionPage,
} from '@wiswork/project-store'
import {
  parsePresentationDeck,
  type PresentationInlineAsset,
} from '@wiswork/pptx-engine/presentation'
import {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'
import { checkPresentationPageContent } from '@wiswork/pptx-engine/presentation-content-check'
import type { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'

const check = (signal: AbortSignal) => {
  if (signal.aborted) throw new Error('aborted')
}
export function presentationProductionSummary(record: PresentationProductionRecord) {
  const deck = parsePresentationDeck(record.deck)
  const compiledCount = record.pages.filter((page) => page.state === 'compiled').length
  return {
    projectId: record.projectId,
    requestId: record.requestId,
    planRevision: record.plan.revision,
    ...(record.revision ? { revision: record.revision } : {}),
    status:
      compiledCount === record.pages.length
        ? 'compiled'
        : record.pages.some((p) => p.state === 'building')
          ? 'building'
          : compiledCount || record.pages.some((p) => p.state === 'failed')
            ? 'partial'
            : 'pending',
    compiledCount,
    total: record.pages.length,
    pages: record.pages.map((page, index) => ({
      id: page.pageId,
      title: deck.slides[index]!.title,
      state: page.state,
      attempt: page.attempt,
      ...(page.error ? { error: page.error } : {}),
    })),
  }
}
function pageResult(
  compiled: Awaited<ReturnType<typeof compilePresentationDeck>>,
  projectId: string,
): NonNullable<PresentationProductionPage['result']> {
  const sourceSlideId = compiled.sourceSlideIds?.[0]
  if (
    !(compiled.bytes instanceof Uint8Array) ||
    !compiled.bytes.byteLength ||
    compiled.bytes.byteLength > 10 * 1024 * 1024
  )
    throw new Error('output_too_large')
  if (
    !Array.isArray(compiled.sourceSlideIds) ||
    compiled.sourceSlideIds.length !== 1 ||
    typeof sourceSlideId !== 'string' ||
    !/^[1-9]\d*#$/.test(sourceSlideId) ||
    !Number.isSafeInteger(Number(sourceSlideId.slice(0, -1))) ||
    Number(sourceSlideId.slice(0, -1)) < 256 ||
    Number(sourceSlideId.slice(0, -1)) > 0xffffffff ||
    compiled.report?.slideCount !== 1 ||
    compiled.report?.deckId !== projectId
  )
    throw new Error('compile_failed')
  if (Buffer.byteLength(JSON.stringify(compiled.report)) > 48 * 1024)
    throw new Error('output_too_large')
  return {
    pptxBase64: Buffer.from(compiled.bytes).toString('base64'),
    sourceSlideId,
    report: compiled.report,
  }
}
export async function handlePresentationProduction(
  request: Record<string, unknown>,
  options: {
    store: PresentationStore
    compile: typeof compilePresentationDeck
    attachments: (request: Record<string, unknown>, signal: AbortSignal) => Promise<unknown>
  },
  signal: AbortSignal,
) {
  const { store, compile, attachments } = options
  const projectId = request.projectId as string,
    documentId = request.documentId as string,
    requestId = request.requestId as string | undefined
  let record = store.production(projectId, documentId, requestId)
  if (request.operation === 'production_rebuild_page') {
    const parent = store.production(projectId, documentId, request.parentRequestId as string)
    if (!parent) throw new Error('not_found')
    if (parent.pages.some((page) => page.state !== 'compiled')) throw new Error('page_not_ready')
    const original = parsePresentationDeck(parent.deck)
    const index = original.slides.findIndex((slide) => slide.id === request.pageId)
    if (index < 0) throw new Error('not_found')
    const slide = request.slide
    if (
      !slide ||
      typeof slide !== 'object' ||
      Array.isArray(slide) ||
      (slide as { id?: unknown }).id !== request.pageId
    )
      throw new Error('invalid_deck')
    let revised: ReturnType<typeof parsePresentationDeck>
    try {
      revised = parsePresentationDeck({
        ...original,
        slides: original.slides.map((page, i) => (i === index ? slide : page)),
      })
    } catch {
      throw new Error('invalid_deck')
    }
    try {
      assertDeckMatchesPresentationPlan(revised, parsePresentationPlan(parent.plan.plan))
    } catch {
      throw new Error('plan_mismatch')
    }
    record = store.deriveProduction(
      projectId,
      documentId,
      parent.requestId,
      requestId!,
      request.pageId as string,
      revised,
    )
  }
  if (request.operation === 'production_begin') {
    const saved = record ? record.plan : store.plan(projectId, documentId)
    if (!saved || saved.revision !== request.planRevision)
      throw new Error(record ? 'request_conflict' : 'revision_conflict')
    const deck = parsePresentationDeck(request.deck)
    try {
      assertDeckMatchesPresentationPlan(deck, parsePresentationPlan(saved.plan))
    } catch {
      throw new Error(record ? 'request_conflict' : 'plan_mismatch')
    }
    record = store.beginProduction(projectId, documentId, requestId!, deck, {
      revision: saved.revision,
      plan: saved.plan,
    })
  }
  if (!record) throw new Error('not_found')
  const deck = parsePresentationDeck(record.deck)
  const plan = parsePresentationPlan(record.plan.plan)
  assertDeckMatchesPresentationPlan(deck, plan)
  if (request.operation === 'production_content_check') {
    check(signal)
    return {
      projectId,
      requestId: record.requestId,
      planRevision: record.plan.revision,
      inputDigest: record.inputDigest,
      planDigest: record.planDigest,
      report: checkPresentationPageContent(plan, deck, request.pageId as string),
    }
  }
  if (request.operation === 'production_page') {
    const page = record.pages.find((p) => p.pageId === request.pageId)
    if (!page) throw new Error('not_found')
    if (page.state !== 'compiled' || !page.result) throw new Error('page_not_ready')
    return {
      projectId,
      requestId: record.requestId,
      pageId: page.pageId,
      planRevision: record.plan.revision,
      status: 'compiled',
      ...page.result,
    }
  }
  if (request.operation !== 'production_run') return presentationProductionSummary(record)
  for (const slide of deck.slides) {
    check(signal)
    const current = record.pages.find((p) => p.pageId === slide.id)!
    if (current.state === 'compiled') continue
    const attempt = current.attempt + 1
    record = store.updateProductionPage(record, slide.id, { state: 'building', attempt })
    let result: NonNullable<PresentationProductionPage['result']>
    let failure: string | undefined
    try {
      const assetIds = new Set(
        slide.elements.flatMap((el) => (el.kind === 'image' ? [el.assetId] : [])),
      )
      const assets: PresentationInlineAsset[] = []
      let assetBytes = 0
      for (const asset of deck.assets.filter((asset) => assetIds.has(asset.id))) {
        check(signal)
        let resolved: PresentationInlineAsset
        if ('attachmentId' in asset) {
          try {
            resolved = {
              ...((await attachments(
                { operation: 'attachment_asset', documentId, attachmentId: asset.attachmentId },
                signal,
              )) as PresentationInlineAsset),
              id: asset.id,
            }
          } catch {
            check(signal)
            throw new Error('asset_unavailable')
          }
        } else resolved = asset
        assetBytes += Buffer.byteLength(resolved.base64, 'base64')
        if (assetBytes > 8 * 1024 * 1024) throw new Error('output_too_large')
        assets.push(resolved)
      }
      check(signal)
      const onePage = {
        ...deck,
        assets,
        claims: deck.claims.filter((claim) => slide.claimIds?.includes(claim.id)),
        slides: [slide],
      }
      result = pageResult(await compile(onePage), projectId)
      check(signal)
    } catch (error) {
      const code = error instanceof Error ? error.message : ''
      failure = signal.aborted
        ? 'aborted'
        : ['output_too_large', 'asset_unavailable'].includes(code)
          ? code
          : code.startsWith('presentation_invalid:') || code.startsWith('presentation_geometry:')
            ? 'invalid_deck'
            : 'compile_failed'
    }
    if (failure) {
      record = store.updateProductionPage(record, slide.id, {
        state: 'failed',
        attempt,
        error: failure,
      })
      check(signal)
      continue
    }
    // Storage errors are not transient compiler failures: stop rather than running ahead of receipts.
    try {
      record = store.updateProductionPage(record, slide.id, {
        state: 'compiled',
        attempt,
        result: result!,
      })
    } catch (error) {
      if (!(error instanceof Error) || error.message !== 'output_too_large') throw error
      record = store.updateProductionPage(record, slide.id, {
        state: 'failed',
        attempt,
        error: 'output_too_large',
      })
    }
  }
  return presentationProductionSummary(record)
}
