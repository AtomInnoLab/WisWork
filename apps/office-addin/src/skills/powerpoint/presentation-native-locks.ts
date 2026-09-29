import { parsePresentationPlan } from '@wiswork/pptx-engine/presentation-plan'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import { parsePresentationImportSource } from '@wiswork/project-store/presentation-import-source'
import type { PresentationGenerationOptions } from './presentation-generation.js'
import type { PresentationImportRecord } from './presentation-delivery.js'
import { presentationHostAssociations } from './presentation-host-associations.js'
import { presentationMutationScope } from './presentation-mutation-scope.js'
import { validPresentationImportRecord } from './presentation-page-delivery.js'
import type { PresentationLockReview, StructuredProposal } from '../../agent/proposal-controller.js'

type Ready = Extract<PresentationLockReview, { state: 'ready' }>
type Options = Pick<
  PresentationGenerationOptions,
  'available' | 'request' | 'documentId' | 'lastProject'
> & {
  listReceipts?(): { key: string; record: PresentationImportRecord }[]
  hostSlideIds(signal?: AbortSignal): Promise<string[]>
}

/** These internally constructed proposals only append pages or update local records. */
function preservesExistingPages(proposal: StructuredProposal): boolean {
  return (
    proposal.toolName === proposal.operation &&
    [
      'import_generated_presentation',
      'import_presentation_production',
      'release_existing_presentation_change',
      'release_existing_presentation_batch',
      'release_existing_presentation_page_change',
      'release_slide_chart_values_change',
      'stage_presentation_page_replacement',
      'stage_existing_presentation_page_change',
      'discard_presentation_page_replacement',
      'discard_existing_presentation_page_change',
    ].includes(proposal.operation)
  )
}

/** Fresh execution evidence; never consume the project card's cached associations. */
export async function readPresentationNativeLocks(
  options: Options,
  proposal: StructuredProposal,
  signal: AbortSignal,
): Promise<Ready | undefined> {
  try {
    return await readNativeLocks(options, proposal, signal)
  } catch {
    throw new Error('presentation_lock_review_unavailable')
  }
}

async function readNativeLocks(
  options: Options,
  proposal: StructuredProposal,
  signal: AbortSignal,
): Promise<Ready | undefined> {
  if (proposal.impact.host !== 'powerpoint' || preservesExistingPages(proposal)) return
  const fail = () => new Error('presentation_lock_review_unavailable')
  const check = () => {
    if (signal.aborted) throw fail()
  }
  const documentId = await options.documentId()
  check()
  const projectId = options.lastProject()
  const receipts = options.listReceipts?.() ?? []
  const receiptValue = canonicalPresentationValue(receipts)
  if (receipts.length > 64 || (projectId && !/^[A-Za-z0-9_-]{1,128}$/.test(projectId))) throw fail()
  const imports = receipts.flatMap((entry) => {
    const match = /^(production\/)?([A-Za-z0-9_-]{1,128})\/([A-Za-z0-9_-]{1,128})$/.exec(entry.key)
    if (!match || !validPresentationImportRecord(entry.record)) throw fail()
    return entry.record.documentId === documentId
      ? [
          {
            ...entry,
            projectId: match[2]!,
            requestId: match[3]!,
            source: match[1] ? ('production' as const) : ('compiled' as const),
          },
        ]
      : []
  })
  const importedIds = imports.flatMap(({ record }) =>
    record.checkpoint
      ? record.checkpoint.completed.map((page) => page.slideId)
      : (record.slideIds ?? []),
  )
  if (
    new Set(importedIds).size !== importedIds.length ||
    new Set(receipts.map((entry) => entry.key)).size !== receipts.length
  )
    throw fail()
  const projects = [
    ...new Set([...imports.map((entry) => entry.projectId), ...(projectId ? [projectId] : [])]),
  ].sort()
  if (!projects.length) return
  if (!options.listReceipts || !options.available()) throw fail()
  const current = async () => {
    check()
    if (
      (await options.documentId()) !== documentId ||
      options.lastProject() !== projectId ||
      canonicalPresentationValue(options.listReceipts!()) !== receiptValue ||
      !options.available()
    )
      throw fail()
    check()
  }
  const read = async (body: Record<string, unknown>, maxBytes: number) => {
    check()
    const response = await options.request({ ...body, documentId }, signal)
    check()
    if (!response.ok) throw fail()
    const text = await response.text()
    check()
    if (new TextEncoder().encode(text).byteLength > maxBytes) throw fail()
    return JSON.parse(text)
  }
  const plan = async (id: string) => {
    const value = await read({ operation: 'get_plan', projectId: id }, 512 * 1024)
    if (value.error === 'not_found' && Object.keys(value).length === 1) return null
    if (
      Object.keys(value).sort().join(',') !== 'plan,projectId,revision' ||
      value.projectId !== id ||
      !Number.isSafeInteger(value.revision) ||
      value.revision < 1
    )
      throw fail()
    const parsed = parsePresentationPlan(value.plan)
    if (parsed.projectId !== id) throw fail()
    return { revision: value.revision as number, plan: parsed }
  }
  const pages: Ready['pages'] = []
  const plans = []
  const sourceGroups = []
  for (const id of projects) {
    const saved = await plan(id)
    plans.push({ projectId: id, saved })
    await current()
    if (!saved || !saved.plan.slides.some((slide) => slide.locked)) continue
    const own = imports.filter((entry) => entry.projectId === id)
    const sources = []
    for (let start = 0; start < own.length; start += 4) {
      const batch = await Promise.all(
        own.slice(start, start + 4).map(async (entry) => {
          const value = parsePresentationImportSource(
            await read(
              {
                operation: 'read_import_source',
                projectId: id,
                requestId: entry.requestId,
                source: entry.source,
              },
              64 * 1024,
            ),
          )
          if (
            value.documentId !== documentId ||
            value.projectId !== id ||
            value.requestId !== entry.requestId ||
            value.source !== entry.source
          )
            throw fail()
          return value
        }),
      )
      sources.push(...batch)
      await current()
    }
    sourceGroups.push({ saved, own, sources })
  }
  await current()
  const slideIds = sourceGroups.length ? await options.hostSlideIds(signal) : []
  if (
    !Array.isArray(slideIds) ||
    slideIds.length > 4096 ||
    new Set(slideIds).size !== slideIds.length ||
    slideIds.some((id) => typeof id !== 'string' || !id || id.length > 256)
  )
    throw fail()
  await current()
  const scope = presentationMutationScope(proposal)
  for (const { saved, own, sources } of sourceGroups) {
    const associations = presentationHostAssociations(
      saved.plan,
      saved.revision,
      documentId,
      own,
      [],
      slideIds,
      sources,
    )
    if (
      associations.legacyImports ||
      associations.uncertainImports ||
      associations.unverifiedSources
    )
      throw fail()
    for (const slide of saved.plan.slides.filter((page) => page.locked)) {
      const hostPages = associations.pages
        .find((page) => page.pageId === slide.id)!
        .hostPages.filter(
          (page) =>
            page.presence === 'present' && (scope === undefined || scope.includes(page.slideId)),
        )
      if (hostPages.length)
        pages.push({
          projectId: saved.plan.projectId,
          pageId: slide.id,
          title: slide.title,
          slideIds: hostPages.map((page) => page.slideId).sort(),
        })
    }
  }
  // A plan may be locked/unlocked or created while source/host reads were pending.
  for (const entry of plans) {
    if (
      canonicalPresentationValue(await plan(entry.projectId)) !==
      canonicalPresentationValue(entry.saved)
    )
      throw fail()
    await current()
  }
  const bytes = new TextEncoder().encode(
    canonicalPresentationValue({ documentId, projectId, receipts, plans, slideIds, pages }),
  )
  const token = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)))
    .map((byte) => byte.toString(16).padStart(2, '0'))
    .join('')
  await current()
  return { state: 'ready', token, pages }
}
