import type { PresentationPlan } from '@wiswork/pptx-engine/presentation-plan'
import type { PresentationImportRecord } from './presentation-delivery.js'
import { validPresentationImportRecord } from './presentation-page-delivery.js'
import {
  parsePresentationImportSource,
  type PresentationImportSource,
} from '@wiswork/project-store/presentation-import-source'

export interface PresentationHostAssociations {
  pages: {
    pageId: string
    hostPages: {
      requestId: string
      slideId: string
      planRevision?: number
      revisionRelation: 'current' | 'historical' | 'unknown'
      presence: 'present' | 'missing'
      position?: number
      sourceProof?: 'digest' | 'identity' | 'unverified'
    }[]
  }[]
  legacyImports: number
  uncertainImports: number
  unverifiedSources?: number
}
export function presentationHostAssociations(
  plan: PresentationPlan,
  revision: number,
  documentId: string,
  receipts: { key: string; record: PresentationImportRecord }[],
  tasks: { requestId: string; planRevision: number }[],
  slideIds: string[],
  sources: PresentationImportSource[] = [],
): PresentationHostAssociations {
  const invalid = () => new Error('presentation_host_association_invalid')
  if (
    !Number.isSafeInteger(revision) ||
    revision < 1 ||
    receipts.length > 64 ||
    tasks.length > 32 ||
    slideIds.length > 4096 ||
    slideIds.some((id) => typeof id !== 'string' || !id || id.length > 256) ||
    new Set(slideIds).size !== slideIds.length ||
    new Set(receipts.map((entry) => entry.key)).size !== receipts.length ||
    new Set(tasks.map((task) => task.requestId)).size !== tasks.length ||
    tasks.some(
      (task) =>
        !Number.isSafeInteger(task.planRevision) ||
        task.planRevision < 1 ||
        task.planRevision > revision,
    )
  )
    throw invalid()
  let descriptors: PresentationImportSource[]
  try {
    descriptors = sources.map(parsePresentationImportSource)
  } catch {
    throw invalid()
  }
  if (
    descriptors.length > 64 ||
    new Set(descriptors.map((p) => `${p.documentId}/${p.source}/${p.projectId}/${p.requestId}`))
      .size !== descriptors.length
  )
    throw invalid()
  const pages: PresentationHostAssociations['pages'] = plan.slides.map((page) => ({
    pageId: page.id,
    hostPages: [],
  }))
  const present = new Set(slideIds)
  const claimed = new Set<string>()
  let legacyImports = 0,
    uncertainImports = 0,
    unverifiedSources = 0
  for (const { key, record } of receipts) {
    const match = /^(production\/)?([A-Za-z0-9_-]{1,128})\/([A-Za-z0-9_-]{1,128})$/.exec(key)
    if (
      !match ||
      !validPresentationImportRecord(record) ||
      Boolean(match[1]) !== (record.checkpoint?.version === 2)
    )
      throw invalid()
    if (record.documentId !== documentId || match[2] !== plan.projectId) continue
    if (record.state === 'pending' && (!record.checkpoint || record.checkpoint.inFlight))
      uncertainImports++
    const requestId = match[3]!
    const taskRevision = match[1]
      ? tasks.find((task) => task.requestId === requestId)?.planRevision
      : undefined
    const descriptor = descriptors.find(
      (p) =>
        p.documentId === documentId &&
        p.projectId === plan.projectId &&
        p.requestId === requestId &&
        p.source === (match[1] ? 'production' : 'compiled'),
    )
    let sourceProof: 'digest' | 'identity' | 'unverified' = 'unverified'
    if (descriptor) {
      if (
        (descriptor.planRevision !== undefined && descriptor.planRevision > revision) ||
        (taskRevision !== undefined && descriptor.planRevision !== taskRevision)
      )
        throw invalid()
      if (record.checkpoint) {
        if (
          record.checkpoint.artifactDigest !== descriptor.artifactDigest ||
          JSON.stringify(record.checkpoint.sourceSlideIds) !==
            JSON.stringify(descriptor.pages.map((page) => page.sourceSlideId)) ||
          (record.checkpoint.version === 2 &&
            JSON.stringify(record.checkpoint.pageIds) !==
              JSON.stringify(descriptor.pages.map((page) => page.id)))
        )
          throw invalid()
        sourceProof = 'digest'
      } else if (record.state === 'complete') {
        if (record.slideIds!.length !== descriptor.pages.length) throw invalid()
        sourceProof = 'identity'
      }
    } else unverifiedSources++
    const completedPages =
      record.checkpoint?.version === 2
        ? record.checkpoint.completed.map((completed, index) => ({
            pageId: record.checkpoint!.pageIds![index]!,
            slideId: completed.slideId,
          }))
        : descriptor && sourceProof !== 'unverified'
          ? record.checkpoint
            ? record.checkpoint.completed.map((completed, index) => ({
                pageId: descriptor.pages[index]!.id,
                slideId: completed.slideId,
              }))
            : record.slideIds!.map((slideId, index) => ({
                pageId: descriptor.pages[index]!.id,
                slideId,
              }))
          : undefined
    if (!completedPages) {
      legacyImports++
      continue
    }
    const planRevision = descriptor?.planRevision ?? taskRevision
    for (const completed of completedPages) {
      if (claimed.has(completed.slideId)) throw invalid()
      claimed.add(completed.slideId)
      const page = pages.find((page) => page.pageId === completed.pageId)
      if (!page) continue
      page.hostPages.push({
        requestId,
        slideId: completed.slideId,
        sourceProof,
        ...(planRevision === undefined ? {} : { planRevision }),
        revisionRelation:
          planRevision === undefined
            ? 'unknown'
            : planRevision === revision
              ? 'current'
              : 'historical',
        presence: present.has(completed.slideId) ? 'present' : 'missing',
        ...(present.has(completed.slideId)
          ? { position: slideIds.indexOf(completed.slideId) + 1 }
          : {}),
      })
    }
  }
  return { pages, legacyImports, uncertainImports, unverifiedSources }
}
