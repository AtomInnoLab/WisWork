import { MAX_PRESENTATION_SOURCE_TEXT_CHARS } from '@wiswork/pptx-engine/presentation-source-limits'
import type { PresentationPlan } from '@wiswork/pptx-engine/presentation-plan'
import { presentationSourceAttachmentId } from '@wiswork/pptx-engine/presentation-plan'
import type { PresentationSourceAudit } from '@wiswork/pptx-engine/presentation-delivery-report'
import { createHash } from 'node:crypto'

export function canonicalSourceLocator(value: string | undefined): string | undefined {
  const match = /^第\s*([1-9]\d{0,5})\s*(页|段)$/.exec(value?.trim() ?? '')
  return match ? `第 ${Number(match[1])} ${match[2]}` : undefined
}

/** A fetched snapshot must be attributed to the exact requested URL, including its query. */
export function matchesFetchedSourceUrl(uri: string, sourceUrlHash: unknown): boolean {
  if (sourceUrlHash === undefined) return true // User-uploaded originals have no observed URL.
  if (typeof sourceUrlHash !== 'string' || !/^[a-f0-9]{64}$/.test(sourceUrlHash))
    throw new Error('invalid_state')
  try {
    const url = new URL(uri)
    return createHash('sha256').update(url.toString()).digest('hex') === sourceUrlHash
  } catch {
    return false
  }
}

/** Read-only, document-bound audit. Literal presence never verifies factual support. */
export async function auditPresentationSources(
  plan: PresentationPlan,
  documentId: string,
  attachments: (body: Record<string, unknown>, signal: AbortSignal) => Promise<unknown>,
  signal: AbortSignal,
): Promise<PresentationSourceAudit[]> {
  const sources: PresentationSourceAudit[] = []
  for (const source of plan.sources) {
    if (signal.aborted) throw new Error('aborted')
    const attachmentId = presentationSourceAttachmentId(source)
    if (!attachmentId) continue
    try {
      const details = (await attachments(
        { operation: 'attachment_metadata', documentId, attachmentId },
        signal,
      )) as { attachmentId?: unknown; sourceUrlHash?: unknown }
      if (details.attachmentId !== attachmentId) throw new Error('invalid_state')
      if (!matchesFetchedSourceUrl(source.uri, details.sourceUrlHash)) {
        sources.push({ sourceId: source.id, attachmentId, status: 'source_mismatch' })
        continue
      }
      const result = (await attachments(
        {
          operation: 'attachment_match_excerpt',
          documentId,
          attachmentId,
          excerpt: source.excerpt,
          ...(canonicalSourceLocator(source.locator)
            ? { locator: canonicalSourceLocator(source.locator) }
            : {}),
        },
        signal,
      )) as {
        attachmentId: string
        status: PresentationSourceAudit['status']
        offset?: number
        locator?: string
      }
      if (
        result.attachmentId !== attachmentId ||
        !['found', 'not_found', 'empty_excerpt', 'not_ready', 'unsupported'].includes(
          result.status,
        ) ||
        (result.locator !== undefined && !/^第 [1-9]\d{0,5} (页|段)$/.test(result.locator)) ||
        (result.status === 'found'
          ? !Number.isSafeInteger(result.offset) ||
            result.offset! < 0 ||
            result.offset! > MAX_PRESENTATION_SOURCE_TEXT_CHARS
          : result.offset !== undefined || result.locator !== undefined)
      )
        throw new Error('invalid_state')
      sources.push({
        sourceId: source.id,
        attachmentId,
        status: result.status,
        ...(result.status === 'found' ? { offset: result.offset } : {}),
        ...(result.locator ? { locator: result.locator } : {}),
      })
    } catch (error) {
      if (!(error instanceof Error) || error.message !== 'not_found') throw error
      sources.push({ sourceId: source.id, attachmentId, status: 'missing' })
    }
  }
  return sources
}
