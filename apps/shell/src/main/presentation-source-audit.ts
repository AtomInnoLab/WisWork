import type { PresentationPlan } from '@wiswork/pptx-engine/presentation-plan'
import type { PresentationSourceAudit } from '@wiswork/pptx-engine/presentation-delivery-report'

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
    const match = /^attachment:([a-f0-9]{64})$/.exec(source.uri)
    if (!match) continue
    try {
      const result = (await attachments(
        {
          operation: 'attachment_match_excerpt',
          documentId,
          attachmentId: match[1],
          excerpt: source.excerpt,
        },
        signal,
      )) as { attachmentId: string; status: PresentationSourceAudit['status']; offset?: number }
      if (
        result.attachmentId !== match[1] ||
        !['found', 'not_found', 'empty_excerpt', 'not_ready', 'unsupported'].includes(
          result.status,
        ) ||
        (result.status === 'found'
          ? !Number.isSafeInteger(result.offset) || result.offset! < 0 || result.offset! > 1_000_000
          : result.offset !== undefined)
      )
        throw new Error('invalid_state')
      sources.push({
        sourceId: source.id,
        attachmentId: match[1],
        status: result.status,
        ...(result.status === 'found' ? { offset: result.offset } : {}),
      })
    } catch (error) {
      if (!(error instanceof Error) || error.message !== 'not_found') throw error
      sources.push({ sourceId: source.id, attachmentId: match[1], status: 'missing' })
    }
  }
  return sources
}
