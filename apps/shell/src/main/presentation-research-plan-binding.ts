import {
  parsePresentationResearchRecord,
  type PresentationResearchRecord,
} from '@wiswork/project-store/presentation-research'
import { assertPresentationResearchBinding } from '@wiswork/pptx-engine/presentation-research-binding'
import type { PresentationPlan } from '@wiswork/pptx-engine/presentation-plan'
export type PresentationResearchReader = (ledgerId: string) => Promise<PresentationResearchRecord>
/** Resolve the exact archived version named by a plan; never substitute latest research. */
export async function readBoundPresentationResearch(
  plan: PresentationPlan,
  documentId: string,
  projectId: string,
  readResearch?: PresentationResearchReader,
  signal?: AbortSignal,
): Promise<PresentationResearchRecord | undefined> {
  if (!plan.research) return undefined
  if (signal?.aborted) throw new Error('aborted')
  if (!readResearch) throw new Error('research_unavailable')
  let raw: PresentationResearchRecord
  try {
    raw = await readResearch(plan.research.ledgerId)
  } catch {
    if (signal?.aborted) throw new Error('aborted')
    throw new Error('research_unavailable')
  }
  if (signal?.aborted) throw new Error('aborted')
  try {
    const record = parsePresentationResearchRecord(raw)
    if (record.documentId !== documentId || record.projectId !== projectId)
      throw new Error('research_binding_invalid')
    assertPresentationResearchBinding(plan, record)
    return record
  } catch {
    throw new Error('research_binding_invalid')
  }
}
