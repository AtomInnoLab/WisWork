import {
  readBoundPresentationResearch,
  type PresentationResearchReader,
} from './presentation-research-plan-binding'
import { PresentationStore } from '@wiswork/project-store'
import { parsePresentationIssueActionInput } from '@wiswork/project-store/presentation-issue'
import { parsePresentationPlan } from '@wiswork/pptx-engine/presentation-plan'
import { parsePresentationDeck } from '@wiswork/pptx-engine/presentation'
import { parsePresentationClaimReview } from '@wiswork/pptx-engine/presentation-claim-review'
import { buildPresentationDeliveryReport } from '@wiswork/pptx-engine/presentation-delivery-report'
import { auditPresentationSources } from './presentation-source-audit'

/** Called under the service's project lock; this path never invokes Office or compilation. */
export async function handlePresentationDeliveryReport(
  request: Record<string, unknown>,
  store: PresentationStore,
  attachments: (body: Record<string, unknown>, signal: AbortSignal) => Promise<unknown>,
  signal: AbortSignal,
  readResearch?: PresentationResearchReader,
) {
  const projectId = request.projectId as string
  const documentId = request.documentId as string
  const requestId = request.requestId as string
  const production = store.production(projectId, documentId, requestId)
  if (!production) throw new Error('not_found')
  const plan = parsePresentationPlan(production.plan.plan)
  const researchRecord = await readBoundPresentationResearch(
    plan,
    documentId,
    projectId,
    readResearch,
    signal,
  )
  const reviews = store.listClaimReviews(projectId, documentId, requestId).map((record) =>
    parsePresentationClaimReview({
      ...(record.review as Record<string, unknown>),
      version: 1,
      projectId: record.projectId,
      requestId: record.requestId,
      reviewId: record.reviewId,
      planRevision: record.planRevision,
      inputDigest: record.inputDigest,
      planDigest: record.planDigest,
      createdAt: record.createdAt,
      checks: {
        support: 'agent_reviewed',
        sourceAuthority: 'not_verified',
        timeliness: 'not_verified',
        host: 'not_checked',
      },
    }),
  )
  let issueLedger = store.issueActions(projectId, documentId, requestId)
  const build = async () =>
    buildPresentationDeliveryReport({
      plan,
      ...(researchRecord ? { researchRecord } : {}),
      sourceAudit: await auditPresentationSources(
        parsePresentationPlan(production.plan.plan),
        documentId,
        attachments,
        signal,
      ),
      deck: parsePresentationDeck(production.deck),
      metadata: {
        projectId,
        documentId,
        requestId,
        planRevision: production.plan.revision,
        inputDigest: production.inputDigest,
        planDigest: production.planDigest,
      },
      pageStates: production.pages,
      reviews,
      issueLedger,
    })
  const report = await build()
  if (signal.aborted) throw new Error('aborted')
  if (request.operation === 'production_delivery_report') return report
  const action = parsePresentationIssueActionInput(request.action)
  const previous = issueLedger.actions.find((item) => item.actionId === action.actionId)
  if (previous) {
    // Store checks exact content before CAS; retry remains valid after the evidence changes.
    issueLedger = store.appendIssueAction(
      projectId,
      documentId,
      requestId,
      request.expectedRevision as number,
      action,
    )
    return build()
  }
  const issue = report.pages
    .flatMap((page) => page.issues)
    .find((item) => item.id === action.issueId)
  if (!issue || issue.digest !== action.issueDigest) throw new Error('issue_changed')
  if (signal.aborted) throw new Error('aborted')
  issueLedger = store.appendIssueAction(
    projectId,
    documentId,
    requestId,
    request.expectedRevision as number,
    action,
  )
  // Cancellation after persistence must not claim the action was not saved.
  return build()
}
