import { buildPresentationDeliveryReport } from '@wiswork/pptx-engine/presentation-delivery-report'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark.js'
export async function deliveryReportFixture() {
  const plan = benchmarkPlan()
  const metadata = {
    projectId: plan.projectId,
    documentId: 'd',
    requestId: 'r',
    planRevision: 1,
    inputDigest: 'a'.repeat(64),
    planDigest: 'b'.repeat(64),
  }
  const identity = {
    projectId: metadata.projectId,
    documentId: metadata.documentId,
    requestId: metadata.requestId,
    inputDigest: metadata.inputDigest,
    planDigest: metadata.planDigest,
  }
  return buildPresentationDeliveryReport({
    plan,
    deck: benchmarkDeck(),
    metadata,
    pageStates: plan.slides.map((page) => ({ pageId: page.id, state: 'pending' })),
    reviews: [],
    issueLedger: { version: 1, ...identity, revision: 0, actions: [] },
  })
}
