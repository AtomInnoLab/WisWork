import { presentationPlanClaims } from '@wiswork/pptx-engine/presentation-plan'
import { mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it, vi } from 'vitest'
import { PresentationStore } from '@wiswork/project-store'
import { handlePresentationDeliveryReport } from '../src/main/presentation-delivery-report'
import { createPresentationService } from '../src/main/presentation-service'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const decode = (value: Uint8Array) => JSON.parse(Buffer.from(value).toString())
const files = (root: string) =>
  readdirSync(root, { recursive: true, withFileTypes: true })
    .filter((item) => item.isFile())
    .map((item) => [
      join(item.parentPath, item.name),
      readFileSync(join(item.parentPath, item.name)).toString('base64'),
    ])
async function setup() {
  const userDataPath = mkdtempSync(join(tmpdir(), 'delivery-service-'))
  roots.push(userDataPath)
  const compile = vi.fn()
  const service = createPresentationService({ userDataPath, compile })
  const deck = benchmarkPlannedDeck(),
    plan = benchmarkPlan()
  for (const source of plan.sources) source.uri = `attachment:${'a'.repeat(64)}`
  deck.claims = presentationPlanClaims(plan)
  const base = { documentId: 'doc', projectId: deck.id, requestId: 'run' }
  const store = new PresentationStore(userDataPath)
  store.beginProduction(deck.id, 'doc', 'run', deck, { revision: 1, plan })
  const call = (operation: string, extra = {}, signal = new AbortController().signal) =>
    service({ ...base, operation, ...extra }, signal).then(decode)
  return { userDataPath, compile, service, call, store, base }
}
it('reads frozen reports without writes, records explanations, and retries after restart', async () => {
  const f = await setup(),
    before = files(f.userDataPath)
  const report = await handlePresentationDeliveryReport(
    { ...f.base, operation: 'production_delivery_report' },
    f.store,
    new AbortController().signal,
  )
  expect(report).not.toHaveProperty('error')
  expect(files(f.userDataPath)).toEqual(before)
  const issue = report.pages.flatMap((page: { issues: unknown[] }) => page.issues)[0]
  const action = {
    actionId: 'a',
    issueId: issue.id,
    issueDigest: issue.digest,
    state: 'explained',
    note: 'Needs human confirmation',
  }
  const updated = await f.call('production_record_issue_action', { expectedRevision: 0, action })
  expect(updated.issueLedger.revision).toBe(1)
  expect(updated.checks.content).toBe('needs_review')
  const restarted = createPresentationService({ userDataPath: f.userDataPath, compile: f.compile })
  expect(
    decode(
      await restarted(
        { ...f.base, operation: 'production_record_issue_action', expectedRevision: 0, action },
        new AbortController().signal,
      ),
    ),
  ).toEqual(updated)
  expect(
    await f.call('production_record_issue_action', {
      expectedRevision: 1,
      action: { ...action, note: 'changed' },
    }),
  ).toEqual({ error: 'request_conflict' })
  expect(
    await f.call('production_record_issue_action', {
      expectedRevision: 0,
      action: { ...action, actionId: 'b' },
    }),
  ).toEqual({ error: 'revision_conflict' })
  expect(f.compile).not.toHaveBeenCalled()
})
it('rejects stale/missing issues, cross-document/request, extra keys, and aborted writes', async () => {
  const f = await setup()
  const action = {
    actionId: 'a',
    issueId: 'missing',
    issueDigest: 'a'.repeat(64),
    state: 'deferred',
    note: 'Later',
  }
  expect(await f.call('production_record_issue_action', { expectedRevision: 0, action })).toEqual({
    error: 'issue_changed',
  })
  expect(await f.call('production_delivery_report', { pageId: 'extra' })).toEqual({
    error: 'invalid_request',
  })
  expect(await f.call('production_delivery_report', { documentId: 'other' })).toEqual({
    error: 'document_mismatch',
  })
  expect(await f.call('production_delivery_report', { requestId: 'missing' })).toEqual({
    error: 'not_found',
  })
  const controller = new AbortController()
  controller.abort()
  expect(
    await f.call(
      'production_record_issue_action',
      { expectedRevision: 0, action },
      controller.signal,
    ),
  ).toEqual({ error: 'aborted' })
  expect(f.store.issueActions(f.base.projectId, 'doc', 'run').revision).toBe(0)
})
it('keeps retries idempotent after reviews change and cancels before persistence', async () => {
  const f = await setup()
  const read = () =>
    handlePresentationDeliveryReport(
      { ...f.base, operation: 'production_delivery_report' },
      f.store,
      new AbortController().signal,
    )
  const report = await read()
  const issue = report.pages
    .flatMap((page) => page.issues)
    .find((item) => item.code === 'source_review_missing')!
  expect(issue).toBeDefined()
  const page = report.pages.find((item) => item.issues.includes(issue))!
  const action = {
    actionId: 'retry',
    issueId: issue.id,
    issueDigest: issue.digest,
    state: 'deferred',
    note: 'Await review',
  }
  expect(
    (await f.call('production_record_issue_action', { expectedRevision: 0, action })).issueLedger
      .revision,
  ).toBe(1)
  f.store.saveClaimReview(f.base.projectId, 'doc', 'run', 'review', {
    pageId: page.pageId,
    claimId: issue.claimId,
    sourceId: issue.sourceId,
    attachmentId: 'a'.repeat(64),
    offset: 0,
    maxChars: 1000,
    evidenceDigest: 'b'.repeat(64),
    outcome: 'supported',
    notes: 'Historical review',
    reviewer: 'agent',
  })
  expect(
    (await f.call('production_record_issue_action', { expectedRevision: 0, action })).issueLedger
      .revision,
  ).toBe(1)
  expect(
    await f.call('production_record_issue_action', {
      expectedRevision: 1,
      action: { ...action, actionId: 'changed' },
    }),
  ).toEqual({ error: 'issue_changed' })
  const current = await read()
  const next = current.pages.flatMap((page) => page.issues)[0]!
  const controller = new AbortController()
  const original = f.store.issueActions.bind(f.store)
  const spy = vi.spyOn(f.store, 'issueActions').mockImplementation((...args) => {
    const ledger = original(...args)
    controller.abort()
    return ledger
  })
  await expect(
    handlePresentationDeliveryReport(
      {
        ...f.base,
        operation: 'production_record_issue_action',
        expectedRevision: 1,
        action: { ...action, actionId: 'cancelled', issueId: next.id, issueDigest: next.digest },
      },
      f.store,
      controller.signal,
    ),
  ).rejects.toThrow('aborted')
  spy.mockRestore()
  expect(original(f.base.projectId, 'doc', 'run').revision).toBe(1)
})
it('serializes competing actions and rejects a stale issue digest without changing the ledger', async () => {
  const f = await setup()
  const report = await f.call('production_delivery_report')
  const issue = report.pages.flatMap(
    (page: { issues: { id: string; digest: string }[] }) => page.issues,
  )[0]
  const action = {
    actionId: 'a',
    issueId: issue.id,
    issueDigest: issue.digest,
    state: 'deferred',
    note: 'Wait',
  }
  expect(
    await f.call('production_record_issue_action', {
      expectedRevision: 0,
      action: { ...action, issueDigest: '0'.repeat(64) },
    }),
  ).toEqual({ error: 'issue_changed' })
  const results = await Promise.all(
    ['a', 'b'].map((actionId) =>
      f.call('production_record_issue_action', {
        expectedRevision: 0,
        action: { ...action, actionId },
      }),
    ),
  )
  expect(results.filter((item) => item.error === 'revision_conflict')).toHaveLength(1)
  expect(results.filter((item) => item.issueLedger?.revision === 1)).toHaveLength(1)
  expect(f.store.issueActions(f.base.projectId, 'doc', 'run').actions).toHaveLength(1)
})
