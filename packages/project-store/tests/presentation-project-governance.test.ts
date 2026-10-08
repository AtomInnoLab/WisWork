import { expect, it } from 'vitest'
import {
  parsePresentationGovernancePreview,
  parsePresentationGovernanceDeletionReport,
  parsePresentationGovernanceLifecycle,
  parsePresentationGovernanceAudit,
  parsePresentationProjectDeletionAttempt,
} from '../src/presentation-project-governance'
const at = '2026-09-30T00:00:00.000Z',
  anonymousProjectId = '12345678-1234-4123-8123-123456789abc'
const counts = { pending: 0, removed: 0, referenceRemoved: 0, retained: 0, failed: 0 }
const resource = {
  kind: 'project',
  ownership: 'project_exclusive',
  disposition: 'candidate',
  resourceId: 'resource_x',
  fileCount: 1,
  bytes: 12,
}
const lifecycle = {
  version: 1,
  documentId: 'doc',
  projectId: 'p',
  revision: 0,
  state: 'active',
  anonymousProjectId,
  createdAt: at,
  updatedAt: at,
  policy: { contentRetentionDays: null, auditRetentionDays: null },
  audit: [{ sequence: 0, at, action: 'created', result: 'accepted', counts }],
}
it('accepts immutable bounded real public preview and full lifecycle/audit shapes', () => {
  const preview = {
    expectedRevision: null,
    confirmationToken: 'a'.repeat(64),
    resources: [resource],
    governanceRetained: true,
  }
  expect(parsePresentationGovernancePreview(preview)).toEqual(preview)
  expect(parsePresentationGovernanceLifecycle(lifecycle)).toEqual(lifecycle)
  const audit = {
    version: 1,
    anonymousProjectId,
    auditRetentionDays: null,
    events: lifecycle.audit,
  }
  expect(parsePresentationGovernanceAudit(audit)).toEqual(audit)
  preview.resources[0]!.bytes = 33
  expect(parsePresentationGovernancePreview({ ...preview, resources: [resource] })).not.toBe(
    preview,
  )
})
it('accepts partial and deleted receipts only with internally consistent retained counts', () => {
  const partial = {
    state: 'partial',
    revision: 2,
    deletionId: 'delete',
    projectContentRetained: true,
    counts: { removed: 0, pending: 0, failed: 0, retained: 1 },
    retained: [(({ disposition: _, ...r }) => r)(resource)],
  }
  expect(parsePresentationGovernanceDeletionReport(partial)).toEqual(partial)
  const deleted = {
    ...partial,
    state: 'deleted',
    projectContentRetained: false,
    counts: { removed: 1, pending: 0, failed: 0, retained: 0 },
    retained: [],
  }
  expect(parsePresentationGovernanceDeletionReport(deleted)).toEqual(deleted)
  expect(() =>
    parsePresentationGovernanceDeletionReport({ ...deleted, retained: partial.retained }),
  ).toThrow()
})
it('rejects unbounded, unknown, scope-leaking and forged public data', () => {
  const preview = {
    expectedRevision: 0,
    confirmationToken: 'a'.repeat(64),
    resources: [resource],
    governanceRetained: true,
  }
  for (const value of [
    { ...preview, rawPath: '/private' },
    { ...preview, resources: [{ ...resource, kind: 'unknown' }] },
    { ...preview, resources: [{ ...resource, disposition: 'retained' }] },
    {
      ...preview,
      resources: Array.from({ length: 4097 }, (_, i) => ({ ...resource, resourceId: 'r' + i })),
    },
  ])
    expect(() => parsePresentationGovernancePreview(value)).toThrow()
  expect(() => parsePresentationGovernanceLifecycle({ ...lifecycle, revision: 1 })).toThrow()
  expect(() =>
    parsePresentationGovernanceAudit({
      version: 1,
      anonymousProjectId,
      auditRetentionDays: null,
      events: lifecycle.audit,
      documentId: 'doc',
    }),
  ).toThrow()
})
it('retains only finite frozen intent with a legal multibyte 4096-character document', () => {
  const attempt = {
    version: 1,
    scope: { documentId: '文'.repeat(4096), projectId: 'p' },
    expectedRevision: null,
    confirmationToken: 'a'.repeat(64),
    deletionId: 'delete',
  }
  expect(parsePresentationProjectDeletionAttempt(attempt)).toEqual(attempt)
  expect(() =>
    parsePresentationProjectDeletionAttempt({ ...attempt, resourcePaths: ['/private'] }),
  ).toThrow()
  expect(() =>
    parsePresentationProjectDeletionAttempt({
      ...attempt,
      scope: { ...attempt.scope, documentId: '文'.repeat(4097) },
    }),
  ).toThrow()
})
it('accepts real long audit above 512KiB and rejects 2MiB overflow', () => {
  const events = Array.from({ length: 8192 }, (_, sequence) => ({
    sequence,
    at,
    action: sequence ? 'policy_updated' : 'created',
    result: 'accepted',
    counts,
  }))
  const audit = { version: 1, anonymousProjectId, auditRetentionDays: null, events }
  expect(new TextEncoder().encode(JSON.stringify(audit)).byteLength).toBeGreaterThan(512 * 1024)
  expect(parsePresentationGovernanceAudit(audit)).toEqual(audit)
  expect(() =>
    parsePresentationGovernanceAudit({ ...audit, extra: 'x'.repeat(2 * 1024 * 1024) }),
  ).toThrow()
})
