import { expect, it } from 'vitest'
import { createPresentationProjectGovernanceController } from '../src/agent/presentation-project-governance'
const preview = {
  expectedRevision: null,
  confirmationToken: 'a'.repeat(64),
  resources: [],
  governanceRetained: true,
}
it('requires capability and explicit preview and persists original intent before confirm request', async () => {
  const calls: unknown[] = []
  let available = false,
    saved: unknown
  const controller = createPresentationProjectGovernanceController({
    available: () => available,
    documentId: async () => 'doc',
    currentProjectId: () => 'p',
    readAttempt: () => undefined,
    writeAttempt: (_s, v) => {
      saved = v
    },
    request: async (body) => {
      calls.push(body)
      if (body.operation === 'project_deletion_confirm') {
        expect(saved).toMatchObject({
          scope: { documentId: 'doc', projectId: 'p' },
          confirmationToken: preview.confirmationToken,
        })
        throw Error('lost_ack')
      }
      return new Response(JSON.stringify({ preview }))
    },
  })
  await controller.preview()
  expect(calls).toEqual([])
  available = true
  await controller.confirmDeletion()
  expect(calls).toEqual([])
  await controller.preview()
  await controller.confirmDeletion()
  expect(calls).toHaveLength(2)
  expect(controller.snapshot().phase).toBe('unknown')
  expect(saved).toBeDefined()
})
it('clears stale preview when document changes during a response and never auto initializes', async () => {
  let doc = 'doc'
  const calls: unknown[] = []
  const controller = createPresentationProjectGovernanceController({
    available: () => true,
    documentId: async () => doc,
    currentProjectId: () => 'p',
    readAttempt: () => undefined,
    writeAttempt: () => {},
    request: async (body) => {
      calls.push(body)
      doc = 'other'
      return new Response(JSON.stringify({ preview }))
    },
  })
  await controller.preview()
  expect(controller.snapshot().preview).toBeUndefined()
  expect(calls).toHaveLength(1)
})
it('never rebinds a preview token to a changed project or deletes when intent storage fails', async () => {
  let project = 'p',
    failStorage = false
  const calls: Record<string, unknown>[] = []
  const controller = createPresentationProjectGovernanceController({
    available: () => true,
    documentId: async () => 'doc',
    currentProjectId: () => project,
    readAttempt: () => undefined,
    writeAttempt: () => {
      if (failStorage) throw Error('quota')
    },
    request: async (body) => {
      calls.push(body)
      return new Response(JSON.stringify({ preview }))
    },
  })
  await controller.preview()
  project = 'other'
  await controller.confirmDeletion()
  expect(calls).toHaveLength(1)
  expect(controller.snapshot().preview).toBeUndefined()
  await controller.preview()
  failStorage = true
  await controller.confirmDeletion()
  expect(calls).toHaveLength(2)
})
it('reopens original intent without re-previewing or automatically replaying confirmation', async () => {
  const attempt = {
    version: 1 as const,
    scope: { documentId: 'doc', projectId: 'p' },
    expectedRevision: null,
    confirmationToken: 'a'.repeat(64),
    deletionId: 'original',
  }
  const calls: Record<string, unknown>[] = []
  const controller = createPresentationProjectGovernanceController({
    available: () => true,
    documentId: async () => 'doc',
    currentProjectId: () => 'p',
    readAttempt: () => attempt,
    writeAttempt: () => {},
    request: async (body) => {
      calls.push(body)
      return new Response(JSON.stringify({ lifecycle: null }))
    },
  })
  await controller.preview()
  expect(calls).toEqual([])
  expect(controller.snapshot().attempt?.deletionId).toBe('original')
  await controller.checkAttempt()
  await controller.confirmDeletion()
  await controller.resumeDeletion()
  expect(calls.map((x) => x.operation)).toEqual(['project_lifecycle_read'])
})
it('initializes only explicitly after a real preview and never saves policy using a guessed revision', async () => {
  const at = '2026-09-30T00:00:00.000Z',
    counts = { pending: 0, removed: 0, referenceRemoved: 0, retained: 0, failed: 0 }
  const lifecycle = {
    version: 1,
    documentId: 'doc',
    projectId: 'p',
    revision: 0,
    state: 'active',
    anonymousProjectId: '12345678-1234-4123-8123-123456789abc',
    createdAt: at,
    updatedAt: at,
    policy: { contentRetentionDays: null, auditRetentionDays: null },
    audit: [{ sequence: 0, at, action: 'created', result: 'accepted', counts }],
  }
  const calls: Record<string, unknown>[] = []
  const controller = createPresentationProjectGovernanceController({
    available: () => true,
    documentId: async () => 'doc',
    currentProjectId: () => 'p',
    readAttempt: () => undefined,
    writeAttempt: () => {},
    request: async (body) => {
      calls.push(body)
      return new Response(
        JSON.stringify(body.operation === 'project_deletion_preview' ? { preview } : { lifecycle }),
      )
    },
  })
  await controller.setPolicy({ contentRetentionDays: null, auditRetentionDays: null })
  await controller.initializePolicy()
  expect(calls).toEqual([])
  await controller.preview()
  expect(Object.isFrozen(controller.snapshot().preview)).toBe(true)
  await controller.initializePolicy()
  expect(calls.map((x) => x.operation)).toEqual([
    'project_deletion_preview',
    'project_lifecycle_initialize',
  ])
  expect(controller.snapshot().lifecycle?.state).toBe('active')
})
const time = '2026-09-30T00:00:00.000Z',
  zero = { pending: 0, removed: 0, referenceRemoved: 0, retained: 0, failed: 0 }
function activeLifecycle() {
  return {
    version: 1,
    documentId: 'doc',
    projectId: 'p',
    revision: 0,
    state: 'active',
    anonymousProjectId: '12345678-1234-4123-8123-123456789abc',
    createdAt: time,
    updatedAt: time,
    policy: { contentRetentionDays: null, auditRetentionDays: null },
    audit: [{ sequence: 0, at: time, action: 'created', result: 'accepted', counts: zero }],
  }
}
it('saves policy with original exact CAS and cloned values after explicit lifecycle read', async () => {
  const calls: Record<string, unknown>[] = []
  const policy = { contentRetentionDays: 30, auditRetentionDays: null }
  const controller = createPresentationProjectGovernanceController({
    available: () => true,
    documentId: async () => 'doc',
    currentProjectId: () => 'p',
    readAttempt: () => undefined,
    writeAttempt: () => {},
    request: async (body) => {
      calls.push(body)
      const record = activeLifecycle()
      return new Response(
        JSON.stringify({
          lifecycle:
            body.operation === 'project_lifecycle_read'
              ? record
              : {
                  ...record,
                  revision: 1,
                  policy: body.policy,
                  audit: [
                    ...record.audit,
                    {
                      sequence: 1,
                      at: time,
                      action: 'policy_updated',
                      result: 'accepted',
                      counts: zero,
                    },
                  ],
                },
        }),
      )
    },
  })
  await controller.refresh()
  const saving = controller.setPolicy(policy)
  policy.contentRetentionDays = 999
  await saving
  expect(calls[1]).toMatchObject({
    operation: 'project_lifecycle_set_policy',
    expectedRevision: 0,
    policy: { contentRetentionDays: 30, auditRetentionDays: null },
  })
  expect(controller.snapshot().lifecycle?.policy.contentRetentionDays).toBe(30)
})
it('reads unknown acknowledgement and requires separate explicit resume with same deletion id', async () => {
  const attempt = {
    version: 1 as const,
    scope: { documentId: 'doc', projectId: 'p' },
    expectedRevision: null,
    confirmationToken: 'a'.repeat(64),
    deletionId: 'original',
  }
  const active = activeLifecycle(),
    retainedCounts = { ...zero, retained: 1 }
  const lifecycle = {
    ...active,
    revision: 2,
    state: 'deleting',
    deletion: {
      deletionId: 'original',
      reason: 'user',
      resources: [
        { resourceId: 'r', kind: 'project', ownership: 'project_exclusive', status: 'retained' },
      ],
    },
    audit: [
      ...active.audit,
      {
        sequence: 1,
        at: time,
        action: 'deletion_started',
        result: 'accepted',
        counts: { ...zero, pending: 1 },
      },
      {
        sequence: 2,
        at: time,
        action: 'resource_result',
        result: 'partial',
        counts: retainedCounts,
      },
    ],
  }
  const calls: Record<string, unknown>[] = []
  const controller = createPresentationProjectGovernanceController({
    available: () => true,
    documentId: async () => 'doc',
    currentProjectId: () => 'p',
    readAttempt: () => attempt,
    writeAttempt: () => {
      throw Error('must not replace original')
    },
    request: async (body) => {
      calls.push(body)
      return new Response(
        JSON.stringify(
          body.operation === 'project_lifecycle_read'
            ? { lifecycle }
            : {
                deletion: {
                  state: 'partial',
                  revision: 2,
                  deletionId: 'original',
                  projectContentRetained: true,
                  counts: { removed: 0, pending: 0, failed: 0, retained: 1 },
                  retained: [
                    {
                      kind: 'project',
                      ownership: 'project_exclusive',
                      resourceId: 'r',
                      fileCount: 1,
                      bytes: 1,
                    },
                  ],
                },
              },
        ),
      )
    },
  })
  await controller.checkAttempt()
  expect(calls).toHaveLength(1)
  expect(controller.snapshot().phase).toBe('partial')
  await controller.resumeDeletion()
  expect(calls[1]).toEqual({
    operation: 'project_deletion_resume',
    documentId: 'doc',
    projectId: 'p',
    expectedRevision: 2,
    deletionId: 'original',
  })
  expect(controller.snapshot().attempt).toEqual(attempt)
})
it.each(['cancel', 'capability', 'document'] as const)(
  'discards late read after %s change',
  async (mode) => {
    let available = true,
      doc = 'doc',
      resolve!: (value: Response) => void
    const pending = new Promise<Response>((r) => {
      resolve = r
    })
    const controller = createPresentationProjectGovernanceController({
      available: () => available,
      documentId: async () => doc,
      currentProjectId: () => 'p',
      readAttempt: () => undefined,
      writeAttempt: () => {},
      request: async () => pending,
    })
    const reading = controller.preview()
    await new Promise((r) => setTimeout(r, 0))
    if (mode === 'cancel') controller.cancel()
    else if (mode === 'capability') available = false
    else doc = 'other'
    resolve(new Response(JSON.stringify({ preview })))
    await reading
    expect(controller.snapshot().preview).toBeUndefined()
    expect(controller.snapshot().scope).toBeUndefined()
  },
)
