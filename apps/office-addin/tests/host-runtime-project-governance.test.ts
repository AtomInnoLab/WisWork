import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PresentationStore, PresentationLifecycleStore } from '@wiswork/project-store'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../../shell/src/main/presentation-service'
import { createPresentationProjectGovernanceService } from '../../shell/src/main/presentation-project-governance'
import { createOfficeHostRuntime } from '../src/agent/host-runtime'
const roots: string[] = []
afterEach(() => {
  vi.unstubAllGlobals()
  for (const path of roots.splice(0)) rmSync(path, { recursive: true, force: true })
})
async function fixture(enabled = true) {
  const root = mkdtempSync(join(tmpdir(), 'runtime-governance-'))
  roots.push(root)
  const plan = benchmarkPlan(),
    documentId = 'doc'
  new PresentationStore(root).savePlan(plan.projectId, documentId, 0, plan)
  const service = createPresentationService({ userDataPath: root }),
    governance = createPresentationProjectGovernanceService({
      userDataPath: root,
      acquireProjectLock: async () => () => {},
    })
  let cap = true,
    primary = true,
    sessionId = 'first',
    loseAck = false,
    otherPc = false
  const rpc = vi.fn(async (body: unknown, signal?: AbortSignal) => {
    if (otherPc) return new Response(JSON.stringify({ lifecycle: null }))
    const result = await governance(body, signal ?? new AbortController().signal)
    if (loseAck && (body as { operation: string }).operation === 'project_deletion_confirm')
      throw Error('lost_ack')
    return new Response(JSON.stringify(result))
  })
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: () => false } },
  })
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      projectGovernanceEnabled: enabled,
      available: () => primary,
      documentId: async () => documentId,
      lastProject: () => plan.projectId,
      setLastProject: () => {},
      request: async (body: unknown, signal?: AbortSignal) =>
        new Response(
          new Uint8Array(await service(body, signal ?? new AbortController().signal)).buffer,
        ),
      governanceAvailable: () => cap,
      governanceRequest: rpc,
      governanceSessionId: () => sessionId,
      readGovernanceAttempt: () => saved,
      writeGovernanceAttempt: (_scope: unknown, value: unknown) => {
        saved = value
      },
    } as any,
  })
  let saved: unknown
  await runtime.presentation!.refresh()
  return {
    runtime,
    rpc,
    plan,
    getSaved: () => saved,
    setSaved: (value: unknown) => {
      saved = value
    },
    root,
    documentId,
    async seedPending() {
      const response = await rpc({
        operation: 'project_deletion_preview',
        documentId,
        projectId: plan.projectId,
      })
      const { preview } = await response.json()
      const id = '11111111-1111-4111-8111-111111111111'
      saved = {
        version: 1,
        scope: { documentId, projectId: plan.projectId },
        expectedRevision: preview.expectedRevision,
        confirmationToken: preview.confirmationToken,
        deletionId: id,
      }
      const life = new PresentationLifecycleStore(root)
      const record = life.initialize({ documentId, projectId: plan.projectId })
      life.beginDeletion({ documentId, projectId: plan.projectId }, record.revision, {
        deletionId: id,
        reason: 'user',
        resources: preview.resources.map((r: any) => ({
          resourceId: r.resourceId,
          kind: r.kind,
          ownership: r.ownership === 'project_exclusive' ? 'project_exclusive' : 'unproven',
        })),
      })
      return id
    },
    setCap: (v: boolean) => {
      cap = v
    },
    setPrimary: (v: boolean) => {
      primary = v
    },
    setSession: (v: string) => {
      sessionId = v
    },
    loseAck: () => {
      loseAck = true
    },
    otherPc: () => {
      otherPc = true
    },
  }
}
it('wires actual PC governance into current project without adding Agent tools or host mutations', async () => {
  const f = await fixture()
  try {
    expect(f.runtime.governance).toBeDefined()
    expect(f.runtime.presentation!.snapshot().project?.projectId).toBe(f.plan.projectId)
    expect(f.rpc).not.toHaveBeenCalled()
    await f.runtime.governance!.preview()
    expect(f.runtime.governance!.snapshot().preview).toBeDefined()
    expect(f.runtime.skill.tools.some((t) => t.name.includes('project_deletion'))).toBe(false)
  } finally {
    f.runtime.dispose()
  }
})
it('fails closed when negotiated capability or primary session is unavailable', async () => {
  const f = await fixture()
  try {
    f.setCap(false)
    await f.runtime.governance!.preview()
    expect(f.rpc).not.toHaveBeenCalled()
    f.setCap(true)
    f.setPrimary(false)
    await f.runtime.governance!.preview()
    expect(f.rpc).not.toHaveBeenCalled()
  } finally {
    f.runtime.dispose()
  }
})
it('rejects a preview across re-pairing and preserves unknown original intent without replay', async () => {
  const f = await fixture()
  try {
    await f.runtime.governance!.preview()
    f.setSession('second')
    await f.runtime.governance!.confirmDeletion()
    expect(f.rpc.mock.calls.map((c) => (c[0] as any).operation)).toEqual([
      'project_deletion_preview',
    ])
    await f.runtime.governance!.preview()
    f.loseAck()
    await f.runtime.governance!.confirmDeletion()
    expect(f.getSaved()).toBeDefined()
    f.runtime.clearSession()
    f.setSession('third')
    await f.runtime.presentation!.refresh()
    f.otherPc()
    await f.runtime.governance!.checkAttempt()
    expect((f.rpc.mock.calls.at(-1)?.[0] as any).operation).toBe('project_lifecycle_read')
    await f.runtime.governance!.resumeDeletion()
    expect(
      f.rpc.mock.calls.filter((c) => (c[0] as any).operation === 'project_deletion_confirm'),
    ).toHaveLength(1)
    expect(
      f.rpc.mock.calls.some((c) => (c[0] as any).operation === 'project_deletion_resume'),
    ).toBe(false)
  } finally {
    f.runtime.dispose()
  }
})

it('defaults to no governance controller or RPC', async () => {
  const f = await fixture(false)
  try {
    expect(f.runtime.governance).toBeUndefined()
    expect(f.rpc).not.toHaveBeenCalled()
  } finally {
    f.runtime.dispose()
  }
})
it('recovers accepted same original intent through readControl after new pairing then explicit resume', async () => {
  const f = await fixture()
  try {
    const id = await f.seedPending()
    f.runtime.clearSession()
    f.setSession('new-pair')
    await f.runtime.presentation!.refresh()
    await f.runtime.governance!.checkAttempt()
    expect(f.runtime.governance!.snapshot().phase).toBe('partial')
    expect(
      f.rpc.mock.calls.some((c) => (c[0] as any).operation === 'project_deletion_resume'),
    ).toBe(false)
    await f.runtime.governance!.resumeDeletion()
    expect(
      f.rpc.mock.calls.filter((c) => (c[0] as any).operation === 'project_deletion_resume'),
    ).toHaveLength(1)
    expect(f.runtime.governance!.snapshot().deletion?.deletionId).toBe(id)
    expect(f.runtime.governance!.snapshot().phase).toBe('deleted')
  } finally {
    f.runtime.dispose()
  }
})
