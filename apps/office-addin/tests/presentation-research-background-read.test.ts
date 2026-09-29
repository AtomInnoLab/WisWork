import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { createOfficeHostRuntime } from '../src/agent/host-runtime'
import { createPresentationService } from '../../shell/src/main/presentation-service'
import { researchDraft } from '../../../packages/project-store/tests/fixtures/presentation-research'
const roots: string[] = []
afterEach(() => roots.splice(0).forEach((root) => rmSync(root, { recursive: true, force: true })))
it('a pending actual background summary does not own the Agent build slot', async () => {
  const root = mkdtempSync(join(tmpdir(), 'research-background-read-'))
  roots.push(root)
  const service = createPresentationService({ userDataPath: root })
  let seen!: () => void,
    release!: () => void,
    blocked = false
  const arrived = new Promise<void>((resolve) => {
      seen = resolve
    }),
    gate = new Promise<void>((resolve) => {
      release = resolve
    })
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      lastProject: () => undefined,
      rememberProject: async () => {},
      documentId: async () => 'doc',
      request: async (body, signal) => {
        const response = await service(body, signal ?? new AbortController().signal)
        if ((body as { operation: string }).operation === 'research_list' && !blocked) {
          blocked = true
          seen()
          await gate
        }
        return new Response(Buffer.from(response).toString())
      },
    },
  })
  try {
    const first = await runtime.skill.executeTool({
      id: 'first',
      name: 'build_research_ledger',
      input: { project_id: 'p', ledger_id: 'a', expected_revision: 0, draft: researchDraft() },
    })
    expect(first.isError, first.output).toBeFalsy()
    await arrived
    const second = await runtime.skill.executeTool({
      id: 'second',
      name: 'build_research_ledger',
      input: { project_id: 'p', ledger_id: 'b', expected_revision: 2, draft: researchDraft() },
    })
    expect(second.isError, second.output).toBeFalsy()
    expect(JSON.parse(second.output).record.id).toBe('b')
  } finally {
    release()
    runtime.dispose()
  }
})
