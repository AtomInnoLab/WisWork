import { afterEach, expect, it, vi } from 'vitest'
const hooks = vi.hoisted(() => ({
  options: undefined as any,
  clear: vi.fn(),
  propose: vi.fn(async () => ({ id: 'master-proposal' })),
  execute: vi.fn(async () => ({ output: '{}', mutated: false })),
}))
vi.mock('../src/skills/powerpoint/presentation-native-master.js', () => ({
  createPresentationNativeMasterSkill: (options: any) => {
    hooks.options = options
    return {
      name: 'native_master',
      description: 'Native master',
      tools: [
        {
          name: 'inspect_slide_master_change',
          description: 'inspect',
          inputSchema: { type: 'object' },
        },
      ],
      executeTool: hooks.execute,
      propose: hooks.propose,
      clear: hooks.clear,
      beginMutation: vi.fn(),
      endMutation: vi.fn(),
    }
  },
}))
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'
afterEach(() => {
  vi.unstubAllGlobals()
  hooks.options = undefined
  vi.clearAllMocks()
})
it('connects durable master transport, CAS, dispatch, and lifecycle with live capability guards', async () => {
  let capability = true,
    paired = true,
    api = true
  vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => api } } })
  const request = vi.fn(async () => new Response('{}')),
    write = vi.fn(async () => {}),
    read = vi.fn()
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => paired,
      documentId: async () => 'doc',
      request: vi.fn(async () => new Response('{}')),
      lastProject: () => undefined,
      rememberProject: async () => {},
      masterBackupAvailable: () => capability,
      masterBackupRequest: request,
      readNativeMasterChange: read,
      writeNativeMasterChange: write,
    } as any,
  })
  expect(hooks.options).toBeDefined()
  expect(hooks.options.available()).toBe(true)
  expect(hooks.options.request).toBe(request)
  expect(hooks.options.readNativeMasterChange).toBe(read)
  await hooks.options.writeNativeMasterChange({ changeId: 'change' }, undefined)
  expect(write).toHaveBeenCalledWith({ changeId: 'change' }, undefined)
  await runtime.skill.executeTool({
    id: 'inspect',
    name: 'inspect_slide_master_change',
    input: { change_id: 'change' },
  })
  expect(hooks.execute).toHaveBeenCalled()
  const input = {
    program: {
      version: 2,
      operations: [
        {
          op: 'set_master_background',
          master_id: 'master',
          fill: { type: 'solid', color: '#000000', transparency: 0 },
        },
      ],
    },
  }
  const proposed = await runtime.skill.executeTool({
    id: 'master',
    name: 'edit_slide_master',
    input,
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  expect(hooks.propose).toHaveBeenCalledWith(input.program.operations, undefined, undefined)
  capability = false
  const blocked = await runtime.skill.executeTool({
    id: 'master-blocked',
    name: 'edit_slide_master',
    input,
  })
  expect(blocked.isError).toBe(true)
  expect(hooks.propose).toHaveBeenCalledTimes(1)
  capability = true
  for (const change of [
    () => {
      capability = false
    },
    () => {
      capability = true
      paired = false
    },
    () => {
      paired = true
      api = false
    },
  ]) {
    change()
    expect(hooks.options.available()).toBe(false)
  }
  runtime.clearSession()
  expect(hooks.clear).toHaveBeenCalled()
  runtime.dispose()
})
it('does not construct a durable master skill without the negotiated backup channel', () => {
  vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } })
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      documentId: async () => 'doc',
      request: vi.fn(),
      lastProject: () => undefined,
      rememberProject: async () => {},
    },
  })
  expect(hooks.options).toBeUndefined()
  runtime.dispose()
})
