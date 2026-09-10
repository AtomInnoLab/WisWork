// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { AgentWorkspace, type OfficeWorkspaceUi } from '../src/App.js'
import {
  appendPresentationEvent,
  type OfficePresentationEvent,
} from '../src/agent/presentation-state.js'
import type { OfficeAgentSession, OfficeAgentSnapshot } from '../src/agent/use-office-agent.js'

let root: Root
let container: HTMLDivElement
beforeEach(() => {
  Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })
  vi.useFakeTimers()
  container = document.createElement('div')
  document.body.append(container)
  root = createRoot(container)
})
afterEach(() => {
  act(() => root.unmount())
  container.remove()
  vi.useRealTimers()
})
const original = '# DESIGN.md\n\n## Original design\n\nOcean blue'
const saved = '# DESIGN.md\n\n## PC draft\n\nVolcano black'
const button = (text: string) =>
  Array.from(container.querySelectorAll('button')).find((element) =>
    element.textContent?.includes(text),
  )!

function conversation(markdown = original) {
  const listeners = new Set<() => void>()
  let snapshot: OfficeAgentSnapshot = {
    assistantText: '',
    activity: '',
    busy: false,
    applying: false,
    status: 'idle',
    retryable: false,
    timeline: [{ id: 'user_1', kind: 'user', text: 'Make a presentation' }],
  }
  const publish = () => listeners.forEach((listener) => listener())
  const append = (event: OfficePresentationEvent) => {
    snapshot = { ...snapshot, timeline: appendPresentationEvent(snapshot.timeline, event) }
    publish()
  }
  const plan = (text: string) =>
    append({
      id: 'plan_1',
      kind: 'tool',
      callId: 'plan_call',
      name: 'plan_deck',
      summary: 'Design created',
      state: 'complete',
      output: JSON.stringify({ designMd: text }),
    })
  plan(markdown)
  const session: OfficeAgentSession = {
    snapshot: () => snapshot,
    subscribe: (listener) => {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    send: vi.fn(),
    stop: vi.fn(),
    confirm: vi.fn(),
    reject: vi.fn(),
    retry: vi.fn(),
    logout: vi.fn(),
    authenticationLost: vi.fn(),
    dispose: vi.fn(),
    reviseDesignContract: vi.fn(),
    newTask: vi.fn(() => {
      snapshot = { ...snapshot, timeline: [] }
      publish()
    }),
  }
  return { session, append, plan }
}

const ui: OfficeWorkspaceUi = {
  attachments: () => [],
  skills: () => [],
  skillPackagesEnabled: false,
  upload: vi.fn(),
  clear: vi.fn(),
}
const request = () =>
  vi.fn(async (body: { action: string; documentId: string; markdown?: string }) => ({
    documentId: body.documentId,
    markdown: body.action === 'open' ? body.markdown! : saved,
    revision: body.action === 'open' ? 'original' : 'saved',
  }))

it('retains the current design, PC draft and document identity after more than 100 events, then resets on new task', async () => {
  const c = conversation(),
    pc = request()
  await act(async () =>
    root.render(
      React.createElement(AgentWorkspace, {
        session: c.session,
        ui,
        host: 'powerpoint',
        disconnect: vi.fn(),
        designRequest: pc,
      }),
    ),
  )
  await act(async () => button('DESIGN.md').click())
  await act(async () => button('在 WisWork PC 编辑').click())
  await act(async () => vi.advanceTimersByTimeAsync(5000))
  expect(container.textContent).toContain('PC draft')
  const documentId = pc.mock.calls[0]![0].documentId

  await act(async () => {
    for (let i = 0; i < 105; i++)
      c.append({ id: `activity_${i}`, kind: 'assistant', text: `Step ${i}` })
  })
  expect(c.session.snapshot().timeline).toHaveLength(100)
  expect(container.querySelector('.design-document-entry')).not.toBeNull()
  expect(container.querySelector('[role="dialog"]')?.textContent).toContain('PC draft')
  expect(button('应用设计修改').disabled).toBe(false)
  await act(async () => vi.advanceTimersByTimeAsync(5000))
  expect(pc.mock.calls.at(-1)![0]).toEqual({ action: 'read', documentId })

  await act(async () => button('新对话').click())
  expect(c.session.newTask).toHaveBeenCalledOnce()
  expect(container.querySelector('.design-document-entry')).toBeNull()
  expect(container.querySelector('[role="dialog"]')).toBeNull()
  const calls = pc.mock.calls.length
  await act(async () => vi.advanceTimersByTimeAsync(5000))
  expect(pc).toHaveBeenCalledTimes(calls)
  await act(async () => c.plan('# DESIGN.md\n\n## Next task'))
  await act(async () => button('DESIGN.md').click())
  expect(container.textContent).toContain('Next task')
  expect(container.textContent).not.toContain('PC draft')
  await act(async () => button('在 WisWork PC 编辑').click())
  expect(pc.mock.calls.at(-1)![0].documentId).not.toBe(documentId)
})

it('discards the previous session draft even when replacement timeline event IDs are identical', async () => {
  const first = conversation(),
    second = conversation('# DESIGN.md\n\n## Replacement session'),
    pc = request()
  const render = (session: OfficeAgentSession) =>
    root.render(
      React.createElement(AgentWorkspace, {
        session,
        ui,
        host: 'powerpoint',
        disconnect: vi.fn(),
        designRequest: pc,
      }),
    )
  await act(async () => render(first.session))
  await act(async () => button('DESIGN.md').click())
  await act(async () => button('在 WisWork PC 编辑').click())
  await act(async () => vi.advanceTimersByTimeAsync(5000))
  expect(container.textContent).toContain('PC draft')
  const documentId = pc.mock.calls[0]![0].documentId
  await act(async () => render(second.session))
  expect(container.querySelector('[role="dialog"]')).toBeNull()
  await act(async () => button('DESIGN.md').click())
  expect(container.textContent).toContain('Replacement session')
  expect(container.textContent).not.toContain('PC draft')
  await act(async () => button('在 WisWork PC 编辑').click())
  expect(pc.mock.calls.at(-1)![0].documentId).not.toBe(documentId)
})
