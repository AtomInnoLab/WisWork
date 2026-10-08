import { expect, it } from 'vitest'
import { upsertToolActivity } from '../src/tool-activity'

type Chip = { callId?: string; running?: boolean; summary: string }

it('matches out-of-order parallel completions by call identity, not the last spinner', () => {
  const a = { callId: 'turn:a', running: true, summary: 'Read A' }
  const b = { callId: 'turn:b', running: true, summary: 'Read B' }
  const original: Chip[] = [a, b]
  const first = upsertToolActivity(original, { callId: 'turn:a', summary: 'A complete' })
  expect(first).toEqual([{ callId: 'turn:a', summary: 'A complete' }, b])
  const second = upsertToolActivity(first, { callId: 'turn:b', summary: 'B rejected' })
  expect(second).toEqual([
    { callId: 'turn:a', summary: 'A complete' },
    { callId: 'turn:b', summary: 'B rejected' },
  ])
  expect(original).toEqual([a, b])
})

it('does not duplicate a repeated start or resurrect a completed card', () => {
  const start = { callId: 'turn:a', running: true, summary: 'Read' }
  expect(upsertToolActivity([start], start)).toEqual([start])
  const complete = { callId: 'turn:a', summary: 'Done' }
  expect(upsertToolActivity([complete], start)).toEqual([complete])
})

it('appends a no-execution error without removing an unrelated pending call or legacy history', () => {
  const legacy: Chip = { summary: 'Historical tool' }
  const pending = { callId: 'turn:b', running: true, summary: 'Pending' }
  const error = { callId: 'turn:a', summary: 'Invalid input' }
  expect(upsertToolActivity([legacy, pending], error)).toEqual([legacy, pending, error])
})
