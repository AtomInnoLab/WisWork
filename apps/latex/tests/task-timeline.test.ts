import { describe, expect, it } from 'vitest'
import {
  cancelRunningTimelineEntries,
  completeTimelineEntry,
  failRunningTimelineEntries,
  startTimelineEntry,
} from '../src/renderer/ai/task-timeline.js'

describe('agent task timeline', () => {
  it('moves a tool call from running to success with bounded detail', () => {
    const running = startTimelineEntry([], {
      id: 'read-1',
      name: 'read_project_text',
      input: { path: 'main.tex' },
    })
    expect(running[0]).toMatchObject({ id: 'read-1', kind: 'read', state: 'running' })

    const completed = completeTimelineEntry(running, 'read-1', {
      summary: 'Read main.tex',
      output: 'x'.repeat(10_000),
      mutated: false,
    })
    expect(completed[0]).toMatchObject({ state: 'success', label: 'Read main.tex' })
    expect(completed[0]?.detail?.length).toBeLessThan(10_000)
  })

  it('makes tool failure, agent error and cancellation observable', () => {
    const first = startTimelineEntry([], { id: 'compile-1', name: 'compile_project', input: {} })
    expect(
      completeTimelineEntry(first, 'compile-1', {
        summary: 'Compile failed',
        output: 'Undefined control sequence',
        isError: true,
        mutated: false,
      })[0],
    ).toMatchObject({ kind: 'compile', state: 'error' })

    const second = startTimelineEntry(first, {
      id: 'search-1',
      name: 'search_project_text',
      input: {},
    })
    expect(failRunningTimelineEntries(second, 'Provider disconnected')[1]).toMatchObject({
      state: 'error',
      detail: 'Provider disconnected',
    })

    const third = startTimelineEntry(second, {
      id: 'proposal-1',
      name: 'propose_project_edits',
      input: {},
    })
    expect(cancelRunningTimelineEntries(third)[2]).toMatchObject({
      kind: 'propose',
      state: 'cancelled',
    })
  })

  it('keeps cancellation terminal when a tool completion arrives late', () => {
    const running = startTimelineEntry([], { id: 'read-1', name: 'read_project_text', input: {} })
    const cancelled = cancelRunningTimelineEntries(running)
    expect(
      completeTimelineEntry(cancelled, 'read-1', {
        summary: 'Late read result',
        output: 'ignored after stop',
        mutated: false,
      })[0],
    ).toMatchObject({ state: 'cancelled', label: 'Read project text' })
  })

  it('keeps repeated provider tool ids as separate entries across runs', () => {
    const first = startTimelineEntry(
      [],
      { id: 'tool-1', name: 'read_project_text', input: {} },
      'run-1',
    )
    const second = startTimelineEntry(
      first,
      { id: 'tool-1', name: 'read_project_text', input: {} },
      'run-2',
    )
    expect(second.map((entry) => entry.id)).toEqual(['run-1:tool-1', 'run-2:tool-1'])
  })

  it('truncates detail on a complete Unicode code-point boundary', () => {
    const running = startTimelineEntry([], { id: 'read', name: 'read_project_text', input: {} })
    const completed = completeTimelineEntry(running, 'read', {
      summary: 'Read',
      output: `${'x'.repeat(1_999)}😀tail`,
      mutated: false,
    })
    expect(completed[0]?.detail).toContain(`${'x'.repeat(1_999)}😀\n[detail truncated]`)
    expect(completed[0]?.detail).not.toContain('\ud83d\n')
  })
})
