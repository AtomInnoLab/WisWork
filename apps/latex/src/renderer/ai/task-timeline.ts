import type { AgentToolCall, ToolExecution } from '@wiswork/agent-core'

export type TaskTimelineKind =
  'context' | 'read' | 'search' | 'compile' | 'propose' | 'verify' | 'result' | 'error'

export type TaskTimelineState = 'running' | 'success' | 'error' | 'cancelled'

export interface TaskTimelineEntry {
  id: string
  kind: TaskTimelineKind
  label: string
  state: TaskTimelineState
  detail?: string
}

const MAX_DETAIL_CHARS = 2_000

function boundedDetail(detail: string | undefined): string | undefined {
  if (!detail) return undefined
  let codePoints = 0
  let end = 0
  while (end < detail.length && codePoints < MAX_DETAIL_CHARS) {
    const codePoint = detail.codePointAt(end)
    end += codePoint !== undefined && codePoint > 0xffff ? 2 : 1
    codePoints += 1
  }
  return end < detail.length ? `${detail.slice(0, end)}\n[detail truncated]` : detail
}

function kindForTool(name: string): TaskTimelineKind {
  if (name === 'search_project_text') return 'search'
  if (name === 'compile_project') return 'compile'
  if (name === 'propose_project_edits') return 'propose'
  if (name.includes('verify')) return 'verify'
  return 'read'
}

function toolLabel(name: string): string {
  const labels: Record<string, string> = {
    list_project_files: 'List project files',
    search_project_text: 'Search project text',
    read_project_text: 'Read project text',
    get_compile_diagnostics: 'Read compile diagnostics',
    compile_project: 'Compile project',
    propose_project_edits: 'Prepare edit proposal',
  }
  return labels[name] ?? name.replaceAll('_', ' ')
}

function inputDetail(input: Record<string, unknown>): string | undefined {
  try {
    const value = JSON.stringify(input)
    return value === '{}' ? undefined : boundedDetail(value)
  } catch {
    return undefined
  }
}

export function startTimelineEntry(
  entries: readonly TaskTimelineEntry[],
  call: AgentToolCall,
  runId?: string,
): TaskTimelineEntry[] {
  const entryId = runId ? `${runId}:${call.id}` : call.id
  return [
    ...entries.filter((entry) => entry.id !== entryId),
    {
      id: entryId,
      kind: kindForTool(call.name),
      label: toolLabel(call.name),
      state: 'running',
      ...(inputDetail(call.input) ? { detail: inputDetail(call.input) } : {}),
    },
  ]
}

export function completeTimelineEntry(
  entries: readonly TaskTimelineEntry[],
  callId: string,
  execution: ToolExecution,
): TaskTimelineEntry[] {
  return entries.map((entry) =>
    entry.id === callId && entry.state === 'running'
      ? {
          ...entry,
          label: execution.summary || entry.label,
          state: execution.isError ? 'error' : 'success',
          ...(boundedDetail(execution.output) ? { detail: boundedDetail(execution.output) } : {}),
        }
      : entry,
  )
}

export function cancelRunningTimelineEntries(
  entries: readonly TaskTimelineEntry[],
): TaskTimelineEntry[] {
  return entries.map((entry) =>
    entry.state === 'running' ? { ...entry, state: 'cancelled' as const } : entry,
  )
}

export function failRunningTimelineEntries(
  entries: readonly TaskTimelineEntry[],
  detail: string,
): TaskTimelineEntry[] {
  return entries.map((entry) =>
    entry.state === 'running'
      ? { ...entry, state: 'error' as const, detail: boundedDetail(detail) }
      : entry,
  )
}
