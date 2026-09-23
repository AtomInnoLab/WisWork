import type { SessionNotification, SessionUpdate, ToolKind } from '@agentclientprotocol/sdk'
import type { AgentToolCall, ToolExecution } from '@wiswork/agent-core'

const READ_TOOLS = /^(get|list|read|inspect|screenshot|verify|review)_/
const SEARCH_TOOLS = /^(search|find)_/
const FETCH_TOOLS = /^(fetch|download|insert_web_image)/
const EDIT_TOOLS = /^(add|apply|build|create|duplicate|edit|insert|replace|set|update|write)_/
const DELETE_TOOLS = /^(delete|remove)_/

export function acpToolKind(name: string): ToolKind {
  if (READ_TOOLS.test(name)) return 'read'
  if (SEARCH_TOOLS.test(name)) return 'search'
  if (FETCH_TOOLS.test(name)) return 'fetch'
  if (EDIT_TOOLS.test(name)) return 'edit'
  if (DELETE_TOOLS.test(name)) return 'delete'
  return 'other'
}

export function acpToolTitle(name: string): string {
  const words = name.replace(/[_-]+/g, ' ').trim()
  return words ? words[0]!.toUpperCase() + words.slice(1) : 'Agent tool'
}

export function acpNotification(sessionId: string, update: SessionUpdate): SessionNotification {
  return { sessionId, update }
}

export function acpToolStarted(call: AgentToolCall): SessionUpdate {
  return {
    sessionUpdate: 'tool_call',
    toolCallId: call.id,
    name: call.name,
    title: acpToolTitle(call.name),
    kind: acpToolKind(call.name),
    status: 'in_progress',
  }
}

export function acpToolFinished(call: AgentToolCall, execution: ToolExecution): SessionUpdate {
  return {
    sessionUpdate: 'tool_call_update',
    toolCallId: call.id,
    status: execution.isError ? 'failed' : 'completed',
    title: execution.summary,
    content: [
      {
        type: 'content',
        content: { type: 'text', text: execution.summary },
      },
    ],
    _meta: {
      'com.wiswork/mutated': execution.mutated === true,
    },
  }
}
