import type { AgentMessage, AgentToolCall } from './types.js'

/** A bounded text checkpoint must end after a complete tool batch. */
export function parseAgentResumeMessages(value: unknown): AgentMessage[] | undefined {
  try {
    const raw = JSON.stringify(value)
    if (!raw || new TextEncoder().encode(raw).byteLength > 128 * 1024) return undefined
    const messages = JSON.parse(raw) as AgentMessage[]
    if (
      !Array.isArray(messages) ||
      messages.length < 3 ||
      messages.length > 128 ||
      messages[0]?.role !== 'user' ||
      messages.at(-1)?.role !== 'tool'
    )
      return undefined
    const object = (item: unknown): item is Record<string, unknown> =>
      !!item && typeof item === 'object' && !Array.isArray(item)
    const keys = (item: Record<string, unknown>, allowed: string[]) =>
      Object.keys(item).every((key) => allowed.includes(key))
    const id = (item: unknown) => typeof item === 'string' && item.length > 0 && item.length <= 256
    const boundedInput = (input: unknown, depth = 0): boolean =>
      depth <= 64 &&
      (!input ||
        typeof input !== 'object' ||
        Object.values(input).every((child) => boundedInput(child, depth + 1)))
    let pending: AgentToolCall[] | undefined
    const seen = new Set<string>()
    for (const [index, message] of messages.entries()) {
      if (!object(message)) return undefined
      if (message.role === 'tool') {
        if (
          !keys(message, ['role', 'results']) ||
          !pending ||
          !Array.isArray(message.results) ||
          message.results.length !== pending.length
        )
          return undefined
        for (const [position, result] of message.results.entries()) {
          const call = pending[position]!
          if (
            !object(result) ||
            !keys(result, ['id', 'name', 'output', 'isError']) ||
            result.id !== call.id ||
            result.name !== call.name ||
            typeof result.output !== 'string' ||
            (result.isError !== undefined && typeof result.isError !== 'boolean')
          )
            return undefined
        }
        pending = undefined
      } else if (message.role === 'user') {
        if (
          pending ||
          !keys(message, ['role', 'text']) ||
          typeof message.text !== 'string' ||
          !message.text.trim() ||
          (index > 0 && !['assistant', 'tool'].includes(messages[index - 1]!.role))
        )
          return undefined
      } else if (message.role === 'assistant') {
        if (
          pending ||
          !keys(message, ['role', 'text', 'toolCalls']) ||
          typeof message.text !== 'string' ||
          (index > 0 && messages[index - 1]?.role === 'assistant')
        )
          return undefined
        if (message.toolCalls !== undefined) {
          if (
            !Array.isArray(message.toolCalls) ||
            !message.toolCalls.length ||
            message.toolCalls.length > 32
          )
            return undefined
          for (const call of message.toolCalls) {
            if (
              !object(call) ||
              !keys(call, ['id', 'name', 'input']) ||
              !id(call.id) ||
              !id(call.name) ||
              !object(call.input) ||
              !boundedInput(call.input) ||
              seen.has(call.id)
            )
              return undefined
            seen.add(call.id)
          }
          pending = message.toolCalls
        } else if (!message.text.trim()) return undefined
      } else return undefined
    }
    return pending ? undefined : messages
  } catch {
    return undefined
  }
}
