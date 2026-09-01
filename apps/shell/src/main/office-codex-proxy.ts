import { createHash } from 'node:crypto'
import {
  createToolExecutionSuspensionAuthority,
  type AgentToolDef,
  type ToolExecution,
} from '@wiswork/agent-core'
import {
  createDocumentToolManifest,
  createDocumentToolSession,
  type ToolMutability,
} from '@wiswork/codex-bridge'
import type { MessagesProxyResponse, OfficeEnhancedSessionStatement } from '@wiswork/office-bridge'
import type { EnhancedRolloutPolicy } from '@wiswork/agent-runtime'
import type { ShellCodexRuntime } from './codex-runtime'
import type { OfficeRelayToolCall, OfficeRelayToolResult } from './office-relay-client'

const MAX_BODY_BYTES = 256 * 1024
const MAX_TEXT_BYTES = 128 * 1024
const MAX_TOOLS = 64
const READ_TOOLS = new Set([
  'get_document_text',
  'get_document_structure',
  'get_ooxml',
  'screenshot_document',
  'get_cell_ranges',
  'get_range_as_csv',
  'search_data',
  'screenshot_range',
  'get_all_objects',
  'inspect_slide_masters',
  'screenshot_slide',
  'list_slide_shapes',
  'read_slide_text',
  'verify_slides',
])
const MUTATION_TOOLS = new Set([
  'write_document',
  'execute_office_js',
  'set_cell_range',
  'clear_cell_range',
  'copy_to',
  'modify_sheet_structure',
  'modify_workbook_structure',
  'resize_range',
  'modify_object',
  'eval_officejs',
  'edit_slide_text',
  'edit_slide_xml',
  'edit_slide_chart',
  'edit_slide_master',
  'edit_slide_master_xml',
  'duplicate_slide',
  'propose_raw_office_edit',
])

interface PolicyAuthority {
  issue(value: {
    generation: number
    host: OfficeEnhancedSessionStatement['host']
    policy: unknown
    capabilities: unknown
  }): unknown
  consume(grant: unknown): any
}

const hostName = (value: 'Word' | 'Excel' | 'PowerPoint') =>
  ({ Word: 'office-word', Excel: 'office-excel', PowerPoint: 'office-powerpoint' })[
    value
  ] as OfficeEnhancedSessionStatement['host']

function parseRequest(
  body: unknown,
  rawOffice: boolean,
): {
  text: string
  tools: AgentToolDef[]
  policy: Record<string, ToolMutability>
} {
  if (
    !body ||
    typeof body !== 'object' ||
    Array.isArray(body) ||
    Buffer.byteLength(JSON.stringify(body)) > MAX_BODY_BYTES
  )
    throw new Error('enhanced_request_invalid')
  const value = body as Record<string, unknown>
  if (
    typeof value.system !== 'string' ||
    !Array.isArray(value.messages) ||
    !Array.isArray(value.tools) ||
    value.tools.length > MAX_TOOLS
  )
    throw new Error('enhanced_request_invalid')
  const text = JSON.stringify({ instructions: value.system, messages: value.messages })
  if (Buffer.byteLength(text) > MAX_TEXT_BYTES || Buffer.byteLength(value.system) > MAX_TEXT_BYTES)
    throw new Error('enhanced_request_invalid')
  const tools: AgentToolDef[] = []
  const policy: Record<string, ToolMutability> = {}
  for (const candidate of value.tools) {
    if (!candidate || typeof candidate !== 'object' || Array.isArray(candidate))
      throw new Error('enhanced_request_invalid')
    const tool = candidate as Record<string, unknown>
    const name = tool.name
    if (
      typeof name !== 'string' ||
      typeof tool.description !== 'string' ||
      !tool.input_schema ||
      typeof tool.input_schema !== 'object' ||
      Array.isArray(tool.input_schema)
    )
      throw new Error('enhanced_request_invalid')
    const mutability = READ_TOOLS.has(name)
      ? 'read'
      : MUTATION_TOOLS.has(name)
        ? 'mutate'
        : undefined
    const raw =
      name === 'execute_office_js' || name === 'eval_officejs' || name === 'propose_raw_office_edit'
    if (!mutability || (raw && !rawOffice)) continue
    tools.push({
      name,
      description: tool.description,
      inputSchema: tool.input_schema as Record<string, unknown>,
    })
    policy[name] = mutability
  }
  return { text, tools, policy }
}

export function createOfficeCodexProxy(options: {
  runtime: ShellCodexRuntime
  rollout: EnhancedRolloutPolicy
  policyAuthority: PolicyAuthority
}) {
  return async (request: {
    body: unknown
    signal: AbortSignal
    host: 'Word' | 'Excel' | 'PowerPoint'
    sessionId: string
    requestId: string
    statement: Readonly<OfficeEnhancedSessionStatement>
    executeTool(call: OfficeRelayToolCall): Promise<OfficeRelayToolResult>
  }): Promise<MessagesProxyResponse> => {
    const host = hostName(request.host)
    if (request.statement.host !== host || request.statement.expires_at <= Date.now())
      throw new Error('enhanced_session_stale')
    const parsed = parseRequest(request.body, request.statement.raw_office)
    const capabilities = [
      'semantic-read',
      'transaction-proposal',
      'bounded-render-facts',
      ...(request.statement.raw_office ? ['raw-office-proposal'] : []),
    ]
    const grant = options.policyAuthority.issue({
      generation: request.statement.policy_generation,
      host,
      policy: options.rollout,
      capabilities,
    })
    const manifest = createDocumentToolManifest({
      policyGrant: grant,
      consumePolicyGrant: options.policyAuthority.consume,
      tools: parsed.tools,
      policy: parsed.policy,
    })
    const suspension = createToolExecutionSuspensionAuthority()
    let open = true
    const turnId = `turn_${createHash('sha256').update(`${request.sessionId}:${request.requestId}`).digest('base64url').slice(0, 32)}`
    const execute = async (
      call: { id: string; name: string; input: Record<string, unknown> },
      signal?: AbortSignal,
    ): Promise<ToolExecution> => {
      if (signal?.aborted)
        return { output: 'tool_cancelled', isError: true, summary: 'Tool cancelled' }
      const result = await request.executeTool({
        turnId,
        callId: call.id,
        generation: request.statement.session_generation,
        toolName: call.name,
        input: call.input,
      })
      return {
        output: result.output,
        isError: result.isError,
        summary: result.isError ? 'Office tool failed' : 'Office tool complete',
        mutated: parsed.policy[call.name] === 'mutate' && !result.isError,
      }
    }
    const documentId = `office_${createHash('sha256').update(`${request.statement.runtime_instance}:${request.sessionId}`).digest('base64url').slice(0, 32)}`
    const session = createDocumentToolSession({
      identity: {
        ownerId: request.sessionId,
        host,
        documentId,
        sessionId: request.sessionId,
        generation: request.statement.policy_generation,
      },
      manifest,
      isOpen: () => open && !request.signal.aborted,
      executeRead: execute,
      suspendMutation: suspension.suspend,
      ownsSuspension: suspension.owns,
    })
    const encoder = new TextEncoder()
    const queue: Uint8Array[] = []
    let wake: (() => void) | undefined
    let terminal = false
    let failure: Error | undefined
    const push = (value: string) => {
      queue.push(encoder.encode(value))
      wake?.()
      wake = undefined
    }
    const pump = setInterval(() => {
      const pending = session.mutationAuthority.claimNext()
      if (!pending) return
      void execute(pending.request.call).then(
        (result) => session.mutationAuthority.settle(pending.claim, result),
        () => session.mutationAuthority.reject(pending.claim, 'tool_execution_failed'),
      )
    }, 5)
    pump.unref()
    void options.runtime
      .runOfficeTurn({
        documentId,
        host,
        generation: request.statement.policy_generation,
        text: parsed.text,
        toolSession: session,
        signal: request.signal,
        onEvent(event) {
          if (event.type === 'text')
            push(
              `data: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: { type: 'text_delta', text: event.text } })}\n\n`,
            )
          if (event.type === 'terminal') {
            terminal = true
            wake?.()
            wake = undefined
          }
        },
      })
      .catch((error) => {
        failure = error instanceof Error ? error : new Error('enhanced_turn_failed')
        terminal = true
        wake?.()
        wake = undefined
      })
      .finally(() => {
        clearInterval(pump)
        open = false
        session.close()
      })
    return {
      status: 200,
      contentType: 'text/event-stream',
      body: {
        async *[Symbol.asyncIterator]() {
          while (!terminal || queue.length) {
            if (queue.length) {
              yield queue.shift()!
              continue
            }
            await new Promise<void>((resolve) => {
              wake = resolve
            })
          }
          if (failure) throw failure
          yield encoder.encode(
            `data: ${JSON.stringify({ type: 'message_delta', delta: { stop_reason: 'end_turn' } })}\n\ndata: [DONE]\n\n`,
          )
        },
      },
    }
  }
}
