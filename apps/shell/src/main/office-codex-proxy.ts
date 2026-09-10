import { createHash } from 'node:crypto'
import {
  createToolExecutionSuspensionAuthority,
  decodeOfficeScreenshotResult,
  officeScreenshotBytes,
  OFFICE_SCREENSHOT_PREVIEW_BYTES,
  type AgentToolDef,
  type AgentToolCall,
  type ToolExecution,
  type ToolDisplay,
} from '@wiswork/agent-core'
import {
  compiledDocumentTool,
  createDocumentToolManifest,
  createDocumentToolSession,
  type ToolMutability,
} from '@wiswork/codex-bridge'
import type { MessagesProxyResponse, OfficeEnhancedSessionStatement } from '@wiswork/office-bridge'
import type {
  EnhancedRolloutPolicy,
  EnhancedTelemetry,
  PcHostProposalSummary,
} from '@wiswork/agent-runtime'
import type { ShellCodexRuntime } from './codex-runtime'
import type { OfficeRelayToolCall, OfficeRelayToolResult } from './office-relay-client'
import type { OfficeRetrievalProxy, OfficeWebCapability } from './office-retrieval-proxy'

const MAX_BODY_BYTES = 256 * 1024
const MAX_TEXT_BYTES = 128 * 1024
const MAX_TOOLS = 64
const MAX_RETRIEVAL_DISPLAY_BYTES = 8 * 1024
const MAX_IMAGE_HANDOFF_BYTES = 180 * 1024
const OFFICE_REMOTE_MUTATION_MS = 5 * 60_000
const PRIVATE_IMAGE_FIELD = '_wiswork_image_base64'
const RETRIEVAL_TOOLS = new Set(['web_search', 'web_fetch', 'image_search'])
const IMAGE_PREFETCH_ERRORS = new Set([
  'image_fetch_unavailable',
  'image_limit',
  'invalid_image',
  'image_mime_unsupported',
  'invalid_tool_input',
  'cancelled',
])

/** Project public source links only; model output and upstream error bodies stay on PC. */
function retrievalDisplay(
  output: string,
  toolName: string,
): { result_count?: number; display?: ToolDisplay } {
  try {
    const value = JSON.parse(output)
    const entries: unknown[] = Array.isArray(value?.images)
      ? value.images
      : Array.isArray(value?.results)
        ? value.results
        : toolName === 'web_fetch'
          ? [value]
          : []
    const items: NonNullable<ToolDisplay['items']> = []
    for (const entry of entries.slice(0, 20)) {
      if (!entry || typeof entry !== 'object') continue
      const record = entry as Record<string, unknown>
      const raw = record.source_url ?? record.url
      if (typeof raw !== 'string' || raw.length > 2048) continue
      let url: URL
      try {
        url = new URL(raw)
      } catch {
        continue
      }
      if (
        url.protocol !== 'https:' ||
        url.username ||
        url.password ||
        url.port ||
        !url.hostname.includes('.') ||
        /^[\d.]+$/.test(url.hostname) ||
        url.hostname.includes(':') ||
        /\.(?:localhost|local)$/.test(url.hostname)
      )
        continue
      // Query/fragment may contain signed credentials; source cards do not need them.
      url.search = ''
      url.hash = ''
      const item = {
        url: url.href,
        ...(typeof record.title === 'string' ? { title: record.title.slice(0, 160) } : {}),
      }
      if (Buffer.byteLength(JSON.stringify([...items, item])) > MAX_RETRIEVAL_DISPLAY_BYTES) break
      items.push(item)
      if (items.length === 8) break
    }
    return {
      ...(toolName === 'web_fetch' ? {} : { result_count: Math.min(entries.length, 20) }),
      ...(items.length
        ? {
            display: {
              kind: toolName === 'image_search' ? 'images' : 'links',
              items,
            } as ToolDisplay,
          }
        : {}),
    }
  } catch {
    return {}
  }
}
export const OFFICE_PROXY_KEEPALIVE_MS = 15_000
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

const summarizeOfficeProposal = (
  host: OfficeEnhancedSessionStatement['host'],
  call: { readonly name: string; readonly input: Record<string, unknown> },
): PcHostProposalSummary => {
  if (Object.hasOwn(call.input, PRIVATE_IMAGE_FIELD)) throw new Error('invalid_tool_input')
  const program = call.input.program
  const count = (
    Object.values(call.input).find(Array.isArray) ??
    (program && typeof program === 'object' && !Array.isArray(program)
      ? Object.values(program).find(Array.isArray)
      : undefined)
  )?.length
  const operation =
    call.name === 'clear_cell_range'
      ? 'delete'
      : call.name === 'duplicate_slide' || call.name.includes('structure')
        ? 'restructure'
        : call.name.includes('style') ||
            call.name.includes('master') ||
            call.name === 'resize_range'
          ? 'format'
          : 'replace'
  return {
    operation,
    target: host === 'office-word' ? 'document' : host === 'office-excel' ? 'cells' : 'slides',
    scope: 'bounded-set',
    ...(count ? { count } : {}),
  }
}

function parseRequest(
  body: unknown,
  host: OfficeEnhancedSessionStatement['host'],
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
    const compiled = compiledDocumentTool(host, name)
    if (!compiled) continue
    const [mutability, requiredCapability] = compiled
    if (requiredCapability === 'raw-office-proposal' && !rawOffice) continue
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
  telemetry?: EnhancedTelemetry
  prepareImageHandoff?: (image: {
    bytes: Uint8Array
    mime: 'image/png' | 'image/jpeg'
  }) => Promise<{ bytes: Uint8Array; mime: 'image/png' | 'image/jpeg' }>
}) {
  return async (request: {
    body: unknown
    signal: AbortSignal
    host: 'Word' | 'Excel' | 'PowerPoint'
    sessionId: string
    requestId: string
    statement: Readonly<OfficeEnhancedSessionStatement>
    executeTool(call: OfficeRelayToolCall): Promise<OfficeRelayToolResult>
    executeRetrieval?: OfficeRetrievalProxy
  }): Promise<MessagesProxyResponse> => {
    const host = hostName(request.host)
    const telemetry = (
      phase: 'plan' | 'dispatch' | 'verify' | 'complete' | 'pending',
      outcome: 'started' | 'succeeded' | 'failed' | 'verified' | 'applied_unverified',
    ) => {
      try {
        options.telemetry?.host(host, phase, outcome)
      } catch {
        // Aggregate telemetry must never own an Office turn.
      }
    }
    if (request.statement.host !== host || request.statement.expires_at <= Date.now())
      throw new Error('enhanced_session_stale')
    telemetry('plan', 'started')
    const parsed = parseRequest(request.body, host, request.statement.raw_office)
    const pcRetrieval = new Set(['web_search', 'image_search', 'insert_web_image'])
    if (!request.executeRetrieval) {
      parsed.tools = parsed.tools.filter((tool) => !pcRetrieval.has(tool.name))
      for (const name of pcRetrieval) delete parsed.policy[name]
    }
    telemetry('plan', 'succeeded')
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
    let acceptingEvents = true
    const turnId = `turn_${createHash('sha256').update(`${request.sessionId}:${request.requestId}`).digest('base64url').slice(0, 32)}`
    const wireCallId = (identity: string) =>
      `call_${createHash('sha256')
        .update(JSON.stringify([turnId, identity]))
        .digest('base64url')}`
    const activities = new Map<
      string,
      {
        callId: string
        toolName: string
        startedAt: number
        settled: boolean
        details?: Record<string, unknown>
      }
    >()
    const activity = (
      callId: string,
      toolName: string,
      state: 'running' | 'complete' | 'error',
    ) => {
      if (!open || terminal || !Object.hasOwn(parsed.policy, toolName)) return
      let item = activities.get(callId)
      if (state === 'running') {
        if (request.signal.aborted || item || activities.size >= 1024) return
        item = { callId, toolName, startedAt: Date.now(), settled: false }
        activities.set(callId, item)
      } else {
        if (!item || item.settled || item.toolName !== toolName) return
        item.settled = true
      }
      push(
        `data: ${JSON.stringify({
          // Older Taskpanes understand retrieval activity and ignore unknown event kinds.
          type: RETRIEVAL_TOOLS.has(toolName) ? 'wiswork_tool_activity' : 'wiswork_tool_lifecycle',
          generation: request.statement.session_generation,
          call_id: item.callId,
          tool_name: toolName,
          state,
          started_at: item.startedAt,
          ...(state === 'running'
            ? {}
            : {
                summary: state === 'error' ? 'Office tool failed' : 'Office tool complete',
                ...(state === 'complete'
                  ? item.details
                  : item.details?.summary === 'Retrieval unavailable'
                    ? { summary: 'Retrieval unavailable' }
                    : toolName === 'insert_web_image' &&
                        typeof item.details?.summary === 'string' &&
                        IMAGE_PREFETCH_ERRORS.has(item.details.summary)
                      ? { summary: item.details.summary }
                      : {}),
              }),
        })}\n\n`,
      )
      // Retain a bounded replay tombstone, not potentially large preview metadata.
      if (item.settled) item.details = undefined
    }
    const closeActivities = () => {
      for (const [callId, item] of activities) activity(callId, item.toolName, 'error')
    }
    const execute = async (call: AgentToolCall, signal?: AbortSignal): Promise<ToolExecution> => {
      if (signal?.aborted)
        return { output: 'tool_cancelled', isError: true, summary: 'Tool cancelled' }
      const mutation = parsed.policy[call.name] === 'mutate'
      if (mutation) telemetry('pending', 'started')
      telemetry('dispatch', 'started')
      let result: OfficeRelayToolResult
      const retrievalCapability: Partial<Record<string, OfficeWebCapability>> = {
        web_search: 'web-search.v1',
        web_fetch: 'web-fetch.v1',
        image_search: 'image-search.v1',
      }
      const capability = retrievalCapability[call.name]
      const callId = wireCallId(call.invocationId ?? call.id)
      const enrich = (details: Record<string, unknown>) => {
        const item = activities.get(callId)
        if (item && !item.settled && open && !request.signal.aborted)
          item.details = {
            ...(typeof call.input.query === 'string'
              ? { query: call.input.query.slice(0, 240) }
              : {}),
            ...details,
          }
      }
      const retrievalSignal = signal ?? request.signal
      let dispatched = false
      try {
        if (Object.hasOwn(call.input, PRIVATE_IMAGE_FIELD)) throw new Error('invalid_tool_input')
        if (capability && request.executeRetrieval) {
          const output = await request.executeRetrieval(
            capability,
            call.input,
            signal ?? request.signal,
          )
          result = {
            output: new TextDecoder('utf-8', { fatal: true }).decode(output),
            isError: false,
          }
          if (!retrievalSignal.aborted)
            enrich({
              summary: 'Retrieval complete',
              ...retrievalDisplay(result.output, call.name),
            })
        } else {
          let toolInput = call.input
          if (call.name === 'insert_web_image') {
            if (!request.executeRetrieval) throw new Error('image_fetch_unavailable')
            // Stay within the active agent request: the PC retrieval proxy owns
            // image-search provenance, URL/DNS/redirect validation and decoding.
            const payload = await request.executeRetrieval(
              'image-fetch.v1',
              { url: call.input.url },
              retrievalSignal,
            )
            if (payload.byteLength > 3 * 1024 * 1024) throw new Error('image_limit')
            const image = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(payload))
            if (
              !image ||
              (image.mime !== 'image/png' && image.mime !== 'image/jpeg') ||
              typeof image.data_base64 !== 'string' ||
              !image.data_base64.length ||
              image.data_base64.length % 4 !== 0 ||
              !/^[A-Za-z0-9+/]+={0,2}$/.test(image.data_base64)
            )
              throw new Error('invalid_image')
            let bytes: Uint8Array = Buffer.from(image.data_base64, 'base64')
            if (bytes.byteLength > 2 * 1024 * 1024) throw new Error('image_limit')
            if (options.prepareImageHandoff)
              ({ bytes } = await options.prepareImageHandoff({ bytes, mime: image.mime }))
            if (!bytes.byteLength || bytes.byteLength > MAX_IMAGE_HANDOFF_BYTES)
              throw new Error('image_limit')
            toolInput = {
              ...call.input,
              [PRIVATE_IMAGE_FIELD]: Buffer.from(bytes).toString('base64'),
            }
            if (Buffer.byteLength(JSON.stringify(toolInput)) > MAX_BODY_BYTES)
              throw new Error('image_limit')
            if (retrievalSignal.aborted || request.signal.aborted) throw new Error('cancelled')
          }
          dispatched = true
          result = await request.executeTool({
            turnId,
            // Model carrier IDs may be shorter than Relay identifiers. Keep the
            // protocol boundary deterministic without weakening Relay validation.
            callId,
            generation: request.statement.session_generation,
            toolName: call.name,
            input: toolInput,
          })
        }
      } catch (error) {
        telemetry('dispatch', 'failed')
        if (call.name === 'insert_web_image' && !dispatched) {
          const code =
            retrievalSignal.aborted || request.signal.aborted
              ? 'cancelled'
              : error instanceof Error
                ? error.message
                : ''
          const output = IMAGE_PREFETCH_ERRORS.has(code) ? code : 'image_fetch_unavailable'
          // Reuse the bounded lifecycle summary, never upstream text or image URLs.
          enrich({ summary: output })
          return {
            output,
            isError: true,
            summary: 'Office image unavailable',
            mutated: false,
          }
        }
        if (capability) {
          if (request.executeRetrieval && !retrievalSignal.aborted)
            enrich({ summary: 'Retrieval unavailable' })
          return {
            output:
              error instanceof Error && error.message === 'retrieval_cancelled'
                ? 'retrieval_cancelled'
                : 'retrieval_upstream_error',
            isError: true,
            summary: 'Office retrieval unavailable',
            mutated: false,
          }
        }
        throw error
      }
      telemetry('dispatch', result.isError ? 'failed' : 'succeeded')
      if (call.name === 'screenshot_slide' && !result.isError) {
        if (signal?.aborted || request.signal.aborted)
          return {
            output: 'tool_cancelled',
            isError: true,
            summary: 'Tool cancelled',
            mutated: false,
          }
        try {
          const screenshot = decodeOfficeScreenshotResult(result.output)
          if (!options.prepareImageHandoff) throw new Error('office_screenshot_unavailable')
          const image = screenshot.modelContent[0]!.image
          const bytes = officeScreenshotBytes(image, OFFICE_SCREENSHOT_PREVIEW_BYTES)
          const verified = await options.prepareImageHandoff({
            bytes,
            mime: image.mime as 'image/png' | 'image/jpeg',
          })
          if (signal?.aborted || request.signal.aborted)
            return {
              output: 'tool_cancelled',
              isError: true,
              summary: 'Tool cancelled',
              mutated: false,
            }
          if (
            verified.mime !== image.mime ||
            !Buffer.from(bytes).equals(Buffer.from(verified.bytes))
          )
            throw new Error('office_screenshot_unavailable')
          screenshot.output = JSON.stringify({
            ...JSON.parse(screenshot.output),
            visualAvailableToModel: true,
          })
          telemetry('verify', 'verified')
          return {
            ...screenshot,
            isError: false,
            summary: 'Office screenshot delivered',
            mutated: false,
          }
        } catch {
          telemetry('verify', 'failed')
          return {
            output: 'office_screenshot_unavailable',
            isError: true,
            summary: 'Office screenshot unavailable',
            mutated: false,
          }
        }
      }
      telemetry('verify', result.isError ? 'failed' : mutation ? 'applied_unverified' : 'verified')
      let appliedBeforeFailure = false
      if (mutation && result.isError) {
        try {
          const decision = JSON.parse(result.output)
          appliedBeforeFailure =
            decision?.status === 'failed' && decision.error === 'office_verify_failed'
        } catch {
          // Unknown failures are not evidence of a completed write.
        }
      }
      return {
        output: result.output,
        isError: result.isError,
        summary: result.isError ? 'Office tool failed' : 'Office tool complete',
        mutated: mutation && (!result.isError || appliedBeforeFailure),
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
      // One Relay call spans Taskpane consent and the subsequent Office write.
      maxMutationMs: OFFICE_REMOTE_MUTATION_MS,
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
    request.signal.addEventListener('abort', closeActivities, { once: true })
    const pump = setInterval(() => {
      const pending = session.mutationAuthority.claimNext()
      if (!pending) return
      void execute(pending.request.call, pending.request.signal)
        .then(
          (result) => {
            if (!pending.request.signal.aborted)
              session.mutationAuthority.settle(pending.claim, result)
          },
          () => {
            if (!pending.request.signal.aborted)
              session.mutationAuthority.reject(pending.claim, 'tool_execution_failed')
          },
        )
        .catch(() => {
          // Invalid receipts are settled as tool errors before the router throws.
          // Cancellation/timeouts already consume their claim and abort its signal.
          telemetry('dispatch', 'failed')
        })
    }, 5)
    pump.unref()
    const keepalive = setInterval(() => push(': keepalive\n\n'), OFFICE_PROXY_KEEPALIVE_MS)
    keepalive.unref()
    void options.runtime
      .runOfficeTurn({
        documentId,
        host,
        generation: request.statement.policy_generation,
        text: parsed.text,
        toolSession: session,
        summarizeProposal: (call) => summarizeOfficeProposal(host, call),
        signal: request.signal,
        onEvent(event) {
          if (!open || terminal || !acceptingEvents) return
          if (event.type === 'tool-start' || event.type === 'tool-complete') {
            const identity = event.turnId ? `${event.turnId}:${event.callId}` : event.callId
            activity(
              wireCallId(identity),
              event.toolName,
              event.type === 'tool-start'
                ? 'running'
                : event.isError || request.signal.aborted
                  ? 'error'
                  : 'complete',
            )
          }
          if (event.type === 'text' && !request.signal.aborted)
            push(
              `data: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: { type: 'text_delta', text: event.text } })}\n\n`,
            )
          if (event.type === 'terminal') {
            closeActivities()
            acceptingEvents = false
            telemetry('complete', event.status === 'completed' ? 'succeeded' : 'failed')
            if (event.status !== 'completed') failure = new Error('enhanced_turn_failed')
          }
        },
      })
      .catch((error) => {
        failure = error instanceof Error ? error : new Error('enhanced_turn_failed')
      })
      .finally(() => {
        clearInterval(pump)
        clearInterval(keepalive)
        closeActivities()
        request.signal.removeEventListener('abort', closeActivities)
        open = false
        session.close()
        terminal = true
        wake?.()
        wake = undefined
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
