/** Content-free, bounded parser input sketches, not wire captures. */
const KEYS = new Set([
  'type',
  'message',
  'id',
  'model',
  'role',
  'content',
  'usage',
  'input_tokens',
  'output_tokens',
  'cache_read_input_tokens',
  'cache_creation_input_tokens',
  'index',
  'content_block',
  'delta',
  'stop_reason',
  'stop_sequence',
  'text',
  'data',
  'name',
  'input',
  'partial_json',
  'citations',
  'caller',
  'container',
  'stop_details',
  'context_management',
  'provider',
  'cache_creation',
  'inference_geo',
  'output_tokens_details',
  'server_tool_use',
  'service_tier',
  'speed',
  'cost',
  'is_byok',
  'cost_details',
  'thinking_tokens',
  'upstream_inference_cost',
  'upstream_inference_prompt_cost',
  'upstream_inference_completions_cost',
  '__unknown',
])
const ENUMS: Record<string, readonly string[]> = {
  type: [
    'message_start',
    'message',
    'content_block_start',
    'content_block_delta',
    'content_block_stop',
    'message_delta',
    'message_stop',
    'text',
    'text_delta',
    'input_json_delta',
    'tool_use',
    'redacted_thinking',
    'thinking',
    'ping',
    'error',
    'direct',
  ],
  model: ['openai/gpt-5.6-sol'],
  role: ['assistant'],
  name: ['exec'],
  stop_reason: ['end_turn', 'stop_sequence', 'tool_use', 'max_tokens'],
}
type Sketch = null | boolean | number | string | Sketch[] | { [key: string]: Sketch }
export interface ProtocolRecording {
  readonly schema: 'wiswork-protocol-structure/v1'
  readonly contentRedacted: true
  readonly truncated: boolean
  readonly frames: readonly Sketch[]
}
export interface ProtocolFrameObserver {
  recordFrame(frame: string): void
}
export type ProtocolRecordingOutcome =
  'completed' | 'incomplete' | 'protocol_rejected' | 'interrupted' | 'not_observed'
const MAX_BYTES = 64 * 1024
const invalid = () => new Error('invalid_protocol_recording')

function sanitize(value: unknown, key = '', depth = 0, budget = { nodes: 0 }, path = ''): Sketch {
  if (++budget.nodes > 512 || depth > 12) return null
  if (key === 'input')
    return value &&
      typeof value === 'object' &&
      !Array.isArray(value) &&
      Object.keys(value).length === 0
      ? {}
      : { __unknown: null }
  if (['text', 'data', 'partial_json', 'id', 'stop_sequence'].includes(key)) {
    if (value === null) return null
    if (typeof value !== 'string') return null
    return value === '' ? '' : key === 'partial_json' ? '{' : '<redacted>'
  }
  if (value === null) return value
  if (typeof value === 'boolean') return path === 'usage.is_byok' ? value : null
  if (typeof value === 'number') {
    const integer =
      path === 'index' ||
      /^(?:message\.)?usage\.(?:input_tokens|output_tokens|cache_read_input_tokens|cache_creation_input_tokens)$/.test(
        path,
      ) ||
      path === 'usage.output_tokens_details.thinking_tokens'
    const cost =
      /^usage\.(?:cost|cost_details\.(?:upstream_inference_cost|upstream_inference_prompt_cost|upstream_inference_completions_cost))$/.test(
        path,
      )
    if (!integer && !cost) return null
    return Number.isFinite(value) &&
      (!integer || Number.isSafeInteger(value)) &&
      value >= 0 &&
      value <= 1_000_000
      ? value
      : -1
  }
  if (typeof value === 'string') {
    if (ENUMS[key]) return ENUMS[key].includes(value) ? value : 'unknown'
    if (key === 'partial_json') return value === '' ? '' : '{'
    return value === '' ? '' : '<redacted>'
  }
  if (Array.isArray(value))
    return value.slice(0, 16).map((item) => sanitize(item, '', depth + 1, budget, `${path}[]`))
  if (typeof value !== 'object' || Object.getPrototypeOf(value) !== Object.prototype) return null
  const output: Record<string, Sketch> = {}
  for (const name of Object.keys(value).slice(0, 48)) {
    if (!KEYS.has(name) || name === '__unknown') output.__unknown = null
    else
      output[name] = sanitize(
        (value as Record<string, unknown>)[name],
        name,
        depth + 1,
        budget,
        path ? `${path}.${name}` : name,
      )
  }
  return output
}

export class ProtocolRecorder implements ProtocolFrameObserver {
  #frames: Sketch[] = []
  #truncated = false
  #bytes = 0
  recordFrame(frame: string): void {
    if (this.#truncated) return
    let sketch: Sketch = 'malformed'
    try {
      const lines = frame.split(/\r\n|\r|\n/)
      const data = lines
        .filter((line) => line.startsWith('data:'))
        .map((line) => line.slice(5).trimStart())
        .join('\n')
      const event = lines
        .find((line) => line.startsWith('event:'))
        ?.slice(6)
        .trim()
      if (lines.some((line) => line !== '' && !/^(?:data:|event:|:)/.test(line))) throw invalid()
      if (!data && event === undefined) return
      if (data === '[DONE]') sketch = event === undefined || event === 'data' ? 'done' : 'malformed'
      else {
        const parsed: unknown = JSON.parse(data)
        if (event !== undefined && (parsed as { type?: unknown })?.type !== event) throw invalid()
        sketch = sanitize(parsed)
      }
    } catch {
      /* Only the fixed malformed marker survives. */
    }
    const bytes = JSON.stringify(sketch).length
    if (this.#frames.length >= 256 || this.#bytes + bytes > MAX_BYTES - 1024) {
      this.#truncated = true
      return
    }
    this.#bytes += bytes
    this.#frames.push(sketch)
  }
  snapshot(): ProtocolRecording {
    return structuredClone({
      schema: 'wiswork-protocol-structure/v1',
      contentRedacted: true,
      truncated: this.#truncated,
      frames: this.#frames,
    })
  }
}

export function parseProtocolRecording(value: unknown): ProtocolRecording {
  const serialized = JSON.stringify(value)
  if (!serialized || serialized.length > MAX_BYTES) throw invalid()
  const parsed = JSON.parse(serialized) as ProtocolRecording
  if (
    !parsed ||
    Object.keys(parsed).sort().join(',') !== 'contentRedacted,frames,schema,truncated' ||
    parsed.schema !== 'wiswork-protocol-structure/v1' ||
    parsed.contentRedacted !== true ||
    typeof parsed.truncated !== 'boolean' ||
    !Array.isArray(parsed.frames) ||
    parsed.frames.length > 256
  )
    throw invalid()
  for (const frame of parsed.frames) {
    if (frame === 'malformed' || frame === 'done') continue
    if (JSON.stringify(sanitize(frame)) !== JSON.stringify(frame)) throw invalid()
  }
  return parsed
}

export async function* recordingChunks(value: unknown): AsyncGenerator<string> {
  const recording = parseProtocolRecording(value)
  for (const frame of recording.frames)
    yield frame === 'malformed'
      ? 'data: invalid\n\n'
      : frame === 'done'
        ? 'data: [DONE]\n\n'
        : `data: ${JSON.stringify(frame)}\n\n`
}

export { replayProtocolRecording } from './index.js'
