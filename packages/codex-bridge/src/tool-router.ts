import { createHash, randomBytes, timingSafeEqual } from 'node:crypto'
import type {
  AgentToolCall,
  AgentToolDef,
  ToolExecution,
  ToolExecutionOutcome,
  ToolExecutionSuspension,
} from '@wiswork/agent-core'
import type { EnhancedHost, EnhancedPolicySnapshot } from '@wiswork/agent-runtime'
import {
  ENHANCED_HOSTS,
  parseEnhancedCapabilities,
  parseEnhancedRolloutPolicy,
} from '@wiswork/agent-runtime'
import type { DocumentCarrierHandle, DocumentCarrierIssuer } from './types.js'

const SECRET_BYTES = 32
const MAX_TOOLS = 64
const MAX_NAME_BYTES = 128
const MAX_DESCRIPTION_BYTES = 128_000
const MAX_TOTAL_DESCRIPTION_BYTES = 512_000
const MAX_SCHEMA_BYTES = 512_000
const MAX_TOTAL_SCHEMA_BYTES = 2_000_000
const MAX_LIST_BYTES = 3_000_000
const MAX_INPUT_BYTES = 1_000_000
const MAX_OUTPUT_BYTES = 1_000_000
const MAX_SUMMARY_BYTES = 4_096
const MAX_ID_BYTES = 256
const MAX_GRAPH_NODES = 20_000
const MAX_TOTAL_GRAPH_NODES = 100_000
const MAX_GRAPH_DEPTH = 48
const MAX_CALL_MS = 30_000
const MAX_TOTAL_CALLS = 1_024
const MAX_PENDING_MUTATIONS = 8
const SECRET_PATTERN = /^[A-Za-z0-9_-]{43}$/

export type ToolMutability = 'read' | 'mutate'
type RequiredCapability =
  'semantic-read' | 'transaction-proposal' | 'bounded-render-facts' | 'raw-office-proposal'

const CATALOG = Object.freeze({
  latex: Object.freeze({
    list_project_files: ['read', 'semantic-read'],
    search_project_text: ['read', 'semantic-read'],
    read_project_text: ['read', 'semantic-read'],
    get_compile_diagnostics: ['read', 'semantic-read'],
    compile_project: ['mutate', 'transaction-proposal'],
    propose_project_edits: ['mutate', 'transaction-proposal'],
  }),
  docs: Object.freeze({
    get_document_context: ['read', 'semantic-read'],
    read_blocks: ['read', 'semantic-read'],
    insert_content: ['mutate', 'transaction-proposal'],
    replace_blocks: ['mutate', 'transaction-proposal'],
    apply_commands: ['mutate', 'transaction-proposal'],
    insert_image: ['mutate', 'transaction-proposal'],
    insert_chart: ['mutate', 'transaction-proposal'],
    edit_chart: ['mutate', 'transaction-proposal'],
  }),
  sheets: Object.freeze({
    get_workbook_context: ['read', 'semantic-read'],
    read_range: ['read', 'semantic-read'],
    load_guide: ['read', 'semantic-read'],
    read_formats: ['read', 'semantic-read'],
    read_sheet_features: ['read', 'semantic-read'],
    read_cells: ['read', 'semantic-read'],
    propose_operations: ['mutate', 'transaction-proposal'],
  }),
  slides: Object.freeze(
    Object.fromEntries([
      ...[
        'get_deck_context',
        'read_slide',
        'ask_clarification',
        'plan_deck',
        'list_style_templates',
      ].map((name) => [name, ['read', 'semantic-read']]),
      ...[
        'build_deck',
        'set_element_text',
        'set_element_style',
        'set_element_transform',
        'execute_slide_script',
        'set_element_fill',
        'set_element_stroke',
        'crop_image',
        'set_picture_opacity',
        'replace_image',
        'delete_slide',
        'save_style_template',
        'add_slide',
        'add_text_box',
        'add_shape',
        'add_chart',
        'add_smartart',
        'add_table',
        'edit_table_cell',
        'edit_table_structure',
        'edit_table_style',
        'edit_chart',
        'set_slide_background',
        'set_speaker_notes',
        'delete_element',
        'ungroup_element',
      ].map((name) => [name, ['mutate', 'transaction-proposal']]),
    ]) as Record<string, readonly [ToolMutability, RequiredCapability]>,
  ),
  'office-word': Object.freeze({
    get_document_text: ['read', 'semantic-read'],
    get_document_structure: ['read', 'semantic-read'],
    get_ooxml: ['read', 'semantic-read'],
    screenshot_document: ['read', 'bounded-render-facts'],
    write_document: ['mutate', 'transaction-proposal'],
    execute_office_js: ['mutate', 'raw-office-proposal'],
    propose_raw_office_edit: ['mutate', 'raw-office-proposal'],
  }),
  'office-excel': Object.freeze({
    get_cell_ranges: ['read', 'semantic-read'],
    get_range_as_csv: ['read', 'semantic-read'],
    search_data: ['read', 'semantic-read'],
    screenshot_range: ['read', 'bounded-render-facts'],
    get_all_objects: ['read', 'semantic-read'],
    set_cell_range: ['mutate', 'transaction-proposal'],
    clear_cell_range: ['mutate', 'transaction-proposal'],
    copy_to: ['mutate', 'transaction-proposal'],
    modify_sheet_structure: ['mutate', 'transaction-proposal'],
    modify_workbook_structure: ['mutate', 'transaction-proposal'],
    resize_range: ['mutate', 'transaction-proposal'],
    modify_object: ['mutate', 'transaction-proposal'],
    eval_officejs: ['mutate', 'raw-office-proposal'],
    propose_raw_office_edit: ['mutate', 'raw-office-proposal'],
  }),
  'office-powerpoint': Object.freeze({
    inspect_slide_masters: ['read', 'semantic-read'],
    screenshot_slide: ['read', 'bounded-render-facts'],
    list_slide_shapes: ['read', 'semantic-read'],
    read_slide_text: ['read', 'semantic-read'],
    verify_slides: ['read', 'bounded-render-facts'],
    plan_deck: ['read', 'semantic-read'],
    edit_slide_text: ['mutate', 'transaction-proposal'],
    edit_slide_xml: ['mutate', 'transaction-proposal'],
    edit_slide_chart: ['mutate', 'transaction-proposal'],
    edit_slide_master: ['mutate', 'transaction-proposal'],
    edit_slide_master_xml: ['mutate', 'transaction-proposal'],
    duplicate_slide: ['mutate', 'transaction-proposal'],
    execute_office_js: ['mutate', 'raw-office-proposal'],
    propose_raw_office_edit: ['mutate', 'raw-office-proposal'],
  }),
}) as unknown as Readonly<
  Record<EnhancedHost, Readonly<Record<string, readonly [ToolMutability, RequiredCapability]>>>
>

export interface DocumentToolIdentity {
  readonly ownerId: string
  readonly host: EnhancedHost
  readonly documentId: string
  readonly sessionId: string
  readonly generation: number
}
export interface DocumentToolManifestInput {
  readonly policyGrant: unknown
  readonly consumePolicyGrant: (grant: unknown) => EnhancedPolicySnapshot
  readonly tools: readonly AgentToolDef[]
  readonly policy: Readonly<Record<string, ToolMutability>>
}
declare const manifestBrand: unique symbol
export interface DocumentToolManifest {
  readonly digest: string
  readonly [manifestBrand]: true
}
interface ManifestEntry {
  readonly authorization: EnhancedPolicySnapshot
  readonly tools: readonly AgentToolDef[]
  readonly policy: Readonly<Record<string, ToolMutability>>
  readonly digest: string
}
const manifestLedger = new WeakMap<object, ManifestEntry>()

interface PendingMutation {
  readonly callId: string
  readonly request: DetachedMutationRequest
  readonly controller: AbortController
  readonly resolve: (execution: ToolExecution) => void
  readonly promise: Promise<ToolExecution>
  timer: ReturnType<typeof setTimeout> | undefined
  state: 'queued' | 'claimed' | 'settled'
  finish(execution: ToolExecution): void
}
interface ClaimEntry {
  readonly authority: object
  readonly pending: PendingMutation
  consumed: boolean
}
const mutationClaims = new WeakMap<object, ClaimEntry>()

export interface DocumentToolRegistration {
  readonly identity: DocumentToolIdentity
  readonly manifest: DocumentToolManifest
  readonly isOpen: () => boolean
  readonly executeRead: (
    call: AgentToolCall,
    signal?: AbortSignal,
  ) => ToolExecution | Promise<ToolExecution>
  readonly suspendMutation: (result: Promise<ToolExecution>) => ToolExecutionSuspension
  readonly ownsSuspension: (value: ToolExecutionOutcome) => boolean
  readonly carrier?: Readonly<{ issuer: DocumentCarrierIssuer; capability: unknown }>
  readonly maxCallMs?: number
  readonly maxTotalCalls?: number
  readonly maxPendingMutations?: number
}
export interface ToolSessionCredentials {
  readonly sessionId: string
  readonly secret: string
}
export interface McpToolDefinition extends AgentToolDef {
  readonly annotations: { readonly readOnlyHint: boolean; readonly destructiveHint: boolean }
}
export interface CarrierRequest {
  readonly turnId: string
  readonly sourceNonce: string
  readonly toolName: string
}
declare const mutationClaimBrand: unique symbol
export interface MutationClaim {
  readonly [mutationClaimBrand]: true
}
export interface DetachedMutationRequest {
  readonly identity: Readonly<DocumentToolIdentity>
  readonly call: Readonly<AgentToolCall>
  readonly catalogDigest: string
}
export interface MutationAuthority {
  claimNext(): Readonly<{ claim: MutationClaim; request: DetachedMutationRequest }> | undefined
  settle(claim: MutationClaim, execution: ToolExecution): void
  reject(claim: MutationClaim, code?: string): void
}
export interface DocumentToolSession {
  readonly identity: Readonly<DocumentToolIdentity>
  readonly credentials: ToolSessionCredentials
  readonly catalogDigest: string
  readonly mutationAuthority: MutationAuthority
  authorize(credentials: ToolSessionCredentials): void
  listTools(credentials: ToolSessionCredentials): McpToolDefinition[]
  callTool(
    credentials: ToolSessionCredentials,
    call: AgentToolCall,
    signal?: AbortSignal,
  ): ToolExecutionOutcome | Promise<ToolExecution>
  issueCarrier(credentials: ToolSessionCredentials, turn: CarrierRequest): DocumentCarrierHandle
  cancel(credentials: ToolSessionCredentials, callId: string): boolean
  cancelAll(credentials: ToolSessionCredentials): number
  close(): void
}
export class ToolRouterError extends Error {
  constructor(readonly code: string) {
    super(code)
    this.name = 'ToolRouterError'
  }
}

const utf8Length = (value: string): number => Buffer.byteLength(value, 'utf8')
function plainRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false
  try {
    const prototype = Object.getPrototypeOf(value)
    return (
      (prototype === Object.prototype || prototype === null) &&
      Object.getOwnPropertySymbols(value).length === 0 &&
      Object.values(Object.getOwnPropertyDescriptors(value)).every((item) => 'value' in item)
    )
  } catch {
    return false
  }
}
function strictArrayValues(value: unknown, maximum: number, code: string): unknown[] {
  if (
    !Array.isArray(value) ||
    Object.getPrototypeOf(value) !== Array.prototype ||
    value.length > maximum ||
    Object.getOwnPropertySymbols(value).length > 0
  )
    throw new ToolRouterError(code)
  const descriptors = Object.getOwnPropertyDescriptors(value)
  const allowed = new Set([
    'length',
    ...Array.from({ length: value.length }, (_, index) => String(index)),
  ])
  if (Object.keys(descriptors).some((key) => !allowed.has(key))) throw new ToolRouterError(code)
  return Array.from({ length: value.length }, (_, index) => {
    const descriptor = descriptors[String(index)]
    if (!descriptor || !('value' in descriptor)) throw new ToolRouterError(code)
    return descriptor.value
  })
}
function inspectGraph(value: unknown, maxBytes: number, code: string): number {
  let encoded: string
  try {
    encoded = JSON.stringify(value)
  } catch {
    throw new ToolRouterError(code)
  }
  if (encoded === undefined || utf8Length(encoded) > maxBytes) throw new ToolRouterError(code)
  const pending: Array<{ value: unknown; depth: number }> = [{ value, depth: 0 }]
  const seen = new Set<object>()
  let nodes = 0
  while (pending.length) {
    const item = pending.pop()!
    if (++nodes > MAX_GRAPH_NODES || item.depth > MAX_GRAPH_DEPTH) throw new ToolRouterError(code)
    if (typeof item.value === 'string' && utf8Length(item.value) > maxBytes)
      throw new ToolRouterError(code)
    if (typeof item.value !== 'object' || item.value === null) continue
    if (seen.has(item.value)) throw new ToolRouterError(code)
    seen.add(item.value)
    if (Array.isArray(item.value))
      for (const child of strictArrayValues(item.value, MAX_GRAPH_NODES, code))
        pending.push({ value: child, depth: item.depth + 1 })
    else {
      if (!plainRecord(item.value)) throw new ToolRouterError(code)
      for (const child of Object.values(item.value))
        pending.push({ value: child, depth: item.depth + 1 })
    }
  }
  return nodes
}
function boundedId(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0 && utf8Length(value) <= MAX_ID_BYTES
}
function canonical(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`
  if (plainRecord(value))
    return `{${Object.keys(value)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${canonical(value[key])}`)
      .join(',')}}`
  const encoded = JSON.stringify(value)
  if (encoded === undefined) throw new ToolRouterError('invalid_tool_manifest')
  return encoded
}
function freezeDetached<T>(value: T): T {
  if (typeof value !== 'object' || value === null || Object.isFrozen(value)) return value
  for (const child of Object.values(value as Record<string, unknown>)) freezeDetached(child)
  return Object.freeze(value)
}

export function createDocumentToolManifest(input: DocumentToolManifestInput): DocumentToolManifest {
  if (!plainRecord(input) || !plainRecord(input.policy))
    throw new ToolRouterError('invalid_tool_manifest')
  const suppliedTools = strictArrayValues(input.tools, MAX_TOOLS, 'invalid_tool_manifest')
  if (suppliedTools.length < 1) throw new ToolRouterError('invalid_tool_manifest')
  let authorization: EnhancedPolicySnapshot
  try {
    if (typeof input.consumePolicyGrant !== 'function') throw new Error('invalid_enhanced_policy')
    const snapshot = input.consumePolicyGrant(input.policyGrant)
    if (!plainRecord(snapshot)) throw new Error('invalid_enhanced_policy')
    const generation = snapshot.generation
    const host = snapshot.host
    if (!Number.isSafeInteger(generation) || generation < 0 || !ENHANCED_HOSTS.includes(host))
      throw new Error('invalid_enhanced_policy')
    const policy = parseEnhancedRolloutPolicy(snapshot.policy)
    const capabilities = parseEnhancedCapabilities(snapshot.capabilities)
    if (
      !policy.globalEnabled ||
      !policy.hosts[host] ||
      (capabilities.includes('raw-office-proposal') &&
        (!host.startsWith('office-') || !policy.rawOfficeEnabled))
    )
      throw new Error('enhanced_policy_denied')
    authorization = Object.freeze({
      generation,
      host,
      policy,
      capabilities: Object.freeze(capabilities),
    })
  } catch (error) {
    throw new ToolRouterError(error instanceof Error ? error.message : 'invalid_enhanced_policy')
  }
  const catalog = CATALOG[authorization.host]
  const capabilities = new Set(authorization.capabilities)
  const names = new Set<string>()
  let descriptions = 0
  let schemas = 0
  let nodes = 0
  const tools = suppliedTools.map((tool) => {
    if (
      !plainRecord(tool) ||
      !boundedId(tool.name) ||
      utf8Length(tool.name) > MAX_NAME_BYTES ||
      names.has(tool.name) ||
      typeof tool.description !== 'string' ||
      utf8Length(tool.description) > MAX_DESCRIPTION_BYTES ||
      !plainRecord(tool.inputSchema)
    )
      throw new ToolRouterError('invalid_tool_definition')
    const compiled = catalog[tool.name]
    if (!compiled) throw new ToolRouterError('tool_not_compiled_for_host')
    if (!capabilities.has(compiled[1])) throw new ToolRouterError('tool_capability_denied')
    if (input.policy[tool.name] !== compiled[0]) throw new ToolRouterError('invalid_tool_policy')
    names.add(tool.name)
    descriptions += utf8Length(tool.description)
    schemas += utf8Length(JSON.stringify(tool.inputSchema))
    nodes += inspectGraph(tool.inputSchema, MAX_SCHEMA_BYTES, 'invalid_tool_definition')
    return Object.freeze({
      name: tool.name,
      description: tool.description,
      inputSchema: structuredClone(tool.inputSchema),
    })
  })
  if (
    Object.keys(input.policy).length !== tools.length ||
    descriptions > MAX_TOTAL_DESCRIPTION_BYTES ||
    schemas > MAX_TOTAL_SCHEMA_BYTES ||
    nodes > MAX_TOTAL_GRAPH_NODES
  )
    throw new ToolRouterError('tool_catalog_limit')
  const list = tools.map((tool) => ({
    ...tool,
    annotations: {
      readOnlyHint: input.policy[tool.name] === 'read',
      destructiveHint: input.policy[tool.name] === 'mutate',
    },
  }))
  if (utf8Length(JSON.stringify(list)) > MAX_LIST_BYTES)
    throw new ToolRouterError('tool_catalog_limit')
  const policy = Object.freeze(
    Object.fromEntries(tools.map((tool) => [tool.name, input.policy[tool.name]!])) as Record<
      string,
      ToolMutability
    >,
  )
  const digest = createHash('sha256')
    .update(
      canonical({
        host: authorization.host,
        rollout: authorization.policy,
        generation: authorization.generation,
        capabilities: [...authorization.capabilities].sort(),
        tools,
        policy,
      }),
    )
    .digest('hex')
  const handle = Object.create(null) as DocumentToolManifest
  Object.defineProperty(handle, 'digest', { value: digest, enumerable: false })
  Object.freeze(handle)
  manifestLedger.set(handle as object, {
    authorization,
    tools: Object.freeze(tools),
    policy,
    digest,
  })
  return handle
}

function validateIdentity(
  value: unknown,
  host: EnhancedHost,
  generation: number,
): Readonly<DocumentToolIdentity> {
  if (
    !plainRecord(value) ||
    Object.keys(value).length !== 5 ||
    !boundedId(value.ownerId) ||
    value.host !== host ||
    !boundedId(value.documentId) ||
    !boundedId(value.sessionId) ||
    value.generation !== generation
  )
    throw new ToolRouterError('invalid_tool_identity')
  return Object.freeze({
    ownerId: value.ownerId,
    host,
    documentId: value.documentId,
    sessionId: value.sessionId,
    generation,
  })
}
function decodeCanonical(value: unknown): Buffer | undefined {
  if (typeof value !== 'string' || !SECRET_PATTERN.test(value)) return undefined
  const decoded = Buffer.from(value, 'base64url')
  return decoded.length === SECRET_BYTES && decoded.toString('base64url') === value
    ? decoded
    : undefined
}
function stable(
  output: string,
  summary = 'Tool failed',
  isError = true,
  mutated = false,
): ToolExecution {
  return { output, summary, isError, mutated }
}
function validExecution(value: unknown): value is ToolExecution {
  if (
    !plainRecord(value) ||
    typeof value.output !== 'string' ||
    utf8Length(value.output) > MAX_OUTPUT_BYTES ||
    typeof value.summary !== 'string' ||
    utf8Length(value.summary) > MAX_SUMMARY_BYTES ||
    (value.isError !== undefined && typeof value.isError !== 'boolean') ||
    (value.mutated !== undefined && typeof value.mutated !== 'boolean')
  )
    return false
  if (value.modelContent !== undefined)
    try {
      inspectGraph(value.modelContent, MAX_OUTPUT_BYTES, 'invalid_tool_result')
    } catch {
      return false
    }
  return true
}
function awaitBounded<T>(
  pending: PromiseLike<T>,
  signal: AbortSignal,
  maxCallMs: number,
): Promise<T> {
  return new Promise((resolve, reject) => {
    let settled = false
    const timer = setTimeout(
      () => finish(() => reject(new ToolRouterError('tool_timeout'))),
      maxCallMs,
    )
    timer.unref()
    const onAbort = () => finish(() => reject(new ToolRouterError('tool_cancelled')))
    const finish = (action: () => void) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      signal.removeEventListener('abort', onAbort)
      action()
    }
    signal.addEventListener('abort', onAbort, { once: true })
    Promise.resolve(pending).then(
      (value) => finish(() => resolve(value)),
      () => finish(() => reject(new ToolRouterError('tool_execution_failed'))),
    )
    if (signal.aborted) onAbort()
  })
}

export function createDocumentToolSession(
  registration: DocumentToolRegistration,
): DocumentToolSession {
  if (
    !plainRecord(registration) ||
    Object.keys(registration).some(
      (key) =>
        ![
          'identity',
          'manifest',
          'isOpen',
          'executeRead',
          'suspendMutation',
          'ownsSuspension',
          'carrier',
          'maxCallMs',
          'maxTotalCalls',
          'maxPendingMutations',
        ].includes(key),
    ) ||
    typeof registration.isOpen !== 'function' ||
    typeof registration.executeRead !== 'function' ||
    typeof registration.suspendMutation !== 'function' ||
    typeof registration.ownsSuspension !== 'function'
  )
    throw new ToolRouterError('invalid_tool_session')
  const manifest = manifestLedger.get(registration.manifest as object)
  if (!manifest) throw new ToolRouterError('invalid_tool_manifest')
  const identity = validateIdentity(
    registration.identity,
    manifest.authorization.host,
    manifest.authorization.generation,
  )
  const maxCallMs = registration.maxCallMs ?? MAX_CALL_MS
  const maxTotalCalls = registration.maxTotalCalls ?? MAX_TOTAL_CALLS
  const maxPendingMutations = registration.maxPendingMutations ?? MAX_PENDING_MUTATIONS
  if (
    !Number.isSafeInteger(maxCallMs) ||
    maxCallMs <= 0 ||
    maxCallMs > MAX_CALL_MS ||
    !Number.isSafeInteger(maxTotalCalls) ||
    maxTotalCalls <= 0 ||
    maxTotalCalls > MAX_TOTAL_CALLS ||
    !Number.isSafeInteger(maxPendingMutations) ||
    maxPendingMutations <= 0 ||
    maxPendingMutations > MAX_PENDING_MUTATIONS
  )
    throw new ToolRouterError('invalid_tool_bounds')
  if (
    registration.carrier !== undefined &&
    (!plainRecord(registration.carrier) ||
      typeof registration.carrier.issuer?.issueForTurn !== 'function')
  )
    throw new ToolRouterError('invalid_carrier_issuer')
  const sessionId = randomBytes(SECRET_BYTES).toString('base64url')
  const secretBytes = randomBytes(SECRET_BYTES)
  const credentials = Object.freeze({ sessionId, secret: secretBytes.toString('base64url') })
  const pending = new Map<string, AbortController>()
  const pendingMutations = new Map<string, PendingMutation>()
  const mutationQueue: PendingMutation[] = []
  const authorityIdentity = Object.freeze(Object.create(null)) as object
  const consumed = new Set<string>()
  let closed = false
  const safeOpen = () => {
    try {
      return registration.isOpen() === true
    } catch {
      return false
    }
  }
  const authenticate = (candidate: ToolSessionCredentials) => {
    const id = decodeCanonical(candidate?.sessionId)
    const provided = decodeCanonical(candidate?.secret)
    if (
      !id ||
      !provided ||
      !timingSafeEqual(id, Buffer.from(sessionId, 'base64url')) ||
      !timingSafeEqual(provided, secretBytes)
    )
      throw new ToolRouterError('tool_unauthorized')
    if (closed || !safeOpen()) throw new ToolRouterError('tool_session_closed')
  }
  const close = () => {
    if (closed) return
    closed = true
    for (const controller of pending.values()) controller.abort()
    for (const mutation of pendingMutations.values()) mutation.controller.abort()
    pending.clear()
  }
  const listTools = (candidate: ToolSessionCredentials): McpToolDefinition[] => {
    authenticate(candidate)
    return manifest.tools.map((tool) => ({
      ...tool,
      inputSchema: structuredClone(tool.inputSchema),
      annotations: {
        readOnlyHint: manifest.policy[tool.name] === 'read',
        destructiveHint: manifest.policy[tool.name] === 'mutate',
      },
    }))
  }
  const callTool = (
    candidate: ToolSessionCredentials,
    call: AgentToolCall,
    outerSignal?: AbortSignal,
  ): ToolExecutionOutcome | Promise<ToolExecution> => {
    authenticate(candidate)
    if (
      !plainRecord(call) ||
      !boundedId(call.id) ||
      !boundedId(call.name) ||
      !plainRecord(call.input) ||
      call.inputError ||
      call.truncated
    )
      return stable('invalid_tool_call')
    if (pending.has(call.id) || pendingMutations.has(call.id) || consumed.has(call.id))
      return stable('tool_call_consumed')
    if (consumed.size >= maxTotalCalls) {
      close()
      return stable('tool_session_call_limit')
    }
    consumed.add(call.id)
    try {
      inspectGraph(call.input, MAX_INPUT_BYTES, 'invalid_tool_call')
    } catch {
      return stable('invalid_tool_call')
    }
    const mutability = manifest.policy[call.name]
    if (!mutability) return stable('unknown_tool')
    if (mutability === 'mutate') {
      if (pendingMutations.size >= maxPendingMutations) return stable('mutation_queue_full')
      const controller = new AbortController()
      const abort = () => controller.abort()
      outerSignal?.addEventListener('abort', abort, { once: true })
      if (outerSignal?.aborted) abort()
      let resolve!: (execution: ToolExecution) => void
      const promise = new Promise<ToolExecution>((done) => (resolve = done))
      const detachedCall = freezeDetached({
        id: call.id,
        name: call.name,
        input: structuredClone(call.input),
        ...(call.invocationId ? { invocationId: call.invocationId } : {}),
      })
      const request = Object.freeze({
        identity,
        call: detachedCall,
        catalogDigest: manifest.digest,
      })
      const mutation: PendingMutation = {
        callId: call.id,
        request,
        controller,
        resolve,
        promise,
        timer: undefined,
        state: 'queued',
        finish: () => undefined,
      }
      const finish = (execution: ToolExecution) => {
        if (mutation.state === 'settled') return
        mutation.state = 'settled'
        if (mutation.timer) clearTimeout(mutation.timer)
        outerSignal?.removeEventListener('abort', abort)
        pendingMutations.delete(call.id)
        resolve(execution)
      }
      mutation.finish = finish
      controller.signal.addEventListener(
        'abort',
        () => finish(stable('tool_cancelled', 'Tool cancelled')),
        { once: true },
      )
      mutation.timer = setTimeout(() => finish(stable('tool_timeout', 'Tool timed out')), maxCallMs)
      mutation.timer.unref()
      pendingMutations.set(call.id, mutation)
      mutationQueue.push(mutation)
      if (controller.signal.aborted) finish(stable('tool_cancelled', 'Tool cancelled'))
      const suspension = registration.suspendMutation(promise)
      if (!registration.ownsSuspension(suspension)) {
        finish(stable('tool_authority_denied', 'Tool authority denied'))
        return stable('tool_authority_denied', 'Tool authority denied')
      }
      return suspension
    }
    if (pending.size > 0 || pendingMutations.size > 0) return stable('tool_call_in_progress')
    const controller = new AbortController()
    const abort = () => controller.abort()
    outerSignal?.addEventListener('abort', abort, { once: true })
    if (outerSignal?.aborted) abort()
    pending.set(call.id, controller)
    return (async () => {
      try {
        if (controller.signal.aborted) return stable('tool_cancelled', 'Tool cancelled')
        let execution: ToolExecution
        try {
          execution = await awaitBounded(
            Promise.resolve(registration.executeRead(call, controller.signal)),
            controller.signal,
            maxCallMs,
          )
        } catch (error) {
          return stable(error instanceof ToolRouterError ? error.code : 'tool_execution_failed')
        }
        return validExecution(execution) && execution.mutated !== true
          ? execution
          : stable('tool_policy_violation')
      } finally {
        outerSignal?.removeEventListener('abort', abort)
        pending.delete(call.id)
      }
    })()
  }
  const settleClaim = (claim: MutationClaim, execution: ToolExecution): void => {
    const entry = mutationClaims.get(claim as object)
    if (!entry) throw new ToolRouterError('invalid_mutation_claim')
    if (entry.authority !== authorityIdentity)
      throw new ToolRouterError('mutation_claim_issuer_mismatch')
    if (entry.consumed || entry.pending.state !== 'claimed')
      throw new ToolRouterError('mutation_claim_consumed')
    entry.consumed = true
    if (!validExecution(execution)) {
      entry.pending.finish(stable('invalid_tool_result', 'Tool failed', true, true))
      throw new ToolRouterError('invalid_tool_result')
    }
    entry.pending.finish(execution)
  }
  const mutationAuthority: MutationAuthority = Object.freeze({
    claimNext() {
      let mutation: PendingMutation | undefined
      while (mutationQueue.length > 0) {
        const candidate = mutationQueue.shift()!
        if (candidate.state === 'queued') {
          mutation = candidate
          break
        }
      }
      if (!mutation) return undefined
      mutation.state = 'claimed'
      const claim = Object.freeze(Object.create(null)) as MutationClaim
      mutationClaims.set(claim as object, {
        authority: authorityIdentity,
        pending: mutation,
        consumed: false,
      })
      return Object.freeze({ claim, request: mutation.request })
    },
    settle: settleClaim,
    reject(claim: MutationClaim, code = 'mutation_rejected') {
      if (!boundedId(code)) throw new ToolRouterError('invalid_mutation_rejection')
      settleClaim(claim, stable(code, 'Mutation rejected'))
    },
  })
  const session: DocumentToolSession = {
    identity,
    credentials,
    catalogDigest: manifest.digest,
    mutationAuthority,
    authorize: authenticate,
    listTools,
    callTool,
    issueCarrier(candidate, turn) {
      authenticate(candidate)
      if (!registration.carrier) throw new ToolRouterError('carrier_not_available')
      if (
        !plainRecord(turn) ||
        Object.keys(turn).length !== 3 ||
        !boundedId(turn.turnId) ||
        typeof turn.sourceNonce !== 'string' ||
        !/^[A-Za-z0-9_-]{43}$/.test(turn.sourceNonce) ||
        typeof turn.toolName !== 'string' ||
        !manifest.policy[turn.toolName]
      )
        throw new ToolRouterError('invalid_carrier_authorization')
      return registration.carrier.issuer.issueForTurn({
        ...turn,
        capability: registration.carrier.capability,
        method: `mcp__wiswork__${turn.toolName}`,
        schemaDigest: manifest.digest,
      })
    },
    cancel(candidate, callId) {
      authenticate(candidate)
      const controller = pending.get(callId)
      const mutation = pendingMutations.get(callId)
      controller?.abort()
      mutation?.controller.abort()
      return controller !== undefined || mutation !== undefined
    },
    cancelAll(candidate) {
      authenticate(candidate)
      const count = pending.size + pendingMutations.size
      for (const controller of pending.values()) controller.abort()
      for (const mutation of pendingMutations.values()) mutation.controller.abort()
      return count
    },
    close,
  }
  return Object.freeze(session)
}
