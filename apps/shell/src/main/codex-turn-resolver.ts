import { randomBytes } from 'node:crypto'
import type { DocumentCarrierIssuer, PreparedResponsesTurn } from '@wiswork/codex-bridge'

interface ResolvedMetadataIds {
  threadId: string
  turnId: string
  sessionId: string
  namespacedMethods: string[]
  advertisedMethods: string[]
}
interface ArmedTurn {
  readonly issuerFor: (ids: ResolvedMetadataIds) => DocumentCarrierIssuer
  readonly capability: unknown
  consumed: boolean
}

function metadataIds(value: unknown): ResolvedMetadataIds {
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    throw new Error('turn_unbound')
  const request = value as Record<string, unknown>
  if (typeof request.client_metadata !== 'object' || request.client_metadata === null)
    throw new Error('turn_unbound')
  const metadata = request.client_metadata as Record<string, unknown>
  let threadId = metadata.thread_id
  let turnId = metadata.turn_id
  let sessionId = metadata.session_id
  let namespacedMethods: string[] = []
  let advertisedMethods: string[] = []
  // Packed metadata is authoritative; the top-level session can identify the process.
  if (typeof metadata['x-codex-turn-metadata'] === 'string') {
    const packed = metadata['x-codex-turn-metadata']
    if (Buffer.byteLength(packed) > 128_000) throw new Error('turn_unbound')
    const parsed = JSON.parse(packed) as unknown
    if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed))
      throw new Error('turn_unbound')
    threadId = (parsed as Record<string, unknown>).thread_id
    turnId = (parsed as Record<string, unknown>).turn_id
    sessionId = (parsed as Record<string, unknown>).session_id
    const methods = (parsed as Record<string, unknown>).code_mode_tool_names
    if (typeof methods === 'object' && methods !== null && !Array.isArray(methods)) {
      advertisedMethods = Object.keys(methods).filter((method) =>
        /^[a-zA-Z0-9_]{1,128}$/.test(method),
      )
      namespacedMethods = advertisedMethods.filter((method) =>
        /^mcp__[a-zA-Z0-9_]{1,128}$/.test(method),
      )
    }
  }
  if (
    typeof threadId !== 'string' ||
    !threadId ||
    typeof turnId !== 'string' ||
    !turnId ||
    typeof sessionId !== 'string' ||
    !sessionId
  )
    throw new Error('turn_unbound')
  return { threadId, turnId, sessionId, namespacedMethods, advertisedMethods }
}

/** Shell-private, one-use binding from app-server metadata to document authority. */
export class CodexTurnResolver {
  readonly #turns = new Map<string, ArmedTurn>()
  constructor(private readonly diagnostics?: (code: string) => void) {}
  arm(threadId: string, issuerFor: ArmedTurn['issuerFor'], capability: unknown): () => void {
    if (!threadId || this.#turns.has(threadId)) throw new Error('turn_already_armed')
    const entry = { issuerFor, capability, consumed: false }
    this.#turns.set(threadId, entry)
    return () => {
      if (this.#turns.get(threadId) === entry) this.#turns.delete(threadId)
    }
  }
  prepare = (input: unknown): PreparedResponsesTurn => {
    let ids: ResolvedMetadataIds
    try {
      ids = metadataIds(input)
    } catch (error) {
      this.diagnostics?.('resolver_metadata_invalid')
      throw new Error('turn_unbound', { cause: error })
    }
    const { threadId, turnId } = ids
    this.diagnostics?.(`resolver_namespaced_count_${Math.min(ids.namespacedMethods.length, 9)}`)
    this.diagnostics?.(`resolver_advertised_count_${Math.min(ids.advertisedMethods.length, 9)}`)
    this.diagnostics?.(
      ids.namespacedMethods.includes('mcp__wiswork__wiswork_call')
        ? 'resolver_method_exact'
        : 'resolver_method_missing',
    )
    const entry = this.#turns.get(threadId)
    if (!entry || entry.consumed) {
      this.diagnostics?.('resolver_authority_unbound')
      throw new Error('turn_unbound')
    }
    // Consume before converter/schema/catalog work so hostile retries cannot race authority.
    entry.consumed = true
    this.#turns.delete(threadId)
    let handle
    const issuer = entry.issuerFor(ids)
    try {
      handle = issuer.issueForTurn({
        turnId,
        sourceNonce: randomBytes(32).toString('base64url'),
        capability: entry.capability,
        method: 'mcp__wiswork__wiswork_call',
        toolName: 'wiswork_call',
        schemaDigest: '0'.repeat(64),
      })
    } catch (error) {
      this.diagnostics?.('resolver_issue_rejected')
      throw new Error('turn_unbound', { cause: error })
    }
    try {
      return issuer.prepareTurn(input, {}, handle)
    } catch (error) {
      const code =
        error instanceof Error && /^[a-z_]{1,64}$/.test(error.message) ? error.message : undefined
      if (code) this.diagnostics?.(`resolver_protocol_${code}`)
      this.diagnostics?.('resolver_protocol_rejected')
      throw new Error('turn_unbound', { cause: error })
    }
  }
  clear(): void {
    this.#turns.clear()
  }
}
