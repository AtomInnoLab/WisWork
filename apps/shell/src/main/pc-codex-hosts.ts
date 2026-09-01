import { randomBytes } from 'node:crypto'
import { createAgentHarness } from '@wiswork/agent-harness'
import type { AgentSkill, AgentToolCall, ToolExecution } from '@wiswork/agent-core'
import {
  PC_HOST_CODEX_CHANNELS,
  type EnhancedRolloutPolicy,
  type PcEnhancedHost,
  type PcHostRegistration,
  type PcHostToolResult,
} from '@wiswork/agent-runtime'
import { createDocumentToolManifest, createDocumentToolSession } from '@wiswork/codex-bridge'
import type { ShellCodexRuntime, CodexOwner } from './codex-runtime'
import { createShellEnhancedPolicyAuthority } from './enhanced-policy-authority'

interface PcOwner extends CodexOwner {
  readonly id: number
  send(channel: string, value: unknown): void
}
interface IpcMainLike {
  handle(
    channel: string,
    handler: (event: { sender: PcOwner }, ...args: unknown[]) => unknown,
  ): void
}

type Pending = {
  readonly call: AgentToolCall
  readonly resolve: (execution: ToolExecution) => void
  readonly reject: () => void
  readonly claim?: object
}
type PendingProposal = {
  readonly proposalId: string
  readonly call: AgentToolCall
  readonly expiresAt: number
}
type HostRecord = {
  readonly owner: PcOwner
  readonly documentId: string
  readonly generation: number
  readonly closeRuntime: () => Promise<void>
  readonly harness: ReturnType<typeof createAgentHarness>
  readonly pending: Map<string, Pending>
  readonly proposals: Map<string, PendingProposal>
  readonly session: ReturnType<typeof createDocumentToolSession>
  text: string
  closed: boolean
}

const ID = /^[A-Za-z0-9._:@/-]{1,256}$/
const PC_HOSTS = new Set<PcEnhancedHost>(['latex', 'slides', 'docs', 'sheets'])
const exactObject = (
  value: unknown,
  keys: readonly string[],
): value is globalThis.Record<string, unknown> => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  if (
    Object.getPrototypeOf(value) !== Object.prototype ||
    Object.getOwnPropertySymbols(value).length
  )
    return false
  const descriptors = Object.getOwnPropertyDescriptors(value)
  return (
    Object.keys(descriptors).length === keys.length &&
    keys.every((key) => 'value' in (descriptors[key] ?? {}))
  )
}
const detachedExecution = (value: unknown): ToolExecution => {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.getPrototypeOf(value) !== Object.prototype
  )
    throw new Error('enhanced_invalid_request')
  const descriptors = Object.getOwnPropertyDescriptors(value)
  const allowed = new Set(['output', 'summary', 'isError', 'mutated', 'stopToolBatch'])
  if (
    Object.getOwnPropertySymbols(value).length ||
    Object.keys(descriptors).some((key) => !allowed.has(key)) ||
    Object.values(descriptors).some((descriptor) => !('value' in descriptor))
  )
    throw new Error('enhanced_invalid_request')
  const output = descriptors.output?.value
  const summary = descriptors.summary?.value
  const isError = descriptors.isError?.value
  const mutated = descriptors.mutated?.value
  const stopToolBatch = descriptors.stopToolBatch?.value
  if (
    typeof output !== 'string' ||
    Buffer.byteLength(output) > 1_000_000 ||
    typeof summary !== 'string' ||
    Buffer.byteLength(summary) > 4_096 ||
    (isError !== undefined && typeof isError !== 'boolean') ||
    (mutated !== undefined && typeof mutated !== 'boolean') ||
    (stopToolBatch !== undefined && typeof stopToolBatch !== 'boolean')
  )
    throw new Error('enhanced_invalid_request')
  return Object.freeze({
    output,
    summary,
    ...(isError === undefined ? {} : { isError }),
    ...(mutated === undefined ? {} : { mutated }),
    ...(stopToolBatch === undefined ? {} : { stopToolBatch }),
  })
}

export function registerPcCodexHosts(options: {
  readonly ipcMain: IpcMainLike
  readonly runtime: ShellCodexRuntime
  readonly policy: EnhancedRolloutPolicy
  readonly hostForOwner: (owner: PcOwner) => PcEnhancedHost | null
}) {
  const records = new Map<PcOwner, HostRecord>()
  const byDocument = new Map<string, HostRecord>()
  const trusted = (owner: PcOwner): PcEnhancedHost => {
    const host = owner.isDestroyed() ? null : options.hostForOwner(owner)
    if (!host || !PC_HOSTS.has(host)) throw new Error('enhanced_untrusted_request')
    return host
  }
  const send = (record: HostRecord, channel: string, value: unknown) => {
    if (record.closed || record.owner.isDestroyed()) return
    try {
      record.owner.send(channel, value)
    } catch {}
  }
  const dispatch = (
    record: HostRecord,
    call: AgentToolCall,
    claim?: object,
  ): Promise<ToolExecution> =>
    new Promise((resolve, reject) => {
      if (record.closed || record.pending.has(call.id))
        return reject(new Error('tool_session_closed'))
      record.pending.set(call.id, {
        call,
        resolve,
        reject: () => reject(new Error('tool_execution_failed')),
        claim,
      })
      send(record, PC_HOST_CODEX_CHANNELS.event, { type: 'tool-start', call })
      send(record, PC_HOST_CODEX_CHANNELS.toolCall, {
        documentId: record.documentId,
        generation: record.generation,
        call,
      })
    })
  const closeRecord = async (record: HostRecord) => {
    if (record.closed) return
    record.closed = true
    records.delete(record.owner)
    byDocument.delete(record.documentId)
    for (const pending of record.pending.values()) pending.reject()
    record.pending.clear()
    for (const proposal of record.proposals.values()) {
      const claimed = record.session.mutationAuthority.claimNext()
      if (claimed && claimed.request.call.id === proposal.call.id)
        record.session.mutationAuthority.reject(claimed.claim, 'mutation_cancelled')
    }
    record.proposals.clear()
    record.session.close()
    record.harness.dispose()
    await record.closeRuntime()
  }
  const handleEngineEvent = (
    record: HostRecord,
    event: import('./codex-runtime').CodexRuntimeEngineEvent,
  ) => {
    if (event.type === 'text') {
      record.text += event.text
      send(record, PC_HOST_CODEX_CHANNELS.event, event)
      return
    }
    if (event.type === 'proposal') {
      if (record.proposals.has(event.proposalId) || record.pending.has(event.call.id)) return
      record.proposals.set(event.proposalId, event)
      send(record, PC_HOST_CODEX_CHANNELS.proposal, {
        proposalId: event.proposalId,
        documentId: record.documentId,
        generation: record.generation,
        toolName: event.call.name,
        summary: 'Review the proposed document change',
        expiresAt: event.expiresAt,
      })
      return
    }
    if (event.type !== 'terminal') return
    if (event.status === 'failed')
      send(record, PC_HOST_CODEX_CHANNELS.event, { type: 'error', code: 'enhanced_turn_failed' })
    else
      send(record, PC_HOST_CODEX_CHANNELS.event, {
        type: 'done',
        result: {
          text: record.text,
          cancelled: event.status === 'cancelled',
          turnLimit: false,
        },
      })
    record.text = ''
  }

  options.ipcMain.handle(PC_HOST_CODEX_CHANNELS.status, (event, ...args) => {
    trusted(event.sender)
    if (args.length) throw new Error('enhanced_invalid_request')
    const record = records.get(event.sender)
    return {
      activeAgentRuntime: options.runtime.configuredAgentRuntime,
      documentId: record?.documentId ?? null,
    }
  })

  options.ipcMain.handle(PC_HOST_CODEX_CHANNELS.register, async (event, value) => {
    const host = trusted(event.sender)
    if (
      !exactObject(value, [
        'host',
        'documentId',
        'generation',
        'systemPrompt',
        'tools',
        'mutatingTools',
      ])
    )
      throw new Error('enhanced_invalid_request')
    const input = value as unknown as PcHostRegistration
    if (input.host !== host) throw new Error('enhanced_untrusted_request')
    if (
      !ID.test(input.documentId) ||
      !Number.isSafeInteger(input.generation) ||
      input.generation < 0 ||
      typeof input.systemPrompt !== 'string' ||
      Buffer.byteLength(input.systemPrompt) > 256_000 ||
      !Array.isArray(input.tools) ||
      !Array.isArray(input.mutatingTools)
    )
      throw new Error('enhanced_invalid_request')
    if (records.has(event.sender) || byDocument.has(input.documentId))
      throw new Error('enhanced_document_exists')
    const authority = createShellEnhancedPolicyAuthority(() => input.generation)
    const grant = authority.issue({
      generation: input.generation,
      host,
      policy: options.policy,
      capabilities: ['semantic-read', 'transaction-proposal'],
    })
    const mutating = new Set(input.mutatingTools)
    const manifest = createDocumentToolManifest({
      policyGrant: grant,
      consumePolicyGrant: (candidate) => authority.consume(candidate as never),
      tools: input.tools,
      policy: Object.fromEntries(
        input.tools.map((tool) => [tool.name, mutating.has(tool.name) ? 'mutate' : 'read']),
      ),
    })
    const inertSkill: AgentSkill = {
      id: `enhanced-${host}`,
      systemPrompt: '',
      tools: [],
      executeTool: async () => ({ output: 'denied', summary: 'denied', isError: true }),
    }
    const harness = createAgentHarness({
      skill: inertSkill,
      maxTurns: 1,
      transport: {
        stream: (_request, callbacks) => {
          callbacks.onError('transport_disabled')
          return { cancel() {} }
        },
      },
    })
    let record!: HostRecord
    const session = createDocumentToolSession({
      identity: {
        ownerId: `renderer-${event.sender.id}`,
        host,
        documentId: input.documentId,
        sessionId: randomBytes(16).toString('hex'),
        generation: input.generation,
      },
      manifest,
      isOpen: () => !record.closed && !event.sender.isDestroyed(),
      executeRead: (call) => dispatch(record, call),
      suspendMutation: (result) => {
        return harness.suspendToolExecution(result)
      },
      ownsSuspension: (value) => harness.ownsToolExecutionSuspension(value),
    })
    const close = options.runtime.registerDocument({
      owner: event.sender,
      documentId: input.documentId,
      host,
      generation: input.generation,
      toolSession: session,
      instructions: input.systemPrompt,
      onEvent: (engineEvent) => handleEngineEvent(record, engineEvent),
    })
    record = {
      owner: event.sender,
      documentId: input.documentId,
      generation: input.generation,
      closeRuntime: close.close,
      harness,
      pending: new Map(),
      proposals: new Map(),
      session,
      text: '',
      closed: false,
    }
    records.set(event.sender, record)
    byDocument.set(input.documentId, record)
  })

  options.ipcMain.handle(
    PC_HOST_CODEX_CHANNELS.unregister,
    async (event, documentId, generation) => {
      trusted(event.sender)
      const record = records.get(event.sender)
      if (!record || record.documentId !== documentId || record.generation !== generation)
        throw new Error('enhanced_untrusted_request')
      await closeRecord(record)
    },
  )
  options.ipcMain.handle(PC_HOST_CODEX_CHANNELS.toolResult, (event, value) => {
    trusted(event.sender)
    if (!exactObject(value, ['documentId', 'generation', 'callId', 'execution']))
      throw new Error('enhanced_invalid_request')
    const result = value as unknown as PcHostToolResult
    const execution = detachedExecution(result.execution)
    const record = records.get(event.sender)
    const pending = record?.pending.get(result.callId)
    if (
      !record ||
      record.documentId !== result.documentId ||
      record.generation !== result.generation ||
      !pending
    )
      throw new Error('enhanced_untrusted_request')
    record.pending.delete(result.callId)
    send(record, PC_HOST_CODEX_CHANNELS.event, {
      type: 'tool-executed',
      event: { call: pending.call, execution },
    })
    if (pending.claim) record.session.mutationAuthority.settle(pending.claim as never, execution)
    pending.resolve(execution)
  })
  const proposalRecord = (
    owner: PcOwner,
    documentId: unknown,
    generation: unknown,
    proposalId: unknown,
  ) => {
    trusted(owner)
    if (
      typeof documentId !== 'string' ||
      typeof proposalId !== 'string' ||
      !Number.isSafeInteger(generation) ||
      !ID.test(proposalId)
    )
      throw new Error('enhanced_invalid_request')
    const record = records.get(owner)
    const proposal = record?.proposals.get(proposalId)
    if (
      !record ||
      record.documentId !== documentId ||
      record.generation !== generation ||
      !proposal
    )
      throw new Error('enhanced_untrusted_request')
    if (proposal.expiresAt <= Date.now()) {
      record.proposals.delete(proposalId)
      throw new Error('enhanced_proposal_expired')
    }
    return { record, proposal }
  }
  options.ipcMain.handle(
    PC_HOST_CODEX_CHANNELS.confirmProposal,
    (event, documentId, generation, proposalId) => {
      const { record, proposal } = proposalRecord(event.sender, documentId, generation, proposalId)
      record.proposals.delete(proposal.proposalId)
      const claimed = record.session.mutationAuthority.claimNext()
      if (
        !claimed ||
        claimed.request.call.id !== proposal.call.id ||
        claimed.request.call.name !== proposal.call.name ||
        claimed.request.catalogDigest !== record.session.catalogDigest ||
        claimed.request.identity.ownerId !== record.session.identity.ownerId ||
        claimed.request.identity.documentId !== record.documentId ||
        claimed.request.identity.generation !== record.generation
      ) {
        if (claimed)
          record.session.mutationAuthority.reject(claimed.claim, 'mutation_binding_mismatch')
        throw new Error('enhanced_untrusted_request')
      }
      return dispatch(record, claimed.request.call, claimed.claim).then(
        () => undefined,
        () => {
          record.session.mutationAuthority.reject(claimed.claim, 'tool_execution_failed')
        },
      )
    },
  )
  options.ipcMain.handle(
    PC_HOST_CODEX_CHANNELS.cancelProposal,
    (event, documentId, generation, proposalId) => {
      const { record, proposal } = proposalRecord(event.sender, documentId, generation, proposalId)
      record.proposals.delete(proposal.proposalId)
      const claimed = record.session.mutationAuthority.claimNext()
      if (!claimed || claimed.request.call.id !== proposal.call.id) {
        if (claimed)
          record.session.mutationAuthority.reject(claimed.claim, 'mutation_binding_mismatch')
        throw new Error('enhanced_untrusted_request')
      }
      record.session.mutationAuthority.reject(claimed.claim, 'mutation_cancelled')
    },
  )
  return Object.freeze({
    documentIdForOwner: (owner: CodexOwner) => records.get(owner as PcOwner)?.documentId ?? null,
    closeOwner: async (owner: CodexOwner) => {
      const record = records.get(owner as PcOwner)
      if (record) await closeRecord(record)
    },
    close: async () => {
      for (const record of [...records.values()]) await closeRecord(record)
    },
  })
}
