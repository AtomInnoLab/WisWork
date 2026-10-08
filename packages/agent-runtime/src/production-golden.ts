import type { AgentSkill, ToolExecutedEvent } from '@wiswork/agent-core'
import type { EnhancedHost } from './contracts'
import { EnhancedAgentRuntime } from './enhanced'
import type { EnhancedRendererBridge } from './renderer'
import { createEnhancedRendererClient } from './renderer'

type PcEnhancedHost = Extract<EnhancedHost, 'latex' | 'slides' | 'docs' | 'sheets'>
type OfficeEnhancedHost = Exclude<EnhancedHost, PcEnhancedHost>

export interface EnhancedGoldenDependencies<
  TConfirmation extends { readonly mutationReceiptId: string },
  TVerification,
  TRollback,
> {
  readonly documentId: string
  readonly generation: number
  readonly instruction: string
  readonly bridge: EnhancedRendererBridge
  readonly skill: AgentSkill
  readonly captureSnapshot?: () => unknown
  confirm(execution: ToolExecutedEvent<unknown>['execution']): Promise<TConfirmation>
  readback(confirmation: TConfirmation): Promise<TVerification>
  rollback(confirmation: TConfirmation): Promise<TRollback>
}

export interface EnhancedGoldenResult<TVerification, TRollback> {
  readonly host: EnhancedHost
  readonly mutationReceiptId: string
  readonly verification: TVerification
  readonly rollback: TRollback
}

export interface OfficeEnhancedGoldenDependencies<
  TExecution,
  TConfirmation extends { readonly mutationReceiptId: string },
  TVerification,
  TRollback,
> {
  executeTurn(): Promise<TExecution>
  confirm(execution: TExecution): Promise<TConfirmation>
  readback(confirmation: TConfirmation): Promise<TVerification>
  rollback(confirmation: TConfirmation): Promise<TRollback>
}

/**
 * Production-path golden orchestrator for desktop hosts. The bridge still owns the remote turn,
 * while the renderer client performs the real registration, allowlisting and skill dispatch.
 */
export async function runEnhancedGolden<
  TConfirmation extends { readonly mutationReceiptId: string },
  TVerification,
  TRollback,
>(
  host: EnhancedHost,
  deps:
    | EnhancedGoldenDependencies<TConfirmation, TVerification, TRollback>
    | OfficeEnhancedGoldenDependencies<unknown, TConfirmation, TVerification, TRollback>,
): Promise<EnhancedGoldenResult<TVerification, TRollback>> {
  if (host.startsWith('office-')) {
    const office = deps as OfficeEnhancedGoldenDependencies<
      unknown,
      TConfirmation,
      TVerification,
      TRollback
    >
    const execution = await office.executeTurn()
    const confirmation = await office.confirm(execution)
    return {
      host: host as OfficeEnhancedHost,
      mutationReceiptId: confirmation.mutationReceiptId,
      verification: await office.readback(confirmation),
      rollback: await office.rollback(confirmation),
    }
  }
  const desktop = deps as EnhancedGoldenDependencies<TConfirmation, TVerification, TRollback>
  const runtime = new EnhancedAgentRuntime(createEnhancedRendererClient(desktop.bridge))
  let resolveExecution!: (execution: ToolExecutedEvent<unknown>['execution']) => void
  let rejectExecution!: (error: Error) => void
  const execution = new Promise<ToolExecutedEvent<unknown>['execution']>((resolve, reject) => {
    resolveExecution = resolve
    rejectExecution = reject
  })
  const session = runtime.createSession({
    host,
    document: { id: desktop.documentId, generation: desktop.generation },
    skill: desktop.skill,
    transport: {
      stream() {
        throw new Error('standard_transport_reached_from_enhanced_golden')
      },
    },
    ...(desktop.captureSnapshot ? { captureSnapshot: desktop.captureSnapshot } : {}),
    events: {
      onToolExecuted: (event) => resolveExecution(event.execution),
      onError: (code) => rejectExecution(new Error(code)),
    },
  })
  try {
    if (!session.run(desktop.instruction)) throw new Error('enhanced_golden_turn_rejected')
    const toolExecution = await execution
    if (toolExecution.isError) throw new Error('enhanced_golden_tool_failed')
    const confirmation = await desktop.confirm(toolExecution)
    const verification = await desktop.readback(confirmation)
    const rollback = await desktop.rollback(confirmation)
    return {
      host,
      mutationReceiptId: confirmation.mutationReceiptId,
      verification,
      rollback,
    }
  } finally {
    await runtime.dispose()
  }
}
