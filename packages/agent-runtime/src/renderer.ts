import type { EnhancedHost } from './contracts'
import { isToolExecutionSuspension, type AgentSkill } from '@wiswork/agent-core'
import type { PcHostRegistration, PcHostToolRequest, PcHostToolResult } from './pc-host'
import type {
  EnhancedRuntimeClient,
  EnhancedRuntimeClientSession,
  EnhancedSessionEvent,
} from './enhanced'

export interface EnhancedRendererBridge {
  status(): Promise<{
    readonly activeAgentRuntime: 'standard' | 'enhanced'
    readonly documentId: string | null
  }>
  startTurn(input: { readonly documentId: string; readonly text: string }): Promise<void>
  cancelTurn(documentId: string): Promise<void>
  subscribe(documentId: string, listener: (event: EnhancedSessionEvent) => void): () => void
  register(input: PcHostRegistration): Promise<void>
  unregister(documentId: string, generation: number): Promise<void>
  onToolCall(listener: (request: PcHostToolRequest) => void): () => void
  toolResult(input: PcHostToolResult): Promise<void>
}

const READ_TOOLS: Readonly<Record<EnhancedHost, ReadonlySet<string>>> = Object.freeze({
  latex: new Set([
    'list_project_files',
    'search_project_text',
    'read_project_text',
    'get_compile_diagnostics',
  ]),
  docs: new Set(['get_document_context', 'read_blocks']),
  sheets: new Set([
    'get_workbook_context',
    'read_range',
    'load_guide',
    'read_formats',
    'read_sheet_features',
    'read_cells',
  ]),
  slides: new Set([
    'get_deck_context',
    'read_slide',
    'ask_clarification',
    'plan_deck',
    'list_style_templates',
  ]),
  'office-word': new Set<string>(),
  'office-excel': new Set<string>(),
  'office-powerpoint': new Set<string>(),
})
const PC_ALLOWED_TOOLS: Readonly<
  Record<'latex' | 'docs' | 'sheets' | 'slides', ReadonlySet<string>>
> = Object.freeze({
  latex: new Set([
    'list_project_files',
    'search_project_text',
    'read_project_text',
    'get_compile_diagnostics',
    'compile_project',
    'propose_project_edits',
  ]),
  docs: new Set([
    'get_document_context',
    'read_blocks',
    'insert_content',
    'replace_blocks',
    'apply_commands',
    'insert_image',
    'insert_chart',
    'edit_chart',
  ]),
  sheets: new Set([
    'get_workbook_context',
    'read_range',
    'load_guide',
    'read_formats',
    'read_sheet_features',
    'read_cells',
    'propose_operations',
  ]),
  slides: new Set([
    'get_deck_context',
    'read_slide',
    'ask_clarification',
    'plan_deck',
    'list_style_templates',
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
  ]),
})

export function createPcHostRegistration(input: {
  readonly host: EnhancedHost
  readonly documentId: string
  readonly generation: number
  readonly skill: AgentSkill
}): PcHostRegistration {
  if (!(input.host in PC_ALLOWED_TOOLS)) throw new Error('enhanced_host_unavailable')
  const tools = input.skill.tools.filter((tool) =>
    PC_ALLOWED_TOOLS[input.host as keyof typeof PC_ALLOWED_TOOLS].has(tool.name),
  )
  return {
    host: input.host as never,
    documentId: input.documentId,
    generation: input.generation,
    systemPrompt: input.skill.systemPrompt,
    tools,
    mutatingTools: tools
      .filter((tool) => !READ_TOOLS[input.host].has(tool.name))
      .map((tool) => tool.name),
  }
}

/** Renderer-only adapter. It holds no component path, process token, or document authority. */
export function createEnhancedRendererClient(
  bridge: EnhancedRendererBridge,
): EnhancedRuntimeClient {
  let closed = false
  const sessions = new Set<EnhancedRuntimeClientSession>()
  return Object.freeze({
    open(input: {
      host: EnhancedHost
      documentId: string
      generation: number
      skill: AgentSkill
      captureSnapshot?: () => unknown
    }) {
      if (closed) throw new Error('enhanced_runtime_closed')
      let ended = false
      const listeners = new Set<(event: EnhancedSessionEvent) => void>()
      const hostRegistration = createPcHostRegistration(input)
      const mutatingTools = new Set(hostRegistration.mutatingTools)
      const snapshots = new Map<string, Readonly<{ id: string; value: unknown }>>()
      const registration = bridge.register(hostRegistration)
      const unsubscribe = bridge.subscribe(input.documentId, (event) => {
        if (ended) return
        let restored = event
        if (event.type === 'tool-executed') {
          const snapshot = snapshots.get(event.event.call.id)
          if (snapshot) {
            if (snapshot.id !== event.event.snapshotBefore) return
            snapshots.delete(event.event.call.id)
            restored = {
              ...event,
              event: { ...event.event, snapshotBefore: snapshot.value },
            }
          } else if (event.event.snapshotBefore !== undefined) {
            return
          }
        }
        for (const listener of [...listeners]) listener(restored)
      })
      const unsubscribeTools = bridge.onToolCall((request) => {
        if (
          ended ||
          request.documentId !== input.documentId ||
          request.generation !== input.generation
        )
          return
        let snapshot: Readonly<{ id: string; value: unknown }> | undefined
        if (
          mutatingTools.has(request.call.name) &&
          (input.host === 'docs' || input.host === 'sheets')
        ) {
          if (!input.captureSnapshot) {
            void bridge.toolResult({
              documentId: input.documentId,
              generation: input.generation,
              callId: request.call.id,
              execution: {
                output: 'snapshot_unavailable',
                summary: 'Tool failed',
                isError: true,
                mutated: false,
              },
            })
            return
          }
          try {
            snapshot = Object.freeze({ id: crypto.randomUUID(), value: input.captureSnapshot() })
            snapshots.set(request.call.id, snapshot)
          } catch {
            void bridge.toolResult({
              documentId: input.documentId,
              generation: input.generation,
              callId: request.call.id,
              execution: {
                output: 'snapshot_failed',
                summary: 'Tool failed',
                isError: true,
                mutated: false,
              },
            })
            return
          }
        }
        void Promise.resolve(input.skill.executeTool(request.call))
          .then(async (outcome) => {
            const execution = isToolExecutionSuspension(outcome) ? await outcome.result : outcome
            await bridge.toolResult({
              documentId: input.documentId,
              generation: input.generation,
              callId: request.call.id,
              execution,
              ...(snapshot ? { snapshotBefore: snapshot.id } : {}),
            })
          })
          .catch(() => {
            snapshots.delete(request.call.id)
            return bridge.toolResult({
              documentId: input.documentId,
              generation: input.generation,
              callId: request.call.id,
              execution: { output: 'tool_execution_failed', summary: 'Tool failed', isError: true },
            })
          })
      })
      const session: EnhancedRuntimeClientSession = Object.freeze({
        async start(turn: {
          readonly text: string
          readonly images?: readonly import('@wiswork/agent-core').AgentImage[]
        }) {
          if (ended) throw new Error('enhanced_session_closed')
          await registration
          const status = await bridge.status()
          if (status.activeAgentRuntime !== 'enhanced' || status.documentId !== input.documentId) {
            throw new Error('enhanced_document_unavailable')
          }
          if (turn.images?.length) throw new Error('enhanced_images_unavailable')
          const context = input.skill.buildContext?.() ?? ''
          await bridge.startTurn({
            documentId: input.documentId,
            text: context
              ? `${turn.text}\n\nAuthoritative current document context:\n${context}`
              : turn.text,
          })
        },
        cancel: () => (ended ? Promise.resolve() : bridge.cancelTurn(input.documentId)),
        subscribe(listener: (event: EnhancedSessionEvent) => void) {
          if (ended) return () => undefined
          listeners.add(listener)
          return () => listeners.delete(listener)
        },
        async close() {
          if (ended) return
          ended = true
          unsubscribe()
          unsubscribeTools()
          snapshots.clear()
          listeners.clear()
          sessions.delete(session)
          await registration.catch(() => undefined)
          await bridge.cancelTurn(input.documentId).catch(() => undefined)
          await bridge.unregister(input.documentId, input.generation).catch(() => undefined)
        },
      })
      sessions.add(session)
      return session
    },
    async close() {
      if (closed) return
      closed = true
      await Promise.all([...sessions].map((session) => session.close()))
    },
  })
}
