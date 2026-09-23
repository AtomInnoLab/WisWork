import { createPresentationPageDeliverySkill } from './presentation-page-delivery.js'
import type { AgentSkill } from '@wiswork/agent-core'
import {
  selectionFingerprint,
  type StructuredProposalController,
} from '../../agent/proposal-controller.js'
import type { PresentationImportAdapter, PresentationImportReceipt } from './presentation-import.js'

export interface CompiledPresentationArtifact {
  documentId: string
  projectId: string
  requestId: string
  pptxBase64: string
  slideCount: number
  pages?: { id: string; title: string; sourceSlideId: string }[]
}
export interface PresentationImportRecord {
  state: 'pending' | 'complete'
  documentId: string
  slideIds?: string[]
  checkpoint?: PresentationImportCheckpoint
}
export interface PresentationImportCheckpoint {
  version: 1
  artifactDigest: string
  sourceSlideIds: string[]
  baselineSlideIds: string[]
  completed: { sourceSlideId: string; slideId: string }[]
  inFlight?: { sourceSlideId: string }
}
export interface PresentationDeliveryOptions {
  adapter: PresentationImportAdapter
  proposals: StructuredProposalController
  available(): boolean
  artifact(projectId?: string): CompiledPresentationArtifact | undefined
  documentId(): Promise<string>
  readReceipt(key: string): PresentationImportRecord | undefined
  writeReceipt(key: string, record: PresentationImportRecord | undefined): Promise<void>
}
const tool = {
  name: 'import_generated_presentation',
  description:
    'Propose appending a compiled or restored presentation to the current PowerPoint document. Existing slides are preserved. User confirms once; a durable receipt prevents duplicate import. This verifies page IDs/count, not rendered quality.',
  inputSchema: {
    type: 'object',
    properties: { project_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' } },
    additionalProperties: false,
  },
}

export function createPresentationDeliverySkill(options: PresentationDeliveryOptions): AgentSkill {
  const pages = createPresentationPageDeliverySkill(options)
  return {
    id: 'office-presentation-delivery',
    systemPrompt:
      'After compilation or restoration, use import_generated_presentation when the user wants the generated pages in the current PowerPoint. This appends pages with one confirmation and preserves existing content. After interruption, read_presentation_import_status to inspect saved page progress. A new confirmation resumes remaining pages only when prior pages and document order match. Do not retry an uncertain import; inspect the document first. After import use screenshot_slide and verify_slides for actual host review; do not equate page-count verification with complete QA.',
    get tools() {
      return options.available() && options.adapter.available() ? [tool, ...pages.tools] : []
    },
    async executeTool(call, signal) {
      if (
        call.name === 'read_presentation_import_status' ||
        (call.name === tool.name &&
          options.adapter.insertPage &&
          options.artifact(call.input.project_id as string | undefined)?.pages)
      )
        return pages.executeTool(call, signal)
      try {
        if (signal?.aborted) throw new Error('cancelled')
        if (!options.available() || !options.adapter.available())
          throw new Error('presentation_unavailable')
        if (
          call.name !== tool.name ||
          call.inputError ||
          call.truncated ||
          Object.keys(call.input).some((key) => key !== 'project_id') ||
          (call.input.project_id !== undefined &&
            (typeof call.input.project_id !== 'string' ||
              !/^[A-Za-z0-9_-]{1,128}$/.test(call.input.project_id)))
        )
          throw new Error('invalid_tool_input')
        const artifact = options.artifact(call.input.project_id as string | undefined)
        if (!artifact) throw new Error('presentation_restore_required')
        const documentId = await options.documentId()
        if (documentId !== artifact.documentId) throw new Error('presentation_document_changed')
        const isCurrent = () =>
          options.available() && options.artifact(artifact.projectId) === artifact
        const key = `${artifact.projectId}/${artifact.requestId}`
        const previous = options.readReceipt(key)
        if (previous) {
          if (previous.documentId !== documentId) throw new Error('presentation_document_changed')
          if (previous.state === 'pending') throw new Error('presentation_import_uncertain')
          return {
            output: JSON.stringify({ status: 'already_imported', slideIds: previous.slideIds }),
            mutated: false,
            summary: '这份文稿已导入，未重复插入页面',
          }
        }
        const before = await options.adapter.snapshot(signal)
        if (!isCurrent() || signal?.aborted) throw new Error('cancelled')
        let receipt: PresentationImportReceipt | undefined
        const proposal = options.proposals.propose({
          operation: tool.name,
          toolName: tool.name,
          title: `添加 ${artifact.slideCount} 页生成内容，保留现有页面`,
          preview: {
            project: artifact.projectId,
            pages: artifact.slideCount,
            existingPages: before.slideIds.length,
          },
          impact: {
            host: 'powerpoint',
            targets: before.slideIds.length ? [before.slideIds.at(-1)!] : ['presentation'],
            count: artifact.slideCount,
          },
          fingerprint: selectionFingerprint(`${documentId}:${before.fingerprint}:${key}`),
          validate: async (validationSignal) => {
            if (
              !isCurrent() ||
              validationSignal?.aborted ||
              (await options.documentId()) !== documentId ||
              options.readReceipt(key)
            )
              return false
            return (
              (await options.adapter.snapshot(validationSignal)).fingerprint === before.fingerprint
            )
          },
          execute: async (writeSignal) => {
            if (
              !isCurrent() ||
              writeSignal?.aborted ||
              (await options.documentId()) !== documentId ||
              options.readReceipt(key)
            )
              throw new Error('proposal_stale')
            // Reserve durably before Office can commit. An uncertain write must never be replayed.
            await options.writeReceipt(key, { state: 'pending', documentId })
            try {
              if (writeSignal?.aborted) throw new Error('cancelled')
              if ((await options.documentId()) !== documentId) throw new Error('proposal_stale')
              receipt = await options.adapter.insert(
                artifact.pptxBase64,
                artifact.slideCount,
                before,
                writeSignal,
              )
            } catch (error) {
              const code = error instanceof Error ? error.message : ''
              if (
                [
                  'cancelled',
                  'proposal_stale',
                  'office_api_unsupported',
                  'invalid_tool_input',
                ].includes(code)
              ) {
                // Adapter guarantees these errors occur before any write is queued.
                if ((await options.documentId()) === documentId)
                  await options.writeReceipt(key, undefined)
              }
              throw error
            }
          },
          verify: async (verificationSignal) => {
            if (
              !receipt ||
              (await options.documentId()) !== documentId ||
              !(await options.adapter.verify(receipt, before, verificationSignal))
            )
              throw new Error('office_state_uncertain')
            await options.writeReceipt(key, {
              state: 'complete',
              documentId,
              slideIds: receipt.slideIds,
            })
          },
        })
        return {
          output: JSON.stringify({ proposalId: proposal.id, status: 'awaiting_confirmation' }),
          mutated: false,
          summary: '已准备添加生成页面，等待确认',
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : ''
        const code = /^(presentation_[a-z_]+|office_[a-z_]+|invalid_tool_input|cancelled)$/.test(
          message,
        )
          ? message
          : 'presentation_operation_failed'
        return {
          output: code,
          isError: true,
          mutated: false,
          summary: '文稿导入未完成，现有内容已保留',
        }
      }
    },
  }
}
