import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import {
  parsePresentationDeliveryReport,
  presentationDeliveryMarkdown,
} from '@wiswork/pptx-engine/presentation-delivery-report'
import { parsePresentationIssueActionInput } from '@wiswork/project-store/presentation-issue'
import type { PresentationGenerationOptions } from './presentation-generation.js'
import type { InMemoryVfs } from '../shared/vfs.js'

const names = [
  'read_presentation_delivery_report',
  'record_presentation_issue_action',
  'export_presentation_delivery_report',
]
const validId = (value: unknown): value is string =>
  typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
export function createPresentationEvidenceDeliverySkill(
  options: Pick<PresentationGenerationOptions, 'request' | 'available' | 'documentId'> & {
    vfs: InMemoryVfs
  },
): AgentSkill & { clear(): void } {
  let epoch = 0
  const tools: AgentToolDef[] = names.map((name) => ({
    name,
    description:
      'Read frozen presentation evidence, record an issue disposition with a reason, or export full JSON and Markdown to session attachments. Arithmetic reproduction is not fact verification; explanations do not close findings. No host or QA changes.',
    inputSchema: {
      type: 'object',
      properties: {
        project_id: { type: 'string' },
        request_id: { type: 'string' },
        ...(name === names[1]
          ? {
              expected_revision: { type: 'integer', minimum: 0 },
              action: {
                type: 'object',
                properties: {
                  actionId: { type: 'string' },
                  issueId: { type: 'string' },
                  issueDigest: { type: 'string' },
                  state: { type: 'string', enum: ['open', 'deferred', 'explained'] },
                  note: { type: 'string' },
                },
                required: ['actionId', 'issueId', 'issueDigest', 'state', 'note'],
                additionalProperties: false,
              },
            }
          : {}),
      },
      required: [
        'project_id',
        'request_id',
        ...(name === names[1] ? ['expected_revision', 'action'] : []),
      ],
      additionalProperties: false,
    },
  }))
  return {
    id: 'office-presentation-evidence-delivery',
    systemPrompt:
      'Delivery reports describe frozen content evidence only. Arithmetic reproduction verifies calculation only. Sources, timeliness, host QA and Office round trip remain unverified. Keep unresolved issues and stale dispositions visible. Recording an explanation never proves a claim.',
    get tools() {
      return options.available() ? tools : []
    },
    clear() {
      epoch++
    },
    async executeTool(call, signal) {
      const captured = epoch
      const check = () => {
        if (captured !== epoch || signal?.aborted) throw new Error('cancelled')
        if (!options.available()) throw new Error('presentation_unavailable')
      }
      try {
        check()
        const input = call.input
        const recording = call.name === names[1]
        if (
          !names.includes(call.name) ||
          call.inputError ||
          call.truncated ||
          !validId(input.project_id) ||
          !validId(input.request_id) ||
          Object.keys(input).some(
            (key) =>
              ![
                'project_id',
                'request_id',
                ...(recording ? ['expected_revision', 'action'] : []),
              ].includes(key),
          )
        )
          throw new Error('invalid_tool_input')
        if (
          recording &&
          (!Number.isSafeInteger(input.expected_revision) || Number(input.expected_revision) < 0)
        )
          throw new Error('invalid_tool_input')
        const action = recording ? parsePresentationIssueActionInput(input.action) : undefined
        const documentId = await options.documentId()
        check()
        const response = await options.request(
          {
            operation: recording ? 'production_record_issue_action' : 'production_delivery_report',
            documentId,
            projectId: input.project_id,
            requestId: input.request_id,
            ...(recording ? { expectedRevision: input.expected_revision, action } : {}),
          },
          signal,
        )
        check()
        if (!response.ok) throw new Error('presentation_service_unavailable')
        const text = await response.text()
        check()
        if (new TextEncoder().encode(text).byteLength > 8 * 1024 * 1024)
          throw new Error('presentation_response_invalid')
        const value = JSON.parse(text)
        if (typeof value?.error === 'string' && /^[a-z_]{1,80}$/.test(value.error))
          throw new Error(
            value.error === 'invalid_request'
              ? 'presentation_upgrade_required'
              : `presentation_${value.error}`,
          )
        const report = parsePresentationDeliveryReport(value)
        if (
          report.documentId !== documentId ||
          report.projectId !== input.project_id ||
          report.requestId !== input.request_id
        )
          throw new Error('presentation_response_invalid')
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed')
        check()
        if (call.name === names[2]) {
          const prefix = `/home/user/presentation-evidence-${crypto.randomUUID()}`
          const paths = [`${prefix}.json`, `${prefix}.md`]
          const markdown = presentationDeliveryMarkdown(report)
          check()
          options.vfs.writeBatch([
            [paths[0]!, JSON.stringify(report, null, 2)],
            [paths[1]!, markdown],
          ])
          return {
            output: JSON.stringify({ paths }),
            mutated: false,
            summary: '完整证据报告已保存到会话附件；尚未核验来源或验收 Office',
          }
        }
        return {
          output: JSON.stringify(report),
          mutated: false,
          summary: '已读取证据与处置历史；算术复现不代表事实通过',
        }
      } catch (error) {
        const code = error instanceof Error ? error.message : ''
        return {
          output: /^(presentation_[a-z_]{1,80}|vfs_[a-z_]+|cancelled|invalid_tool_input)$/.test(
            code,
          )
            ? code
            : 'presentation_response_invalid',
          isError: true,
          mutated: false,
          summary: '证据操作未完成；PC 可能已保存处置，请刷新确认',
        }
      }
    },
  }
}
