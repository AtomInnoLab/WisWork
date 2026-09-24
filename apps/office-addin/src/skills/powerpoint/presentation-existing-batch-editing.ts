import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import type { PresentationBaselineSkill } from './presentation-baseline.js'
import type { PresentationBaselineAdapter } from './browser-presentation-baseline-adapter.js'
import {
  validatePowerPointPageScreenshot,
  type PowerPointAdapter,
  type PresentationPageGeometry,
} from './browser-powerpoint-adapter.js'
import {
  validatePresentationExistingBatch,
  type ExistingBatchOperation,
  type PresentationExistingBatch,
} from './presentation-existing-batch.js'

interface Options {
  baseline: PresentationBaselineSkill
  baselineAdapter: PresentationBaselineAdapter
  adapter: PowerPointAdapter
  proposals: StructuredProposalController
  documentId(): Promise<string>
  readExistingBatch(id: string): PresentationExistingBatch | undefined
  writeExistingBatch(
    record: PresentationExistingBatch,
    expected: PresentationExistingBatch | undefined,
  ): Promise<void>
}

const tools: AgentToolDef[] = [
  {
    name: 'edit_existing_presentation_batch',
    description:
      'Propose 2–8 ordered native text/geometry edits from one fresh scoped baseline. Confirmation creates one durable batch savepoint before host writes. Each step is read back; interrupted writes require explicit recovery.',
    inputSchema: {
      type: 'object',
      properties: {
        baseline_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
        intent: { type: 'string', minLength: 1, maxLength: 300 },
        preserved: {
          type: 'array',
          maxItems: 20,
          items: { type: 'string', minLength: 1, maxLength: 300 },
        },
        validation: {
          type: 'array',
          maxItems: 20,
          items: { type: 'string', minLength: 1, maxLength: 300 },
        },
        risk: { type: 'string', enum: ['medium', 'high'] },
        operations: {
          type: 'array',
          minItems: 2,
          maxItems: 8,
          items: {
            type: 'object',
            properties: {
              slide_id: { type: 'string', minLength: 1, maxLength: 256 },
              shape_id: { type: 'string', minLength: 1, maxLength: 256 },
              kind: { type: 'string', enum: ['text', 'geometry'] },
              text: { type: 'string', maxLength: 12000 },
              geometry: {
                type: 'object',
                properties: Object.fromEntries(
                  ['left', 'top', 'width', 'height'].map((k) => [k, { type: 'number' }]),
                ),
                required: ['left', 'top', 'width', 'height'],
                additionalProperties: false,
              },
            },
            required: ['slide_id', 'shape_id', 'kind'],
            additionalProperties: false,
          },
        },
      },
      required: ['baseline_id', 'intent', 'preserved', 'validation', 'risk', 'operations'],
      additionalProperties: false,
    },
  },
  ...(['inspect', 'resume', 'undo'] as const).map((action): AgentToolDef => ({
    name: `${action}_existing_presentation_batch`,
    description:
      action === 'inspect'
        ? 'Read every saved target and classify batch recovery without writing.'
        : 'Propose confirmed, stepwise batch recovery. Exact host values are checked before each write; ambiguous values stop without replay.',
    inputSchema: {
      type: 'object',
      properties: { change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' } },
      required: ['change_id'],
      additionalProperties: false,
    },
  })),
  {
    name: 'capture_existing_presentation_batch_page',
    description:
      'Capture one affected native page after matching all batch targets against a terminal saved state. Returns a screenshot for human review; does not mark visual QA passed.',
    inputSchema: {
      type: 'object',
      properties: {
        change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
        slide_id: { type: 'string', minLength: 1, maxLength: 256 },
      },
      required: ['change_id', 'slide_id'],
      additionalProperties: false,
    },
  },
  {
    name: 'record_existing_presentation_batch_page_review',
    description:
      'Persist a historical visual assessment of this session’s captured batch page only after fresh screenshot digest and target readback match. Does not certify current or whole-deck QA.',
    inputSchema: {
      type: 'object',
      properties: {
        change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
        slide_id: { type: 'string', minLength: 1, maxLength: 256 },
        screenshot_digest: { type: 'string', pattern: '^[a-f0-9]{64}$' },
        status: { type: 'string', enum: ['pass', 'fail'] },
        notes: { type: 'string', maxLength: 2000 },
      },
      required: ['change_id', 'slide_id', 'screenshot_digest', 'status', 'notes'],
      additionalProperties: false,
    },
  },
]
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const geometry = (g: PresentationPageGeometry) => ({
  left: g.left,
  top: g.top,
  width: g.width,
  height: g.height,
})
const matches = (a: unknown, b: unknown) =>
  typeof a === 'string' || typeof b === 'string'
    ? a === b
    : !!a &&
      !!b &&
      typeof a === 'object' &&
      typeof b === 'object' &&
      (['left', 'top', 'width', 'height'] as const).every(
        (k) =>
          Math.abs((a as PresentationPageGeometry)[k] - (b as PresentationPageGeometry)[k]) <= 0.01,
      )
const output = (v: unknown) => {
  const s = JSON.stringify(v)
  if (new TextEncoder().encode(s).byteLength > 256 * 1024)
    throw new Error('presentation_existing_batch_output_limit')
  return s
}
async function digest(value: string) {
  return Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(value))),
    (b) => b.toString(16).padStart(2, '0'),
  ).join('')
}

export function createPresentationExistingBatchEditingSkill(
  options: Options,
): AgentSkill & { clear(): void; beginMutation(): void; endMutation(): void } {
  let epoch = 0,
    qaEpoch = 0,
    mutating = 0
  const captures = new Map<
    string,
    { record: string; screenshotDigest: string; capturedAt: string; epoch: number }
  >()
  return {
    id: 'presentation-existing-batch-editing',
    tools,
    systemPrompt:
      'For two or more existing-deck text/geometry changes, read_presentation_baseline then edit_existing_presentation_batch. Use exact native IDs. A batch is ordered and recoverable, not atomic. After confirmed writes, capture_existing_presentation_batch_page for every affected page, visually review, then record_existing_presentation_batch_page_review using its screenshot_digest. Historical reviews do not certify current or whole-deck QA. If interrupted, inspect then resume or undo. Never replay ambiguous host values.',
    clear() {
      epoch++
      qaEpoch++
      captures.clear()
    },
    beginMutation() {
      mutating++
      qaEpoch++
      captures.clear()
    },
    endMutation() {
      mutating = Math.max(0, mutating - 1)
    },
    async executeTool(call, signal) {
      const token = epoch
      const active = () => {
        if (signal?.aborted || epoch !== token) throw new Error('cancelled')
      }
      try {
        active()
        const tool = tools.find((t) => t.name === call.name)
        if (!tool || call.inputError || call.truncated) throw new Error('invalid_tool_input')
        const props = tool.inputSchema as {
          properties: Record<string, unknown>
          required: string[]
        }
        if (
          Object.keys(call.input).some((k) => !Object.hasOwn(props.properties, k)) ||
          props.required.some((k) => !Object.hasOwn(call.input, k))
        )
          throw new Error('invalid_tool_input')
        const creating = call.name === 'edit_existing_presentation_batch'
        const documentId = await options.documentId()
        active()
        const current = async () => {
          active()
          if ((await options.documentId()) !== documentId)
            throw new Error('presentation_document_changed')
          active()
        }
        let record: PresentationExistingBatch
        let initialBaseline: ReturnType<PresentationBaselineSkill['snapshot']>
        if (creating) {
          const i = call.input
          if (
            typeof i.baseline_id !== 'string' ||
            !/^[A-Za-z0-9_-]{1,128}$/.test(i.baseline_id) ||
            typeof i.intent !== 'string' ||
            !i.intent.length ||
            i.intent.length > 300 ||
            !Array.isArray(i.operations) ||
            i.operations.length < 2 ||
            i.operations.length > 8 ||
            !Array.isArray(i.preserved) ||
            !Array.isArray(i.validation) ||
            !['medium', 'high'].includes(i.risk as string)
          )
            throw new Error('invalid_tool_input')
          const baseline = options.baseline.snapshot(i.baseline_id)
          if (!baseline || baseline.documentId !== documentId)
            throw new Error('presentation_baseline_changed')
          initialBaseline = structuredClone(baseline)
          const check = await options.baseline.executeTool(
            {
              id: 'batch-check',
              name: 'check_presentation_baseline',
              input: { baseline_id: baseline.baselineId },
            },
            signal,
          )
          await current()
          if (
            check.isError ||
            !JSON.parse(check.output).unchanged ||
            !same(baseline, options.baseline.snapshot(baseline.baselineId))
          )
            throw new Error('presentation_baseline_changed')
          const operations: ExistingBatchOperation[] = i.operations.map((raw: unknown) => {
            if (!raw || typeof raw !== 'object' || Array.isArray(raw))
              throw new Error('invalid_tool_input')
            const op = raw as Record<string, unknown>
            if (
              Object.keys(op).some(
                (k) => !['slide_id', 'shape_id', 'kind', 'text', 'geometry'].includes(k),
              ) ||
              typeof op.slide_id !== 'string' ||
              typeof op.shape_id !== 'string' ||
              !baseline.scope.slideIds.includes(op.slide_id) ||
              (baseline.scope.shapeIds && !baseline.scope.shapeIds.includes(op.shape_id))
            )
              throw new Error('presentation_existing_scope_mismatch')
            const shape = baseline.pages
              .find((p) => p.slideId === op.slide_id)
              ?.shapes.find((s) => s.id === op.shape_id)
            if (!shape) throw new Error('presentation_existing_target_unsupported')
            const base = { hostSlideId: op.slide_id, shapeId: op.shape_id, shapeType: shape.type }
            if (
              op.kind === 'text' &&
              typeof op.text === 'string' &&
              shape.text !== undefined &&
              op.geometry === undefined
            )
              return { ...base, kind: 'text', before: shape.text, after: op.text }
            if (
              op.kind === 'geometry' &&
              op.text === undefined &&
              op.geometry &&
              typeof op.geometry === 'object'
            )
              return {
                ...base,
                kind: 'geometry',
                before: geometry(shape),
                after: op.geometry as PresentationPageGeometry,
              }
            throw new Error('invalid_tool_input')
          })
          record = {
            version: 1,
            changeId: crypto.randomUUID(),
            documentId,
            baselineId: baseline.baselineId,
            baselineDigest: baseline.contentDigest,
            scope: {
              slideIds: [...baseline.scope.slideIds],
              ...(baseline.scope.shapeIds ? { shapeIds: [...baseline.scope.shapeIds] } : {}),
            },
            intent: i.intent,
            preserved: i.preserved,
            validation: i.validation,
            risk: i.risk,
            operations,
            state: 'applying',
            cursor: 0,
            reviewCapacity: true,
          } as PresentationExistingBatch
          if (!validatePresentationExistingBatch(record)) throw new Error('invalid_tool_input')
        } else {
          if (
            typeof call.input.change_id !== 'string' ||
            !/^[A-Za-z0-9_-]{1,128}$/.test(call.input.change_id)
          )
            throw new Error('invalid_tool_input')
          const saved = options.readExistingBatch(call.input.change_id)
          if (
            !saved ||
            !validatePresentationExistingBatch(saved) ||
            saved.documentId !== documentId
          )
            throw new Error('presentation_existing_batch_missing')
          record = structuredClone(saved)
        }
        let expected = creating ? undefined : structuredClone(record)
        const saved = () => {
          if (!same(options.readExistingBatch(record.changeId), expected))
            throw new Error('presentation_existing_batch_stale')
        }
        const value = async (op: ExistingBatchOperation) => {
          await current()
          saved()
          const page = await options.baselineAdapter.readPage(op.hostSlideId, signal)
          await current()
          saved()
          const shape = page.shapes.find((s) => s.id === op.shapeId)
          if (page.slideId !== op.hostSlideId || !shape || shape.type !== op.shapeType)
            throw new Error('presentation_existing_target_changed')
          return op.kind === 'text' ? shape.text : geometry(shape)
        }
        const store = async (next: PresentationExistingBatch) => {
          await current()
          saved()
          await options.writeExistingBatch(next, expected)
          await current()
          if (!same(options.readExistingBatch(next.changeId), next))
            throw new Error('office_state_uncertain')
          record = structuredClone(next)
          expected = structuredClone(next)
        }
        const classify = async () => {
          const values = []
          for (const op of record.operations) {
            const v = await value(op)
            values.push(
              matches(v, op.before) ? 'before' : matches(v, op.after) ? 'after' : 'conflict',
            )
          }
          return values
        }
        const values = await classify()
        if (
          call.name === 'capture_existing_presentation_batch_page' ||
          call.name === 'record_existing_presentation_batch_page_review'
        ) {
          const checkpoint = qaEpoch
          const qaCheck = () => {
            active()
            if (mutating || checkpoint !== qaEpoch)
              throw new Error('presentation_existing_batch_qa_stale')
          }
          qaCheck()
          if (
            typeof call.input.slide_id !== 'string' ||
            !record.operations.some((op) => op.hostSlideId === call.input.slide_id) ||
            !['applied', 'undone'].includes(record.state) ||
            !values.every((v) => v === (record.state === 'applied' ? 'after' : 'before'))
          )
            throw new Error('presentation_existing_batch_conflict')
          const key = JSON.stringify([record.changeId, call.input.slide_id])
          const prior = captures.get(key)
          if (
            call.name === 'record_existing_presentation_batch_page_review' &&
            (!prior ||
              prior.record !== JSON.stringify(record) ||
              prior.epoch !== qaEpoch ||
              call.input.screenshot_digest !== prior.screenshotDigest ||
              !['pass', 'fail'].includes(call.input.status as string) ||
              typeof call.input.notes !== 'string' ||
              call.input.notes.length > 2000)
          )
            throw new Error('presentation_existing_batch_qa_stale')
          if (!options.adapter.inspectPresentationPage) throw new Error('office_api_unsupported')
          const shot = await options.adapter.inspectPresentationPage(call.input.slide_id, signal)
          qaCheck()
          await current()
          saved()
          if (
            shot.slideId !== call.input.slide_id ||
            shot.shapesTruncated ||
            shot.screenshot.mime !== 'image/png'
          )
            throw new Error('office_read_failed')
          validatePowerPointPageScreenshot(shot.screenshot.base64)
          const screenshotDigest = await digest(shot.screenshot.base64)
          qaCheck()
          if (!same(values, await classify()))
            throw new Error('presentation_existing_batch_conflict')
          qaCheck()
          if (call.name === 'record_existing_presentation_batch_page_review') {
            if (screenshotDigest !== prior!.screenshotDigest)
              throw new Error('presentation_existing_batch_qa_stale')
            const review = {
              hostSlideId: shot.slideId,
              screenshotDigest,
              capturedAt: prior!.capturedAt,
              reviewedAt: new Date().toISOString(),
              status: call.input.status as 'pass' | 'fail',
              notes: call.input.notes as string,
            }
            const reviews = [
              ...(record.reviews ?? []).filter((v) => v.hostSlideId !== shot.slideId),
              review,
            ]
            await store({ ...record, reviews })
            qaCheck()
            captures.delete(key)
            return {
              output: output({
                changeId: record.changeId,
                hostSlideId: shot.slideId,
                historicalReview: review,
                currentScreenshotMatched: true,
                wholeDeckQaPassed: false,
              }),
              mutated: false,
              summary: '已保存该批量变更页面的历史视觉复核结果',
            }
          }
          captures.set(key, {
            record: JSON.stringify(record),
            screenshotDigest,
            capturedAt: new Date().toISOString(),
            epoch: qaEpoch,
          })
          return {
            output: output({
              changeId: record.changeId,
              hostSlideId: shot.slideId,
              state: record.state,
              screenshotDigest,
              structure: {
                overflows: shot.overflows,
                overlaps: shot.overlaps,
                overlapsTruncated: shot.overlapsTruncated,
              },
              qaPassed: false,
            }),
            modelContent: [{ type: 'image', image: shot.screenshot }],
            display: {
              kind: 'images',
              items: [{ url: `data:image/png;base64,${shot.screenshot.base64}` }],
            },
            mutated: false,
            summary: '已采集批量变更页面截图，等待视觉复核',
          }
        }
        const freshBaseline = async () => {
          if (!creating || !initialBaseline) return true
          if (!same(initialBaseline, options.baseline.snapshot(initialBaseline.baselineId)))
            return false
          const check = await options.baseline.executeTool(
            {
              id: 'batch-confirm-check',
              name: 'check_presentation_baseline',
              input: { baseline_id: initialBaseline.baselineId },
            },
            signal,
          )
          await current()
          return (
            !check.isError &&
            JSON.parse(check.output).unchanged &&
            same(initialBaseline, options.baseline.snapshot(initialBaseline.baselineId))
          )
        }
        if (call.name === 'inspect_existing_presentation_batch')
          return {
            output: output({
              changeId: record.changeId,
              state: record.state,
              cursor: record.cursor,
              values,
              currentHostVerified: true,
              qaPassed: false,
            }),
            mutated: false,
            summary: '已核对批量变更各目标当前值',
          }
        if (call.name === 'undo_existing_presentation_batch' && record.state !== 'applied')
          throw new Error('presentation_existing_batch_state_invalid')
        if (
          call.name === 'resume_existing_presentation_batch' &&
          !['applying', 'undoing'].includes(record.state)
        )
          throw new Error('presentation_existing_batch_state_invalid')
        const reverse =
          call.name === 'undo_existing_presentation_batch' || record.state === 'undoing'
        const allowed = record.operations.every(
          (_, n) =>
            values[n] === (n < record.cursor ? 'after' : 'before') ||
            (n === (reverse ? record.cursor - 1 : record.cursor) &&
              values[n] === (reverse ? 'before' : 'after')),
        )
        if (!allowed) throw new Error('presentation_existing_batch_conflict')
        const proposal = options.proposals.propose({
          operation: call.name,
          toolName: call.name,
          title: reverse ? '撤销现稿批量修改' : record.intent,
          preview: {
            scope: record.scope,
            operations: record.operations.map((op) => ({
              slideId: op.hostSlideId,
              shapeId: op.shapeId,
              kind: op.kind,
            })),
            cursor: record.cursor,
            risk: record.risk,
            preserved: record.preserved,
            validation: record.validation,
            atomic: false,
            textFormatting: '文字撤销仅恢复内容，不恢复全部富文本格式',
          },
          impact: {
            host: 'powerpoint',
            targets: [...new Set(record.operations.map((op) => op.hostSlideId))],
            count: record.operations.length,
          },
          fingerprint: selectionFingerprint(output(record)),
          before: record.operations.map((op) => op.before),
          after: record.operations.map((op) => op.after),
          validate: async () => {
            try {
              return (await freshBaseline()) && same(values, await classify())
            } catch {
              return false
            }
          },
          execute: async () => {
            if (!(await freshBaseline())) throw new Error('proposal_stale')
            if (!same(values, await classify())) throw new Error('proposal_stale')
            if (creating) await store(record)
            if (call.name === 'undo_existing_presentation_batch') {
              const { reviews: _reviews, ...r } = record
              await store({ ...r, state: 'undoing' })
            }
            while (record.state === 'applying' || record.state === 'undoing') {
              const back = record.state === 'undoing'
              const index = back ? record.cursor - 1 : record.cursor
              const op = record.operations[index]
              const source = back ? op.after : op.before,
                target = back ? op.before : op.after
              const v = await value(op)
              if (!matches(v, target)) {
                if (!matches(v, source)) throw new Error('presentation_existing_batch_conflict')
                if (op.kind === 'text')
                  await options.adapter.editPresentationPageText!(
                    op.hostSlideId,
                    op.shapeId,
                    target as string,
                    source as string,
                    signal,
                  )
                else
                  await options.adapter.editPresentationPageGeometry!(
                    op.hostSlideId,
                    op.shapeId,
                    target as PresentationPageGeometry,
                    source as PresentationPageGeometry,
                    signal,
                  )
              }
              if (!matches(await value(op), target)) throw new Error('office_verify_failed')
              const cursor = record.cursor + (back ? -1 : 1)
              await store({
                ...record,
                cursor,
                state: back
                  ? cursor === 0
                    ? 'undone'
                    : 'undoing'
                  : cursor === record.operations.length
                    ? 'applied'
                    : 'applying',
              })
            }
          },
          verify: async () => {
            if (
              record.state !== (reverse ? 'undone' : 'applied') ||
              !(await classify()).every((v) => v === (reverse ? 'before' : 'after'))
            )
              throw new Error('office_verify_failed')
          },
        })
        return {
          output: output({
            proposalId: proposal.id,
            changeId: record.changeId,
            status: 'awaiting_confirmation',
            operations: record.operations.length,
          }),
          mutated: false,
          summary: '已准备现稿批量变更提案，等待确认',
        }
      } catch (error) {
        const code = error instanceof Error ? error.message : ''
        return {
          output:
            /^(presentation_[a-z_]+|office_[a-z_]+|invalid_tool_input|cancelled|proposal_stale)$/.test(
              code,
            )
              ? code
              : 'presentation_existing_batch_operation_failed',
          isError: true,
          mutated: false,
          summary: '现稿批量操作未完成；未自动重试',
        }
      }
    },
  }
}
