import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import type {
  ProposalPostWriteEvidence,
  StructuredProposalController,
} from '../../agent/proposal-controller.js'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import type { PresentationBaselineSkill } from './presentation-baseline.js'
import {
  nativePlainTextEditable,
  type PresentationBaselineAdapter,
} from './browser-presentation-baseline-adapter.js'
import type { PresentationBaselinePage } from './browser-presentation-baseline-adapter.js'
import {
  nativeGeometryEditable,
  validatePowerPointPageScreenshot,
  type PowerPointAdapter,
  type PresentationPageGeometry,
} from './browser-powerpoint-adapter.js'
import {
  validatePresentationExistingBatch,
  type ExistingBatchOperation,
  type PresentationExistingBatch,
} from './presentation-existing-batch.js'
import { inspectPowerPointTableCellsPackage } from './presentation-complex-page-package.js'
import { presentationPackageDigest } from './powerpoint-package.js'
import {
  describePagePackageBackup,
  readChartPackageBackup,
  saveChartPackageBackup,
} from './presentation-chart-backup.js'

interface Options {
  baseline: PresentationBaselineSkill
  baselineAdapter: PresentationBaselineAdapter
  adapter: PowerPointAdapter
  proposals: StructuredProposalController
  documentId(): Promise<string>
  request(body: unknown, signal?: AbortSignal): Promise<Response>
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
      'Propose 2–8 ordered native text/geometry edits from one fresh scoped baseline. Confirmation backs up every affected original page to the paired PC before any host write. Each step is read back; interrupted writes require explicit recovery.',
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
  {
    name: 'edit_existing_presentation_table_batch',
    description:
      'Propose 2–8 ordered native text cell edits in one existing table. Confirmation backs up its original page to the paired PC before host writes; saves each verified step for recovery and undo.',
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
              row_index: { type: 'integer', minimum: 0, maximum: 19 },
              column_index: { type: 'integer', minimum: 0, maximum: 11 },
              text: { type: 'string', maxLength: 128 },
            },
            required: ['slide_id', 'shape_id', 'row_index', 'column_index', 'text'],
            additionalProperties: false,
          },
        },
      },
      required: ['baseline_id', 'intent', 'preserved', 'validation', 'risk', 'operations'],
      additionalProperties: false,
    },
  },
  ...(['inspect', 'resume', 'undo', 'release'] as const).map((action): AgentToolDef => ({
    name: `${action}_existing_presentation_batch`,
    description:
      action === 'inspect'
        ? 'Read every saved target and classify batch recovery without writing.'
        : action === 'release'
          ? 'After a batch has been fully undone, propose releasing its original page package backups from the paired PC. Requires separate confirmation.'
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
function preservedFields(page: PresentationBaselinePage, operations: ExistingBatchOperation[]) {
  const targets = new Set(
    operations.filter((op) => op.hostSlideId === page.slideId).map((op) => op.shapeId),
  )
  return JSON.stringify({
    slideId: page.slideId,
    masterId: page.masterId,
    layoutId: page.layoutId,
    shapes: page.shapes.filter((shape) => !targets.has(shape.id)),
  })
}
const targetKey = (slideId: string, shapeId: string) => JSON.stringify([slideId, shapeId])
function preservedTargetFields(
  shape: PresentationBaselinePage['shapes'][number],
  slideId: string,
  operations: ExistingBatchOperation[],
) {
  const planned = operations.filter((op) => op.hostSlideId === slideId && op.shapeId === shape.id)
  const value: Record<string, unknown> = { ...shape }
  if (planned.some((op) => op.kind === 'text' || op.kind === 'table_cell')) delete value.text
  if (planned.some((op) => op.kind === 'table_cell')) delete value.font
  if (planned.some((op) => op.kind === 'geometry'))
    for (const key of ['left', 'top', 'width', 'height']) delete value[key]
  return JSON.stringify(value)
}
async function pngDigest(base64: string) {
  const bytes = Uint8Array.from(atob(base64), (char) => char.charCodeAt(0))
  return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (b) =>
    b.toString(16).padStart(2, '0'),
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
      'For two or more existing-deck text/geometry changes, read_presentation_baseline then edit_existing_presentation_batch. For 2–8 cells in one native existing table use edit_existing_presentation_table_batch. Use exact native IDs. Geometry proposals support TextBox, GeometricShape, Image and Line only; use a page rebuild or dedicated validated operation for Chart, Group, SmartArt and placeholders. Whole-range text edits require TextBox or GeometricShape and determinate aggregate font fields; mixed or unknown formatting requires a dedicated validated operation. Confirmation requires all affected original page packages to be saved and read back on the paired PC before any host write. A batch is ordered and recoverable, not atomic; undo restores captured fields and does not replace the original page. After confirmed writes, capture_existing_presentation_batch_page for every affected page, visually review, then record_existing_presentation_batch_page_review using its screenshot_digest. Historical reviews do not certify current or whole-deck QA. If interrupted, inspect then resume or undo. Never replay ambiguous host values.',
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
        const tableBatch = call.name === 'edit_existing_presentation_table_batch'
        const creating = call.name === 'edit_existing_presentation_batch' || tableBatch
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
        const proposalPackages = new Map<string, string>()
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
          const tableOperations = async (): Promise<ExistingBatchOperation[]> => {
            const raw = i.operations as unknown[]
            const first = raw[0] as Record<string, unknown>
            if (
              !first ||
              typeof first.slide_id !== 'string' ||
              typeof first.shape_id !== 'string' ||
              !baseline.scope.slideIds.includes(first.slide_id) ||
              (baseline.scope.shapeIds && !baseline.scope.shapeIds.includes(first.shape_id))
            )
              throw new Error('presentation_existing_scope_mismatch')
            const shape = baseline.pages
              .find((page) => page.slideId === first.slide_id)
              ?.shapes.find((item) => item.id === first.shape_id)
            if (!shape || shape.type !== 'Table')
              throw new Error('presentation_existing_target_unsupported')
            const targets = raw.map((item) => {
              if (!item || typeof item !== 'object' || Array.isArray(item))
                throw new Error('invalid_tool_input')
              const op = item as Record<string, unknown>
              if (
                Object.keys(op).some(
                  (key) =>
                    !['slide_id', 'shape_id', 'row_index', 'column_index', 'text'].includes(key),
                ) ||
                op.slide_id !== first.slide_id ||
                op.shape_id !== first.shape_id ||
                !Number.isSafeInteger(op.row_index) ||
                !Number.isSafeInteger(op.column_index) ||
                (op.row_index as number) < 0 ||
                (op.row_index as number) > 19 ||
                (op.column_index as number) < 0 ||
                (op.column_index as number) > 11 ||
                typeof op.text !== 'string' ||
                op.text.length > 128
              )
                throw new Error('invalid_tool_input')
              return {
                rowIndex: op.row_index as number,
                columnIndex: op.column_index as number,
                after: op.text,
              }
            })
            if (
              !options.adapter.exportPresentationPagePackage ||
              !options.adapter.readPresentationTableCell ||
              !options.adapter.editPresentationTableCell
            )
              throw new Error('office_api_unsupported')
            const exported = await options.adapter.exportPresentationPagePackage(
              first.slide_id,
              signal,
            )
            if (
              exported.slideId !== first.slide_id ||
              !same(exported.slideIds, baseline.context.slideIds)
            )
              throw new Error('presentation_baseline_changed')
            const evidence = await inspectPowerPointTableCellsPackage(
              exported.base64,
              first.shape_id,
              targets,
              signal,
            )
            const digest = await presentationPackageDigest(exported.base64, signal)
            const repeated = await options.adapter.exportPresentationPagePackage(
              first.slide_id,
              signal,
            )
            if (
              repeated.slideId !== first.slide_id ||
              !same(repeated.slideIds, exported.slideIds) ||
              (await presentationPackageDigest(repeated.base64, signal)) !== digest
            )
              throw new Error('presentation_baseline_changed')
            const operations: ExistingBatchOperation[] = []
            for (const [index, target] of targets.entries()) {
              const host = await options.adapter.readPresentationTableCell(
                first.slide_id,
                first.shape_id,
                target.rowIndex,
                target.columnIndex,
                signal,
              )
              if (
                host.slideId !== first.slide_id ||
                host.shapeId !== first.shape_id ||
                host.rowIndex !== target.rowIndex ||
                host.columnIndex !== target.columnIndex ||
                host.text !== evidence.cells[index]?.text
              )
                throw new Error('presentation_baseline_changed')
              operations.push({
                hostSlideId: first.slide_id,
                shapeId: first.shape_id,
                shapeType: 'Table',
                kind: 'table_cell',
                rowIndex: target.rowIndex,
                columnIndex: target.columnIndex,
                tableStructureDigest: evidence.structureDigest,
                before: host.text,
                after: target.after,
              })
            }
            return operations
          }
          const operations: ExistingBatchOperation[] = tableBatch
            ? await tableOperations()
            : i.operations.map((raw: unknown) => {
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
                if (op.kind === 'geometry' && !nativeGeometryEditable(shape.type))
                  throw new Error('presentation_existing_target_unsupported')
                if (op.kind === 'text' && !nativePlainTextEditable(shape))
                  throw new Error('presentation_existing_target_unsupported')
                const base = {
                  hostSlideId: op.slide_id,
                  shapeId: op.shape_id,
                  shapeType: shape.type,
                }
                if (
                  op.kind === 'text' &&
                  typeof op.text === 'string' &&
                  nativePlainTextEditable(shape) &&
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
          const preservedPageDigests = Object.fromEntries(
            await Promise.all(
              [...new Set(operations.map((op) => op.hostSlideId))].map(async (slideId) => {
                const page = baseline.pages.find((item) => item.slideId === slideId)
                if (!page) throw new Error('presentation_baseline_changed')
                return [slideId, await digest(preservedFields(page, operations))] as const
              }),
            ),
          )
          const preservedTargetDigests = Object.fromEntries(
            await Promise.all(
              [...new Set(operations.map((op) => targetKey(op.hostSlideId, op.shapeId)))].map(
                async (key) => {
                  const [slideId, shapeId] = JSON.parse(key) as [string, string]
                  const shape = baseline.pages
                    .find((page) => page.slideId === slideId)
                    ?.shapes.find((item) => item.id === shapeId)
                  if (!shape) throw new Error('presentation_baseline_changed')
                  return [
                    key,
                    await digest(preservedTargetFields(shape, slideId, operations)),
                  ] as const
                },
              ),
            ),
          )
          if (!options.adapter.exportPresentationPagePackage)
            throw new Error('office_api_unsupported')
          const backups: NonNullable<PresentationExistingBatch['backups']> = []
          for (const slideId of new Set(operations.map((op) => op.hostSlideId))) {
            const first = await options.adapter.exportPresentationPagePackage(slideId, signal)
            const second = await options.adapter.exportPresentationPagePackage(slideId, signal)
            await current()
            if (
              first.slideId !== slideId ||
              second.slideId !== slideId ||
              !same(first.slideIds, baseline.context.slideIds) ||
              !same(second.slideIds, first.slideIds)
            )
              throw new Error('presentation_baseline_changed')
            const metadata = await describePagePackageBackup(first.base64, signal)
            if ((await presentationPackageDigest(second.base64, signal)) !== metadata.packageDigest)
              throw new Error('presentation_baseline_changed')
            proposalPackages.set(slideId, first.base64)
            backups.push({ hostSlideId: slideId, backupId: crypto.randomUUID(), ...metadata })
          }
          record = {
            version: 1,
            changeId: crypto.randomUUID(),
            documentId,
            baselineId: baseline.baselineId,
            baselineDigest: baseline.contentDigest,
            beforeSlideIds: [...baseline.context.slideIds],
            backups,
            scope: {
              slideIds: [...baseline.scope.slideIds],
              ...(baseline.scope.shapeIds ? { shapeIds: [...baseline.scope.shapeIds] } : {}),
            },
            intent: i.intent,
            preserved: i.preserved,
            preservedPageDigests,
            preservedTargetDigests,
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
          const preservedDigest = record.preservedPageDigests?.[op.hostSlideId]
          if (
            preservedDigest &&
            (await digest(preservedFields(page, record.operations))) !== preservedDigest
          )
            throw new Error('presentation_existing_preserved_changed')
          const targetDigest =
            record.preservedTargetDigests?.[targetKey(op.hostSlideId, op.shapeId)]
          if (
            targetDigest &&
            (await digest(preservedTargetFields(shape, op.hostSlideId, record.operations))) !==
              targetDigest
          )
            throw new Error('presentation_existing_preserved_changed')
          if (op.kind === 'table_cell') {
            if (
              !options.adapter.readPresentationTableCell ||
              !options.adapter.exportPresentationPagePackage
            )
              throw new Error('office_api_unsupported')
            const positions = record.operations
              .filter(
                (item): item is Extract<ExistingBatchOperation, { kind: 'table_cell' }> =>
                  item.kind === 'table_cell',
              )
              .map((item) => ({ rowIndex: item.rowIndex, columnIndex: item.columnIndex }))
            const native = await options.adapter.readPresentationTableCell(
              op.hostSlideId,
              op.shapeId,
              op.rowIndex,
              op.columnIndex,
              signal,
            )
            const first = await options.adapter.exportPresentationPagePackage(
              op.hostSlideId,
              signal,
            )
            if (first.slideId !== op.hostSlideId)
              throw new Error('presentation_existing_target_changed')
            const evidence = await inspectPowerPointTableCellsPackage(
              first.base64,
              op.shapeId,
              positions,
              signal,
            )
            if (evidence.structureDigest !== op.tableStructureDigest)
              throw new Error('presentation_existing_target_changed')
            const currentCell = evidence.cells.find(
              (cell) => cell.rowIndex === op.rowIndex && cell.columnIndex === op.columnIndex,
            )
            const packageDigest = await presentationPackageDigest(first.base64, signal)
            const repeated = await options.adapter.exportPresentationPagePackage(
              op.hostSlideId,
              signal,
            )
            await current()
            saved()
            if (
              repeated.slideId !== op.hostSlideId ||
              !same(repeated.slideIds, first.slideIds) ||
              (await presentationPackageDigest(repeated.base64, signal)) !== packageDigest ||
              native.slideId !== op.hostSlideId ||
              native.shapeId !== op.shapeId ||
              native.rowIndex !== op.rowIndex ||
              native.columnIndex !== op.columnIndex ||
              native.text !== currentCell?.text
            )
              throw new Error('presentation_existing_target_changed')
            return native.text
          }
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
        if (call.name === 'release_existing_presentation_batch') {
          const originalPackagesRestored = async (signal?: AbortSignal) => {
            if (!options.adapter.exportPresentationPagePackage || !record.beforeSlideIds)
              return false
            for (const backup of record.backups ?? []) {
              const exported = await options.adapter.exportPresentationPagePackage(
                backup.hostSlideId,
                signal,
              )
              if (
                exported.slideId !== backup.hostSlideId ||
                !same(exported.slideIds, record.beforeSlideIds) ||
                (await presentationPackageDigest(exported.base64, signal)) !== backup.packageDigest
              )
                return false
            }
            return true
          }
          if (
            record.state !== 'undone' ||
            !record.backups?.length ||
            !record.beforeSlideIds ||
            record.backupReleasedAt ||
            !values.every((value) => value === 'before')
          )
            throw new Error('presentation_existing_batch_state_invalid')
          if (!(await originalPackagesRestored(signal)))
            throw new Error('presentation_existing_batch_original_page_changed')
          const proposal = options.proposals.propose({
            operation: call.name,
            toolName: call.name,
            title: '释放已撤销批量修改的原页备份',
            preview: {
              changeId: record.changeId,
              pages: record.backups.map(({ hostSlideId }) => hostSlideId),
              backupCount: record.backups.length,
              originalPagePackages: '已核对原页包内容与写前保存点一致；释放后本机备份不可恢复',
            },
            impact: {
              host: 'powerpoint',
              targets: record.backups.map(({ backupId }) => `backup:${backupId}`),
              count: record.backups.length,
            },
            fingerprint: selectionFingerprint(output(record)),
            validate: async () => {
              try {
                return same(values, await classify()) && (await originalPackagesRestored())
              } catch {
                return false
              }
            },
            execute: async (writeSignal) => {
              if (!same(values, await classify()) || !(await originalPackagesRestored(writeSignal)))
                throw new Error('proposal_stale')
              for (const backup of record.backups!) {
                if (writeSignal?.aborted) throw new Error('cancelled')
                await current()
                saved()
                const response = await options.request(
                  {
                    operation: 'existing_page_backup_release',
                    documentId,
                    backupId: backup.backupId,
                    hostSlideId: backup.hostSlideId,
                    slideIds: record.beforeSlideIds,
                    sha256: backup.sha256,
                    sizeBytes: backup.sizeBytes,
                  },
                  writeSignal,
                )
                await current()
                if (!response.ok)
                  throw new Error('presentation_existing_batch_backup_release_failed')
                const receipt = (await response.json()) as Record<string, unknown>
                if (
                  receipt.status !== 'released' ||
                  receipt.documentId !== documentId ||
                  receipt.backupId !== backup.backupId ||
                  receipt.hostSlideId !== backup.hostSlideId ||
                  !same(receipt.slideIds, record.beforeSlideIds) ||
                  receipt.sha256 !== backup.sha256 ||
                  receipt.sizeBytes !== backup.sizeBytes
                )
                  throw new Error('presentation_existing_batch_backup_release_failed')
              }
              if (!same(values, await classify()))
                throw new Error('presentation_existing_batch_conflict')
              await store({ ...record, backupReleasedAt: new Date().toISOString() })
            },
            verify: async () => {
              saved()
              if (!record.backupReleasedAt) throw new Error('office_state_uncertain')
            },
          })
          return {
            output: output({
              proposalId: proposal.id,
              status: 'awaiting_confirmation',
              changeId: record.changeId,
            }),
            mutated: false,
            summary: '已准备释放撤销后的批量原页备份，等待确认',
          }
        }
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
        const ensureBackups = async (writeSignal?: AbortSignal) => {
          if (!record.backups || !record.beforeSlideIds) return // Legacy records predate package savepoints.
          if (!options.adapter.exportPresentationPagePackage)
            throw new Error('office_api_unsupported')
          for (let backup of record.backups) {
            await current()
            saved()
            const scope = {
              request: options.request,
              documentId,
              hostSlideId: backup.hostSlideId,
              slideIds: record.beforeSlideIds,
            }
            let ready = false
            try {
              await readChartPackageBackup(
                { ...scope, backup, expectedPackageDigest: backup.packageDigest },
                writeSignal,
              )
              ready = true
            } catch (error) {
              if (writeSignal?.aborted) throw error
            }
            if (!ready) {
              if (record.state !== 'applying' || record.cursor !== 0)
                throw new Error('presentation_existing_batch_backup_missing')
              const exported = await options.adapter.exportPresentationPagePackage(
                backup.hostSlideId,
                writeSignal,
              )
              await current()
              if (
                exported.slideId !== backup.hostSlideId ||
                !same(exported.slideIds, record.beforeSlideIds) ||
                (await presentationPackageDigest(exported.base64, writeSignal)) !==
                  backup.packageDigest
              )
                throw new Error('presentation_baseline_changed')
              const base64 = proposalPackages.get(backup.hostSlideId) ?? exported.base64
              const metadata = await describePagePackageBackup(base64, writeSignal)
              if (metadata.packageDigest !== backup.packageDigest)
                throw new Error('presentation_existing_batch_backup_missing')
              if (metadata.sha256 !== backup.sha256 || metadata.sizeBytes !== backup.sizeBytes) {
                const status = await options.request(
                  {
                    operation: 'existing_page_backup_status',
                    documentId,
                    backupId: backup.backupId,
                  },
                  writeSignal,
                )
                await current()
                const result = (await status.json()) as Record<string, unknown>
                if (status.ok || result.error !== 'not_found')
                  throw new Error('presentation_existing_batch_backup_missing')
                const replacement = { ...backup, ...metadata, backupId: crypto.randomUUID() }
                await store({
                  ...record,
                  backups: record.backups.map((item) =>
                    item.backupId === backup.backupId ? replacement : item,
                  ),
                })
                backup = replacement
              }
              const stored = await saveChartPackageBackup(
                { ...scope, base64, backupId: backup.backupId },
                writeSignal,
              )
              if (stored.sha256 !== backup.sha256 || stored.sizeBytes !== backup.sizeBytes)
                throw new Error('presentation_existing_batch_backup_missing')
              await readChartPackageBackup(
                { ...scope, backup, expectedPackageDigest: backup.packageDigest },
                writeSignal,
              )
            }
            await current()
            saved()
          }
          if (record.state === 'applying' && record.cursor === 0) {
            for (const backup of record.backups) {
              const exported = await options.adapter.exportPresentationPagePackage(
                backup.hostSlideId,
                writeSignal,
              )
              await current()
              if (
                exported.slideId !== backup.hostSlideId ||
                !same(exported.slideIds, record.beforeSlideIds) ||
                (await presentationPackageDigest(exported.base64, writeSignal)) !==
                  backup.packageDigest
              )
                throw new Error('presentation_baseline_changed')
            }
          }
        }
        if (call.name === 'inspect_existing_presentation_batch')
          return {
            output: output({
              changeId: record.changeId,
              state: record.state,
              cursor: record.cursor,
              values,
              currentHostVerified: true,
              packageSavepoints:
                record.backups?.map(({ hostSlideId, backupId }) => ({ hostSlideId, backupId })) ??
                [],
              packageSavepointsRequireReadbackBeforeWrite: Boolean(record.backups?.length),
              qaPassed: false,
            }),
            mutated: false,
            summary: '已核对批量变更各目标当前值',
          }
        if (
          call.name === 'undo_existing_presentation_batch' &&
          !['applied', 'applying'].includes(record.state)
        )
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
              values[n] === (reverse ? 'before' : 'after')) ||
            (reverse &&
              record.state === 'applying' &&
              n === record.cursor &&
              values[n] === 'after'),
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
              ...(op.kind === 'table_cell'
                ? { rowIndex: op.rowIndex, columnIndex: op.columnIndex }
                : {}),
            })),
            cursor: record.cursor,
            risk: record.risk,
            preserved: record.preserved,
            preservedFields: '已读取的受影响页非目标形状字段和目标形状未计划修改字段自动对照',
            validation: record.validation,
            originalPageBackup: record.backups
              ? `${record.backups.length} 个原页包写前保存到已配对的本机 PC，单页不超过 8 MiB`
              : '旧记录仅保存已读取字段',
            atomic: false,
            textFormatting: '文字与表格单元格撤销仅恢复内容，不恢复全部富文本格式',
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
          execute: async (writeSignal) => {
            if (!(await freshBaseline())) throw new Error('proposal_stale')
            if (!same(values, await classify())) throw new Error('proposal_stale')
            if (writeSignal?.aborted) throw new Error('cancelled')
            if (creating) await store(record)
            await ensureBackups(writeSignal)
            if (call.name === 'undo_existing_presentation_batch') {
              // A completed Office write can lose its durable receipt. Include that
              // target in the reverse prefix before switching direction.
              if (record.state === 'applying' && values[record.cursor] === 'after')
                await store({
                  ...record,
                  cursor: record.cursor + 1,
                  state: record.cursor + 1 === record.operations.length ? 'applied' : 'applying',
                })
              const { reviews: _reviews, ...r } = record
              await store({ ...r, state: record.cursor === 0 ? 'undone' : 'undoing' })
            }
            while (record.state === 'applying' || record.state === 'undoing') {
              if (writeSignal?.aborted) throw new Error('cancelled')
              const back = record.state === 'undoing'
              const index = back ? record.cursor - 1 : record.cursor
              const op = record.operations[index]
              const source = back ? op.after : op.before,
                target = back ? op.before : op.after
              const v = await value(op)
              if (!matches(v, target)) {
                if (!matches(v, source)) throw new Error('presentation_existing_batch_conflict')
                if (writeSignal?.aborted) throw new Error('cancelled')
                if (op.kind === 'table_cell')
                  await options.adapter.editPresentationTableCell!(
                    op.hostSlideId,
                    op.shapeId,
                    op.rowIndex,
                    op.columnIndex,
                    target as string,
                    source as string,
                    writeSignal,
                  )
                else if (op.kind === 'text')
                  await options.adapter.editPresentationPageText!(
                    op.hostSlideId,
                    op.shapeId,
                    target as string,
                    source as string,
                    writeSignal,
                  )
                else
                  await options.adapter.editPresentationPageGeometry!(
                    op.hostSlideId,
                    op.shapeId,
                    target as PresentationPageGeometry,
                    source as PresentationPageGeometry,
                    writeSignal,
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
          postWrite: async (): Promise<ProposalPostWriteEvidence> => {
            const terminal = reverse ? 'undone' : 'applied'
            const check = async () => {
              await current()
              saved()
              if (
                record.state !== terminal ||
                !(await classify()).every((v) => v === (reverse ? 'before' : 'after'))
              )
                throw new Error('presentation_existing_batch_conflict')
            }
            await check()
            if (!options.adapter.inspectPresentationPage)
              return { status: 'unavailable', reason: 'capture_unavailable' }
            const pages: Extract<
              ProposalPostWriteEvidence,
              { status: 'captured' }
            >['pages'][number][] = []
            for (const slideId of new Set(record.operations.map((op) => op.hostSlideId))) {
              const shot = await options.adapter.inspectPresentationPage(slideId)
              if (
                shot.slideId !== slideId ||
                shot.shapesTruncated ||
                shot.screenshot.mime !== 'image/png'
              )
                throw new Error('office_read_failed')
              const pngBase64 = validatePowerPointPageScreenshot(shot.screenshot.base64)
              await check()
              pages.push({ slideId, pngBase64, digest: await pngDigest(pngBase64) })
            }
            await check()
            return { status: 'captured', pages }
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
