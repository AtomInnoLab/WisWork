import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import JSZip from 'jszip'
import {
  selectionFingerprint,
  type ProposalPostWriteEvidence,
  type StructuredProposalController,
} from '../../agent/proposal-controller.js'
import {
  validatePowerPointPageScreenshot,
  type PowerPointPageInspection,
} from './browser-powerpoint-adapter.js'
import type { InMemoryVfs } from '../shared/vfs.js'
import type { PresentationBaselineSkill } from './presentation-baseline.js'
import type { PresentationPageReplacementAdapter } from './browser-presentation-page-replacement-adapter.js'
import type { PresentationPageReplacement } from './presentation-page-replacement-record.js'
import {
  presentationPackageDigest,
  MAX_PPTX_PACKAGE_BYTES,
  resolvePowerPointPictureIdentity,
} from './powerpoint-package.js'
import {
  validatePresentationExistingPageChange,
  type PresentationExistingPageChange,
} from './presentation-existing-page.js'
import { readBoundedImage } from '../shared/import-media.js'
import { replacePowerPointPictureMediaPackage } from './presentation-picture-package.js'
import { replacePowerPointTextRangePackage } from './presentation-text-revision-package.js'
import { preparePowerPointCompositePagePackage } from './presentation-composite-revision-package.js'
import { readChartPackageBackup } from './presentation-chart-backup.js'
import {
  validatePresentationExistingChange,
  type PresentationExistingChange,
} from './presentation-existing-change.js'
import {
  validatePresentationExistingBatch,
  type PresentationExistingBatch,
} from './presentation-existing-batch.js'

interface Options {
  baseline: PresentationBaselineSkill
  adapter: PresentationPageReplacementAdapter
  inspectPage(
    slideId: string,
    signal?: AbortSignal,
  ): Promise<Pick<PowerPointPageInspection, 'slideId' | 'shapesTruncated' | 'screenshot'>>
  exportAdapter: {
    exportPresentationPagePackage(
      slideId: string,
      signal?: AbortSignal,
    ): Promise<{ slideId: string; slideIds: string[]; base64: string }>
  }
  vfs: InMemoryVfs
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  proposals: StructuredProposalController
  documentId(): Promise<string>
  readExistingPageChange(id: string): PresentationExistingPageChange | undefined
  readExistingChange?(id: string): PresentationExistingChange | undefined
  readExistingBatch?(id: string): PresentationExistingBatch | undefined
  writeExistingPageChange(
    record: PresentationExistingPageChange,
    expected: PresentationExistingPageChange | undefined,
  ): Promise<void>
  available(): boolean
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const id = (v: unknown): v is string => typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
const host = (v: unknown): v is string =>
  typeof v === 'string' &&
  v.length > 0 &&
  v.length <= 256 &&
  !Array.from(v).some((c) => c.charCodeAt(0) < 32)
const b64 = (bytes: Uint8Array) => btoa(Array.from(bytes, (x) => String.fromCharCode(x)).join(''))
const bytes = (base64: string) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
const sha = async (value: Uint8Array) =>
  Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', Uint8Array.from(value).buffer)),
    (x) => x.toString(16).padStart(2, '0'),
  ).join('')
const output = (value: unknown) => JSON.stringify(value)
const exact = (value: unknown, keys: string[]): value is Record<string, unknown> =>
  !!value &&
  typeof value === 'object' &&
  !Array.isArray(value) &&
  Object.keys(value).length === keys.length &&
  Object.keys(value).every((key) => keys.includes(key))
const CHUNK = 128 * 1024

function projected(r: PresentationExistingPageChange): PresentationPageReplacement {
  return {
    version: 1,
    changeId: r.changeId,
    documentId: r.documentId,
    projectId: 'existing',
    parentRequestId: 'baseline',
    requestId: r.changeId,
    pageId: 'existing',
    backupId: r.backup.backupId,
    parentArtifactDigest: r.baselineDigest,
    backupDigest: r.backup.sha256,
    originalPackageDigest: r.originalPackageDigest,
    replacementPackageDigest: r.replacementPackageDigest,
    sourceSlideId: r.sourceSlideId,
    oldSlideId: r.oldSlideId,
    beforeSlideIds: r.beforeSlideIds,
    state: r.state,
    ...(r.newSlideId ? { newSlideId: r.newSlideId } : {}),
    ...(r.restoredSlideId ? { restoredSlideId: r.restoredSlideId } : {}),
  }
}
async function oneSlide(bytesValue: Uint8Array, signal?: AbortSignal) {
  if (!bytesValue.length || bytesValue.length > MAX_PPTX_PACKAGE_BYTES)
    throw new Error('presentation_page_source_invalid')
  const base64 = b64(bytesValue)
  const digest = await presentationPackageDigest(base64, signal)
  const zip = await JSZip.loadAsync(bytesValue)
  const xml = await zip.file('ppt/presentation.xml')?.async('string')
  const ids = [...(xml ?? '').matchAll(/<p:sldId\b[^>]*\bid="([0-9]+)"[^>]*\/?\s*>/g)]
  const slidePaths = Object.keys(zip.files).filter((path) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(path),
  )
  const relations = await zip.file('ppt/_rels/presentation.xml.rels')?.async('string')
  if (
    ids.length !== 1 ||
    Number(ids[0]![1]) < 256 ||
    Number(ids[0]![1]) > 4294967295 ||
    slidePaths.length !== 1 ||
    !relations?.includes(`Target="slides/${slidePaths[0]!.split('/').at(-1)}"`)
  )
    throw new Error('presentation_page_source_invalid')
  const slideXml = await zip.file(slidePaths[0]!)!.async('string')
  const imageOnly =
    /<(?:[A-Za-z][\w.-]*:)?pic\b/.test(slideXml) &&
    !/<(?:[A-Za-z][\w.-]*:)?(?:sp|graphicFrame|grpSp|cxnSp)\b/.test(slideXml)
  return {
    base64,
    digest,
    sourceSlideId: `${ids[0]![1]}#`,
    sha256: await sha(bytesValue),
    sizeBytes: bytesValue.length,
    imageOnly,
  }
}
const names = [
  'stage',
  'reapply',
  'inspect',
  'reconcile',
  'resume',
  'commit',
  'discard',
  'undo',
  'release',
  'capture',
  'record',
] as const
const tools: AgentToolDef[] = names.map((action) => ({
  name: `${action}_existing_presentation_page_change`,
  description:
    action === 'reapply'
      ? 'Confirm staging the durable source of an undone page as a new independent change; separately inspect and commit.'
      : action === 'stage'
        ? 'Confirm backup and stage a one-slide PPTX after an existing native page; original remains.'
        : action === 'inspect'
          ? 'Inspect current native page order and content against a saved change.'
          : action === 'reconcile'
            ? 'Read-only recovery for a pending page insertion: identify one exact inserted page and journal its host ID without replaying a write.'
            : action === 'resume'
              ? 'Finish journal for a known inserted page without replaying insertion.'
              : action === 'commit'
                ? 'Confirm replacing the original with the verified staged page.'
                : action === 'discard'
                  ? 'Confirm deleting the verified staged page while keeping the original.'
                  : action === 'undo'
                    ? 'Confirm restoring the backed-up original page and removing the replacement.'
                    : action === 'release'
                      ? 'Confirm releasing the PC backup of a discarded or undone page change.'
                      : action === 'capture'
                        ? 'Capture one affected page for visual judgment; does not pass QA.'
                        : 'Persist a visual judgment for one unchanged captured page.',
  inputSchema: {
    type: 'object',
    properties:
      action === 'stage'
        ? {
            baseline_id: { type: 'string' },
            slide_id: { type: 'string' },
            path: { type: 'string' },
            picture_shape_id: { type: 'string' },
            explanation: { type: 'string' },
            restore_source_kind: { type: 'string', enum: ['single', 'batch'] },
            restore_source_change_id: { type: 'string' },
          }
        : {
            change_id: { type: 'string' },
            ...(action === 'capture' || action === 'record'
              ? { slide_id: { type: 'string' } }
              : {}),
            ...(action === 'record'
              ? {
                  screenshot_digest: { type: 'string' },
                  status: { type: 'string', enum: ['pass', 'fail'] },
                  notes: { type: 'string', maxLength: 2000 },
                }
              : {}),
          },
    required:
      action === 'stage'
        ? ['baseline_id', 'slide_id', 'path']
        : action === 'capture'
          ? ['change_id', 'slide_id']
          : action === 'record'
            ? ['change_id', 'slide_id', 'screenshot_digest', 'status', 'notes']
            : ['change_id'],
    additionalProperties: false,
  },
}))
tools.unshift({
  name: 'prepare_existing_presentation_composite_revision',
  description:
    'Prepare one editable, single-page PPTX revision with native text, geometry, and ordinary picture media changes on three distinct shapes. No host write. Stage the returned path once through the confirmed existing-page change flow for one durable backup and undo record.',
  inputSchema: {
    type: 'object',
    properties: {
      baseline_id: { type: 'string' },
      slide_id: { type: 'string' },
      text: {
        type: 'object',
        properties: {
          shape_id: { type: 'string' },
          start: { type: 'integer', minimum: 0 },
          before: { type: 'string' },
          after: { type: 'string' },
        },
        required: ['shape_id', 'start', 'before', 'after'],
        additionalProperties: false,
      },
      geometry: {
        type: 'object',
        properties: {
          shape_id: { type: 'string' },
          before: { type: 'object' },
          after: { type: 'object' },
        },
        required: ['shape_id', 'before', 'after'],
        additionalProperties: false,
      },
      picture: {
        type: 'object',
        properties: { shape_id: { type: 'string' }, path: { type: 'string' } },
        required: ['shape_id', 'path'],
        additionalProperties: false,
      },
    },
    required: ['baseline_id', 'slide_id', 'text', 'geometry', 'picture'],
    additionalProperties: false,
  },
})
tools.unshift({
  name: 'prepare_existing_presentation_text_revision',
  description:
    'Prepare an equal-length native text edit across formatting runs as a one-slide PPTX. This writes only to VFS; stage the returned PPTX with the existing page change flow.',
  inputSchema: {
    type: 'object',
    properties: {
      baseline_id: { type: 'string' },
      slide_id: { type: 'string' },
      shape_id: { type: 'string' },
      start: { type: 'integer', minimum: 0 },
      before: { type: 'string', minLength: 1, maxLength: 128 },
      after: { type: 'string', minLength: 1, maxLength: 128 },
    },
    required: ['baseline_id', 'slide_id', 'shape_id', 'start', 'before', 'after'],
    additionalProperties: false,
  },
})
tools.unshift({
  name: 'prepare_existing_presentation_image_revision',
  description:
    'Prepare a one-slide PPTX with one native picture media replaced. This writes only to VFS; stage the returned PPTX with the existing page change flow.',
  inputSchema: {
    type: 'object',
    properties: {
      baseline_id: { type: 'string' },
      slide_id: { type: 'string' },
      shape_id: { type: 'string' },
      path: { type: 'string' },
    },
    required: ['baseline_id', 'slide_id', 'shape_id', 'path'],
    additionalProperties: false,
  },
})
tools.unshift({
  name: 'prepare_existing_presentation_original_page_restore',
  description:
    'Read and verify one original native page PPTX from an exact single or batch change savepoint on the paired PC. Prepare a VFS source for a separately confirmed existing-page stage and commit; this does not write PowerPoint.',
  inputSchema: {
    type: 'object',
    properties: {
      source_kind: { type: 'string', enum: ['single', 'batch'] },
      change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
      slide_id: { type: 'string', minLength: 1, maxLength: 256 },
    },
    required: ['source_kind', 'change_id', 'slide_id'],
    additionalProperties: false,
  },
})

export function createPresentationExistingPageEditingSkill(
  options: Options,
): AgentSkill & { clear(): void } {
  let epoch = 0
  let reviewCapture:
    | {
        changeId: string
        slideId: string
        record: string
        digest: string
        capturedAt: string
        epoch: number
      }
    | undefined
  return {
    id: 'presentation-existing-page-editing',
    get tools() {
      return options.available()
        ? tools
        : tools.filter((t) =>
            ['inspect_', 'capture_', 'record_'].some((prefix) => t.name.startsWith(prefix)),
          )
    },
    systemPrompt:
      'For an undone page with retained sourceBackup, reapply_existing_presentation_page_change creates a new independent staged change with fresh original and source backups. Inspect and separately confirm commit; reapply never commits automatically and never recreates missing historical backups. For one text, one geometry, and one ordinary embedded picture change on the same existing page, prepare_existing_presentation_composite_revision creates a single native page revision. Stage its returned path with picture_shape_id through the confirmed existing-page change flow so the three objects share one durable backup, commit, and undo record. A length-changing text replacement is supported only within one native text run. ' +
      'Existing page rebuild uses a validated one-slide VFS PPTX. For an ordinary embedded native picture, prepare_existing_presentation_image_revision creates a one-slide revision in VFS; then stage the returned path with picture_shape_id through the existing page change flow so post-write media can be read back. For equal-length text spanning multiple formatting runs, prepare_existing_presentation_text_revision preserves each run and produces a one-slide VFS revision; stage that path through the same confirmed page change flow. To restore a whole page from a completed single or batch existing-edit savepoint, call prepare_existing_presentation_original_page_restore for its exact change and slide; read a fresh page baseline, then stage using the returned path and restore_source_kind/restore_source_change_id, inspect both pages, and separately confirm commit. The edited page is backed up before stage and remains until commit; the restored page receives a new host slide ID. Preparation does not modify PowerPoint. An unverified image-only source is rejected when the original page has native content; keep editable text and complex objects native where possible. Stage retains the original. A pending insertion with unknown host ID must use reconcile_pending_existing_presentation_page_change before any retry; it only accepts an exact page/order/package match and does not replay a write. Inspect and resume recorded interrupted insertions before further action. Commit and undo require separate confirmation. After a confirmed write, capture_existing_presentation_page_change for each affected slide_id, visually inspect the image, then record_existing_presentation_page_change with the same slide_id and screenshot_digest plus pass/fail notes. Inspect compares current screenshots to historical captures per page when possible; a match is not current or whole-deck QA.',
    clear() {
      epoch++
      reviewCapture = undefined
    },
    async executeTool(call, signal) {
      const token = epoch
      const active = () => {
        if (epoch !== token || signal?.aborted) throw new Error('cancelled')
      }
      try {
        const tool = tools.find((t) => t.name === call.name)
        if (
          !tool ||
          call.inputError ||
          call.truncated ||
          !call.input ||
          typeof call.input !== 'object'
        )
          throw new Error('invalid_tool_input')
        const schema = tool.inputSchema as {
          properties: Record<string, unknown>
          required: string[]
        }
        if (
          Object.keys(call.input).some((k) => !(k in schema.properties)) ||
          schema.required.some((k) => !(k in call.input))
        )
          throw new Error('invalid_tool_input')
        const reapplying = call.name === 'reapply_existing_presentation_page_change'
        const action = reapplying ? 'stage' : call.name.split('_')[0]
        if (!['inspect', 'capture', 'record'].includes(action) && !options.available())
          throw new Error('presentation_page_backup_unavailable')
        const documentId = await options.documentId()
        const current = async () => {
          active()
          if ((await options.documentId()) !== documentId)
            throw new Error('presentation_document_changed')
          active()
        }
        const request = async (operation: string, data: Record<string, unknown>) => {
          await current()
          const response = await options.request({ operation, documentId, ...data }, signal)
          await current()
          const value = (await response.json()) as Record<string, unknown>
          if (
            operation === 'existing_page_backup_begin' &&
            value &&
            typeof value === 'object' &&
            !Array.isArray(value) &&
            value.error === 'quota_exceeded'
          )
            throw new Error('presentation_existing_backup_capacity')
          if (!response.ok || !value || typeof value !== 'object' || 'error' in value)
            throw new Error('presentation_page_backup_failed')
          return value
        }
        const restoreSource = async (kind: unknown, changeId: unknown, slideId: unknown) => {
          if (!['single', 'batch'].includes(kind as string) || !id(changeId) || !host(slideId))
            throw new Error('invalid_tool_input')
          const sourceRecord =
            kind === 'single'
              ? options.readExistingChange?.(changeId)
              : options.readExistingBatch?.(changeId)
          const savedRecord = sourceRecord && structuredClone(sourceRecord)
          if (
            !savedRecord ||
            savedRecord.documentId !== documentId ||
            savedRecord.changeId !== changeId ||
            (kind === 'batch' && (savedRecord as PresentationExistingBatch).version === 4) ||
            !(kind === 'single'
              ? validatePresentationExistingChange(savedRecord)
              : validatePresentationExistingBatch(savedRecord)) ||
            !(
              kind === 'batch' && savedRecord.version !== 1
                ? ['applying', 'applied', 'undoing', 'undone']
                : ['applied', 'undone']
            ).includes(savedRecord.state) ||
            savedRecord.backupReleasedAt
          )
            throw new Error('presentation_original_restore_source_invalid')
          const backup =
            kind === 'single'
              ? (savedRecord as PresentationExistingChange).backup
              : (savedRecord as PresentationExistingBatch).backups?.find(
                  (item) => item.hostSlideId === slideId,
                )
          const slideIds = savedRecord.beforeSlideIds
          if (
            !backup ||
            !slideIds ||
            backup.hostSlideId !== slideId ||
            (kind === 'single' &&
              (savedRecord as PresentationExistingChange).hostSlideId !== slideId)
          )
            throw new Error('presentation_original_restore_source_invalid')
          const base64 = await readChartPackageBackup(
            {
              request: options.request,
              documentId,
              hostSlideId: slideId,
              slideIds,
              backup,
              expectedPackageDigest: backup.packageDigest,
            },
            signal,
          )
          await current()
          if (
            !same(
              savedRecord,
              kind === 'single'
                ? options.readExistingChange?.(changeId)
                : options.readExistingBatch?.(changeId),
            )
          )
            throw new Error('presentation_original_restore_source_stale')
          const packageSource = await oneSlide(bytes(base64), signal)
          if (
            packageSource.digest !== backup.packageDigest ||
            packageSource.sha256 !== backup.sha256
          )
            throw new Error('presentation_original_restore_source_invalid')
          return {
            base64,
            packageSource,
            backup,
            sourceKind: kind as 'single' | 'batch',
            sourceChangeId: changeId,
            slideId,
          }
        }
        if (call.name === 'prepare_existing_presentation_original_page_restore') {
          const input = call.input
          const source = await restoreSource(input.source_kind, input.change_id, input.slide_id)
          const currentPage = await options.exportAdapter.exportPresentationPagePackage(
            source.slideId,
            signal,
          )
          await current()
          if (
            currentPage.slideId !== source.slideId ||
            !currentPage.slideIds.includes(source.slideId)
          )
            throw new Error('presentation_original_restore_target_changed')
          const path = `/home/user/presentation-original-restore-${crypto.randomUUID()}.pptx`
          options.vfs.writeFile(path, bytes(source.base64))
          return {
            output: output({
              path,
              slideId: source.slideId,
              sourceKind: source.sourceKind,
              sourceChangeId: source.sourceChangeId,
              packageDigest: source.backup.packageDigest,
              nextTool: 'stage_existing_presentation_page_change',
              nextInput: {
                path,
                slide_id: source.slideId,
                restore_source_kind: source.sourceKind,
                restore_source_change_id: source.sourceChangeId,
              },
            }),
            mutated: false,
            summary: '已准备原页包恢复来源，尚未修改演示文稿',
          }
        }
        if (call.name === 'prepare_existing_presentation_text_revision') {
          const input = call.input
          if (
            !id(input.baseline_id) ||
            !host(input.slide_id) ||
            typeof input.shape_id !== 'string' ||
            !/^[1-9]\d{0,9}$/.test(input.shape_id) ||
            !Number.isSafeInteger(input.start) ||
            (input.start as number) < 0 ||
            typeof input.before !== 'string' ||
            typeof input.after !== 'string'
          )
            throw new Error('invalid_tool_input')
          const baseline = options.baseline.snapshot(input.baseline_id)
          if (
            !baseline ||
            baseline.documentId !== documentId ||
            !baseline.scope.slideIds.includes(input.slide_id) ||
            baseline.scope.shapeIds?.length
          )
            throw new Error('presentation_existing_scope_mismatch')
          const original = await options.exportAdapter.exportPresentationPagePackage(
            input.slide_id,
            signal,
          )
          await current()
          if (
            original.slideId !== input.slide_id ||
            !same(original.slideIds, baseline.context.slideIds)
          )
            throw new Error('presentation_baseline_changed')
          const check = await options.baseline.executeTool(
            {
              id: 'page-text-check',
              name: 'check_presentation_baseline',
              input: { baseline_id: baseline.baselineId },
            },
            signal,
          )
          await current()
          if (check.isError || JSON.parse(check.output).unchanged !== true)
            throw new Error('presentation_baseline_changed')
          const revision = await replacePowerPointTextRangePackage(
            original.base64,
            input.shape_id,
            input.start as number,
            input.before,
            input.after,
            signal,
          )
          await current()
          const path = `/home/user/presentation-text-revision-${crypto.randomUUID()}.pptx`
          options.vfs.writeFile(path, bytes(revision.base64))
          return {
            output: output({
              path,
              baselineId: baseline.baselineId,
              slideId: input.slide_id,
              shapeId: input.shape_id,
              beforeDigest: revision.beforeDigest,
              afterDigest: revision.afterDigest,
              changedRuns: revision.changedRuns,
              nextTool: 'stage_existing_presentation_page_change',
            }),
            mutated: false,
            summary: '已准备保留格式的原生文字修订稿，尚未修改演示文稿',
          }
        }
        if (call.name === 'prepare_existing_presentation_composite_revision') {
          const input = call.input
          if (
            !id(input.baseline_id) ||
            !host(input.slide_id) ||
            !exact(input.text, ['shape_id', 'start', 'before', 'after']) ||
            !exact(input.geometry, ['shape_id', 'before', 'after']) ||
            !exact(input.picture, ['shape_id', 'path'])
          )
            throw new Error('invalid_tool_input')
          const textEdit = input.text
          const geometryEdit = input.geometry
          const pictureEdit = input.picture
          const shapeId = (value: unknown): value is string =>
            typeof value === 'string' && /^[1-9]\d{0,9}$/.test(value)
          const coordinates = (
            value: unknown,
          ): value is { left: number; top: number; width: number; height: number } =>
            exact(value, ['left', 'top', 'width', 'height']) &&
            Object.values(value).every((n) => typeof n === 'number' && Number.isFinite(n))
          if (
            !shapeId(textEdit.shape_id) ||
            !Number.isSafeInteger(textEdit.start) ||
            (textEdit.start as number) < 0 ||
            typeof textEdit.before !== 'string' ||
            typeof textEdit.after !== 'string' ||
            !shapeId(geometryEdit.shape_id) ||
            !coordinates(geometryEdit.before) ||
            !coordinates(geometryEdit.after) ||
            !shapeId(pictureEdit.shape_id) ||
            typeof pictureEdit.path !== 'string' ||
            new Set([textEdit.shape_id, geometryEdit.shape_id, pictureEdit.shape_id]).size !== 3
          )
            throw new Error('invalid_tool_input')
          const baseline = options.baseline.snapshot(input.baseline_id)
          if (
            !baseline ||
            baseline.documentId !== documentId ||
            !baseline.scope.slideIds.includes(input.slide_id) ||
            baseline.scope.shapeIds?.length
          )
            throw new Error('presentation_existing_scope_mismatch')
          const image = await readBoundedImage(options.vfs, pictureEdit.path)
          const original = await options.exportAdapter.exportPresentationPagePackage(
            input.slide_id,
            signal,
          )
          await current()
          if (
            original.slideId !== input.slide_id ||
            !same(original.slideIds, baseline.context.slideIds)
          )
            throw new Error('presentation_baseline_changed')
          const check = await options.baseline.executeTool(
            {
              id: 'page-composite-check',
              name: 'check_presentation_baseline',
              input: { baseline_id: baseline.baselineId },
            },
            signal,
          )
          await current()
          if (check.isError || JSON.parse(check.output).unchanged !== true)
            throw new Error('presentation_baseline_changed')
          const revision = await preparePowerPointCompositePagePackage(
            original.base64,
            {
              text: {
                shapeId: textEdit.shape_id,
                start: textEdit.start as number,
                before: textEdit.before,
                after: textEdit.after,
              },
              geometry: {
                shapeId: geometryEdit.shape_id,
                before: geometryEdit.before,
                after: geometryEdit.after,
              },
              picture: { shapeId: pictureEdit.shape_id, image },
            },
            signal,
          )
          await current()
          const path = `/home/user/presentation-composite-revision-${crypto.randomUUID()}.pptx`
          options.vfs.writeFile(path, bytes(revision.base64))
          return {
            output: output({
              path,
              baselineId: baseline.baselineId,
              slideId: input.slide_id,
              beforeDigest: revision.beforeDigest,
              afterDigest: revision.afterDigest,
              changedRuns: revision.changedRuns,
              mediaDigest: revision.mediaDigest,
              nextTool: 'stage_existing_presentation_page_change',
              nextInput: {
                baseline_id: baseline.baselineId,
                slide_id: input.slide_id,
                path,
                picture_shape_id: pictureEdit.shape_id,
              },
            }),
            mutated: false,
            summary: '已准备文字、几何和图片的单页组合修订稿，尚未修改演示文稿',
          }
        }
        if (action === 'prepare') {
          const input = call.input
          if (
            !id(input.baseline_id) ||
            !host(input.slide_id) ||
            typeof input.shape_id !== 'string' ||
            !/^[1-9]\d{0,9}$/.test(input.shape_id) ||
            typeof input.path !== 'string'
          )
            throw new Error('invalid_tool_input')
          const baseline = options.baseline.snapshot(input.baseline_id)
          if (
            !baseline ||
            baseline.documentId !== documentId ||
            !baseline.scope.slideIds.includes(input.slide_id) ||
            baseline.scope.shapeIds?.length
          )
            throw new Error('presentation_existing_scope_mismatch')
          const image = await readBoundedImage(options.vfs, input.path)
          const original = await options.exportAdapter.exportPresentationPagePackage(
            input.slide_id,
            signal,
          )
          await current()
          if (
            original.slideId !== input.slide_id ||
            !same(original.slideIds, baseline.context.slideIds)
          )
            throw new Error('presentation_baseline_changed')
          const check = await options.baseline.executeTool(
            {
              id: 'page-check',
              name: 'check_presentation_baseline',
              input: { baseline_id: baseline.baselineId },
            },
            signal,
          )
          await current()
          if (check.isError || JSON.parse(check.output).unchanged !== true)
            throw new Error('presentation_baseline_changed')
          const revision = await replacePowerPointPictureMediaPackage(
            original.base64,
            input.shape_id,
            image,
            signal,
          )
          await current()
          const path = `/home/user/presentation-image-revision-${crypto.randomUUID()}.pptx`
          options.vfs.writeFile(path, bytes(revision.base64))
          return {
            output: output({
              path,
              baselineId: baseline.baselineId,
              slideId: input.slide_id,
              shapeId: input.shape_id,
              beforeDigest: revision.beforeDigest,
              afterDigest: revision.afterDigest,
              mediaDigest: revision.mediaDigest,
              nextTool: 'stage_existing_presentation_page_change',
              nextPictureShapeId: input.shape_id,
            }),
            mutated: false,
            summary: '已准备原生图片修订稿，尚未修改演示文稿',
          }
        }
        const loadRetainedPackage = async (
          owner: PresentationExistingPageChange,
          metadata: PresentationExistingPageChange['backup'],
          expectedDigest: string,
        ) => {
          const status = await request('existing_page_backup_status', {
            backupId: metadata.backupId,
          })
          if (
            status.status !== 'ready' ||
            status.backupId !== metadata.backupId ||
            status.documentId !== documentId ||
            status.hostSlideId !== owner.oldSlideId ||
            !same(status.slideIds, owner.beforeSlideIds) ||
            status.sha256 !== metadata.sha256 ||
            status.sizeBytes !== metadata.sizeBytes ||
            status.receivedBytes !== status.sizeBytes
          )
            throw new Error('presentation_page_backup_invalid')
          const content = new Uint8Array(metadata.sizeBytes)
          for (let offset = 0; offset < content.length; offset += CHUNK) {
            const length = Math.min(CHUNK, content.length - offset)
            const part = await request('existing_page_backup_read', {
              backupId: metadata.backupId,
              offset,
              length,
            })
            if (
              part.backupId !== metadata.backupId ||
              part.offset !== offset ||
              part.sizeBytes !== content.length ||
              part.sha256 !== metadata.sha256 ||
              typeof part.base64 !== 'string'
            )
              throw new Error('presentation_page_backup_invalid')
            const chunk = bytes(part.base64)
            if (chunk.length !== length || b64(chunk) !== part.base64)
              throw new Error('presentation_page_backup_invalid')
            content.set(chunk, offset)
          }
          if (
            (await sha(content)) !== metadata.sha256 ||
            (await presentationPackageDigest(b64(content), signal)) !== expectedDigest
          )
            throw new Error('presentation_page_backup_invalid')
          await current()
          return b64(content)
        }
        const savePackage = async (
          owner: PresentationExistingPageChange,
          metadata: PresentationExistingPageChange['backup'],
          content: Uint8Array,
          packageDigest: string,
        ) => {
          const match = (value: Record<string, unknown>) =>
            value.backupId === metadata.backupId &&
            value.documentId === documentId &&
            value.hostSlideId === owner.oldSlideId &&
            same(value.slideIds, owner.beforeSlideIds) &&
            value.sha256 === metadata.sha256 &&
            value.sizeBytes === content.length
          let meta = await request('existing_page_backup_begin', {
            backupId: metadata.backupId,
            hostSlideId: owner.oldSlideId,
            slideIds: owner.beforeSlideIds,
            sha256: metadata.sha256,
            sizeBytes: content.length,
          })
          if (!match(meta)) throw new Error('presentation_page_backup_invalid')
          while (meta.status !== 'ready' && Number(meta.receivedBytes) < content.length) {
            const offset = Number(meta.receivedBytes)
            if (!Number.isSafeInteger(offset) || offset < 0)
              throw new Error('presentation_page_backup_invalid')
            const chunk = content.subarray(offset, Math.min(offset + CHUNK, content.length))
            meta = await request('existing_page_backup_chunk', {
              backupId: metadata.backupId,
              offset,
              base64: b64(chunk),
            })
            if (!match(meta) || meta.receivedBytes !== offset + chunk.length)
              throw new Error('presentation_page_backup_invalid')
          }
          if (meta.status !== 'ready')
            meta = await request('existing_page_backup_finish', { backupId: metadata.backupId })
          if (!match(meta) || meta.status !== 'ready' || meta.receivedBytes !== content.length)
            throw new Error('presentation_page_backup_invalid')
          if ((await loadRetainedPackage(owner, metadata, packageDigest)) !== b64(content))
            throw new Error('presentation_page_backup_invalid')
        }
        const releasePackage = async (
          owner: PresentationExistingPageChange,
          metadata: PresentationExistingPageChange['backup'],
        ) => {
          const receipt = await request('existing_page_backup_release', {
            backupId: metadata.backupId,
            hostSlideId: owner.oldSlideId,
            slideIds: owner.beforeSlideIds,
            sha256: metadata.sha256,
            sizeBytes: metadata.sizeBytes,
          })
          if (
            receipt.status !== 'released' ||
            receipt.backupId !== metadata.backupId ||
            receipt.documentId !== documentId ||
            receipt.hostSlideId !== owner.oldSlideId ||
            !same(receipt.slideIds, owner.beforeSlideIds) ||
            receipt.sha256 !== metadata.sha256 ||
            receipt.sizeBytes !== metadata.sizeBytes
          )
            throw new Error('presentation_page_backup_invalid')
        }
        let reappliedRecord: PresentationExistingPageChange | undefined
        let restoredObservation: { slideId: string; slideIds: string[]; base64: string } | undefined
        let record: PresentationExistingPageChange
        let source: Awaited<ReturnType<typeof oneSlide>> | undefined
        let sourcePath: string | undefined
        let restoreProof: Awaited<ReturnType<typeof restoreSource>> | undefined
        let baseline = undefined as ReturnType<PresentationBaselineSkill['snapshot']>
        if (action === 'stage' && !reapplying) {
          const input = call.input
          if (
            !id(input.baseline_id) ||
            !host(input.slide_id) ||
            typeof input.path !== 'string' ||
            !input.path.endsWith('.pptx') ||
            input.path.length > 1024 ||
            (input.picture_shape_id !== undefined &&
              (typeof input.picture_shape_id !== 'string' ||
                !/^[1-9]\d{0,9}$/.test(input.picture_shape_id))) ||
            (input.explanation !== undefined &&
              (typeof input.explanation !== 'string' || input.explanation.length > 300)) ||
            (input.restore_source_kind === undefined) !==
              (input.restore_source_change_id === undefined) ||
            (input.restore_source_kind !== undefined &&
              (input.picture_shape_id !== undefined ||
                !['single', 'batch'].includes(input.restore_source_kind as string) ||
                !id(input.restore_source_change_id)))
          )
            throw new Error('invalid_tool_input')
          baseline = options.baseline.snapshot(input.baseline_id)
          if (
            !baseline ||
            baseline.documentId !== documentId ||
            !baseline.scope.slideIds.includes(input.slide_id) ||
            baseline.scope.shapeIds?.length
          )
            throw new Error('presentation_existing_scope_mismatch')
          sourcePath = input.path
          source = await oneSlide(
            options.vfs.readBytes(sourcePath, { maxBytes: MAX_PPTX_PACKAGE_BYTES + 1 }),
            signal,
          )
          if (input.restore_source_kind !== undefined) {
            restoreProof = await restoreSource(
              input.restore_source_kind,
              input.restore_source_change_id,
              input.slide_id,
            )
            if (
              restoreProof.packageSource.sha256 !== source.sha256 ||
              restoreProof.packageSource.digest !== source.digest
            )
              throw new Error('presentation_original_restore_source_invalid')
          } else if (sourcePath.startsWith('/home/user/presentation-original-restore-'))
            throw new Error('presentation_original_restore_source_invalid')
          if (
            !restoreProof &&
            source.imageOnly &&
            baseline.pages
              ?.find((page) => page.slideId === input.slide_id)
              ?.shapes.some((shape) => shape.type !== 'Image')
          )
            throw new Error('presentation_page_source_rasterized')
          const original = await options.exportAdapter.exportPresentationPagePackage(
            input.slide_id,
            signal,
          )
          await current()
          if (
            original.slideId !== input.slide_id ||
            !same(original.slideIds, baseline.context.slideIds)
          )
            throw new Error('presentation_baseline_changed')
          const originalPackageDigest = await presentationPackageDigest(original.base64, signal)
          let pictureTarget: PresentationExistingPageChange['pictureTarget']
          if (input.picture_shape_id) {
            const [before, after] = await Promise.all([
              resolvePowerPointPictureIdentity(
                original.base64,
                { shapeId: input.picture_shape_id },
                signal,
              ),
              resolvePowerPointPictureIdentity(
                source.base64,
                { shapeId: input.picture_shape_id },
                signal,
              ),
            ])
            if (before.name !== after.name || before.mediaDigest === after.mediaDigest)
              throw new Error('presentation_page_source_invalid')
            pictureTarget = {
              shapeId: input.picture_shape_id,
              name: before.name,
              beforeDigest: before.mediaDigest,
              afterDigest: after.mediaDigest,
            }
          }
          record = {
            version: 1,
            changeId: crypto.randomUUID(),
            documentId,
            baselineId: baseline.baselineId,
            baselineDigest: baseline.contentDigest,
            scope: { slideIds: [...baseline.scope.slideIds] },
            oldSlideId: input.slide_id,
            beforeSlideIds: [...baseline.context.slideIds],
            originalPackageDigest,
            replacementPackageDigest: source.digest,
            sourceSlideId: source.sourceSlideId,
            ...(pictureTarget ? { pictureTarget } : {}),
            ...(restoreProof
              ? {
                  restores: {
                    sourceKind: restoreProof.sourceKind,
                    sourceChangeId: restoreProof.sourceChangeId,
                    sourceHostSlideId: restoreProof.slideId,
                    originalBackupId: restoreProof.backup.backupId,
                    originalPackageDigest: restoreProof.backup.packageDigest,
                  },
                }
              : {}),
            backup: { backupId: crypto.randomUUID(), sha256: '0'.repeat(64), sizeBytes: 1 },
            state: 'pending',
          }
        } else {
          if (!id(call.input.change_id)) throw new Error('invalid_tool_input')
          const saved = options.readExistingPageChange(call.input.change_id)
          if (
            !saved ||
            !validatePresentationExistingPageChange(saved) ||
            saved.documentId !== documentId
          )
            throw new Error('presentation_existing_page_missing')
          record = structuredClone(saved)
          if (reapplying) {
            if (record.state !== 'undone' || record.backupReleasedAt || !record.sourceBackup)
              throw new Error('presentation_existing_page_state_invalid')
            reappliedRecord = structuredClone(record)
            const original = await loadRetainedPackage(
              record,
              record.backup,
              record.originalPackageDigest,
            )
            source = await oneSlide(
              bytes(
                await loadRetainedPackage(
                  record,
                  record.sourceBackup,
                  record.replacementPackageDigest,
                ),
              ),
              signal,
            )
            if (source.sourceSlideId !== record.sourceSlideId)
              throw new Error('presentation_page_backup_invalid')
            restoredObservation = await options.exportAdapter.exportPresentationPagePackage(
              record.restoredSlideId!,
              signal,
            )
            await current()
            if (
              restoredObservation.slideId !== record.restoredSlideId ||
              !same(
                restoredObservation.slideIds,
                record.beforeSlideIds.map((slideId) =>
                  slideId === record.oldSlideId ? record.restoredSlideId! : slideId,
                ),
              ) ||
              (await presentationPackageDigest(restoredObservation.base64, signal)) !==
                record.originalPackageDigest
            )
              throw new Error('presentation_existing_page_conflict')
            if (record.pictureTarget) {
              const target = record.pictureTarget
              const before = await resolvePowerPointPictureIdentity(
                original,
                { shapeId: target.shapeId, name: target.name },
                signal,
              )
              const after = await resolvePowerPointPictureIdentity(
                source.base64,
                { shapeId: target.shapeId, name: target.name },
                signal,
              )
              if (
                before.mediaDigest !== target.beforeDigest ||
                after.mediaDigest !== target.afterDigest ||
                before.name !== target.name ||
                after.name !== target.name
              )
                throw new Error('presentation_page_backup_invalid')
            }
            record = {
              version: 1,
              changeId: crypto.randomUUID(),
              documentId,
              baselineId: crypto.randomUUID(),
              baselineDigest: await sha(
                new TextEncoder().encode(JSON.stringify(restoredObservation)),
              ),
              scope: { slideIds: [record.restoredSlideId!] },
              oldSlideId: record.restoredSlideId!,
              beforeSlideIds: [...restoredObservation.slideIds],
              originalPackageDigest: record.originalPackageDigest,
              replacementPackageDigest: record.replacementPackageDigest,
              sourceSlideId: record.sourceSlideId,
              ...(record.pictureTarget ? { pictureTarget: record.pictureTarget } : {}),
              ...(record.restores ? { restores: { ...record.restores } } : {}),
              backup: { backupId: crypto.randomUUID(), sha256: '0'.repeat(64), sizeBytes: 1 },
              sourceBackup: {
                backupId: crypto.randomUUID(),
                sha256: source.sha256,
                sizeBytes: source.sizeBytes,
              },
              reapplies: record.changeId,
              state: 'pending',
            }
          }
        }
        let expected = action === 'stage' ? undefined : structuredClone(record)
        const saved = () => {
          if (
            reappliedRecord &&
            !same(options.readExistingPageChange(reappliedRecord.changeId), reappliedRecord)
          )
            throw new Error('presentation_existing_page_stale')
          if (!same(options.readExistingPageChange(record.changeId), expected))
            throw new Error('presentation_existing_page_stale')
        }
        const assertCurrent = async () => {
          await current()
          saved()
        }
        const store = async (next: PresentationExistingPageChange) => {
          await current()
          saved()
          await options.writeExistingPageChange(next, expected)
          await current()
          if (!same(options.readExistingPageChange(next.changeId), next))
            throw new Error('office_state_uncertain')
          record = structuredClone(next)
          expected = structuredClone(next)
          reviewCapture = undefined
        }
        if (action === 'reconcile') {
          if (record.state !== 'pending' || record.newSlideId)
            throw new Error('presentation_existing_page_state_invalid')
          const found = await options.adapter.reconcilePending(projected(record), signal)
          await current()
          if (found.status === 'conflict')
            throw new Error('presentation_existing_page_manual_review')
          if (found.status === 'baseline')
            return {
              output: output({
                changeId: record.changeId,
                status: 'pending_no_insert_observed',
                hostWrite: false,
              }),
              mutated: false,
              summary: '未观察到暂存页；保留待核对记录，不重放写入',
            }
          if (!host(found.newSlideId)) throw new Error('office_state_uncertain')
          await store({ ...record, state: 'inserted', newSlideId: found.newSlideId })
          if ((await options.adapter.inspect(projected(record), signal)).status !== 'staged')
            throw new Error('presentation_existing_page_manual_review')
          await store({ ...record, state: 'staged' })
          return {
            output: output({
              changeId: record.changeId,
              status: 'staged',
              newSlideId: found.newSlideId,
              hostWrite: false,
            }),
            mutated: false,
            summary: '已核对并认领暂存页，可继续检查与确认提交',
          }
        }
        if (action === 'release') {
          if (!['discarded', 'undone'].includes(record.state) || record.backupReleasedAt)
            throw new Error('presentation_existing_page_state_invalid')
          const proposal = options.proposals.propose({
            operation: call.name,
            toolName: call.name,
            title: record.sourceBackup
              ? '释放已结束整页变更的原页与替换源备份'
              : '释放已结束整页变更的原页备份',
            preview: {
              changeId: record.changeId,
              state: record.state,
              backupId: record.backup.backupId,
              sizeBytes: record.backup.sizeBytes + (record.sourceBackup?.sizeBytes ?? 0),
              sourceBackupId: record.sourceBackup?.backupId,
              backupCount: record.sourceBackup ? 2 : 1,
            },
            impact: {
              host: 'powerpoint',
              targets: [record.backup, ...(record.sourceBackup ? [record.sourceBackup] : [])].map(
                (backup) => `backup:${backup.backupId}`,
              ),
              count: record.sourceBackup ? 2 : 1,
            },
            fingerprint: selectionFingerprint(output(record)),
            validate: async () => {
              try {
                await assertCurrent()
                return true
              } catch {
                return false
              }
            },
            execute: async () => {
              await assertCurrent()
              await releasePackage(record, record.backup)
              if (record.sourceBackup) await releasePackage(record, record.sourceBackup)
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
            summary: '已准备释放整页备份，等待确认',
          }
        }
        const baselineFresh = async () => {
          if (reapplying) {
            await assertCurrent()
            const observation = await options.exportAdapter.exportPresentationPagePackage(
              record.oldSlideId,
              signal,
            )
            await assertCurrent()
            return (
              !!restoredObservation &&
              observation.slideId === restoredObservation.slideId &&
              same(observation.slideIds, restoredObservation.slideIds) &&
              (await presentationPackageDigest(observation.base64, signal)) ===
                record.originalPackageDigest
            )
          }
          if (!baseline || !same(options.baseline.snapshot(baseline.baselineId), baseline))
            return false
          const result = await options.baseline.executeTool(
            {
              id: 'page-check',
              name: 'check_presentation_baseline',
              input: { baseline_id: baseline.baselineId },
            },
            signal,
          )
          await current()
          return (
            !result.isError &&
            JSON.parse(result.output).unchanged === true &&
            same(options.baseline.snapshot(baseline.baselineId), baseline)
          )
        }
        const sourceFresh = async () => {
          if (reappliedRecord) {
            await assertCurrent()
            await loadRetainedPackage(
              reappliedRecord,
              reappliedRecord.backup,
              reappliedRecord.originalPackageDigest,
            )
            const retained = await oneSlide(
              bytes(
                await loadRetainedPackage(
                  reappliedRecord,
                  reappliedRecord.sourceBackup!,
                  reappliedRecord.replacementPackageDigest,
                ),
              ),
              signal,
            )
            if (
              retained.sha256 !== source!.sha256 ||
              retained.sourceSlideId !== record.sourceSlideId
            )
              throw new Error('presentation_page_backup_invalid')
            await assertCurrent()
            return
          }
          if (!source || !sourcePath) return
          const value = options.vfs.readBytes(sourcePath, { maxBytes: MAX_PPTX_PACKAGE_BYTES + 1 })
          if ((await sha(value)) !== source.sha256 || value.length !== source.sizeBytes)
            throw new Error('proposal_stale')
          if (restoreProof) {
            const latest = await restoreSource(
              restoreProof.sourceKind,
              restoreProof.sourceChangeId,
              restoreProof.slideId,
            )
            if (
              latest.packageSource.sha256 !== source.sha256 ||
              latest.backup.backupId !== restoreProof.backup.backupId
            )
              throw new Error('proposal_stale')
          }
          await current()
        }
        const inspect = async () => {
          await current()
          saved()
          const state = await options.adapter.inspect(projected(record), signal)
          await current()
          saved()
          return state
        }
        if (action === 'inspect') {
          const observed = await inspect()
          const expectedStatuses: Record<PresentationExistingPageChange['state'], string[]> = {
            pending: [],
            inserted: ['staged'],
            staged: ['staged'],
            commit_pending: ['staged', 'applied'],
            applied: ['applied'],
            discard_pending: ['staged', 'baseline'],
            discarded: ['baseline'],
            undo_pending: ['applied'],
            restore_inserted: ['restore_staged', 'undone'],
            undone: ['undone'],
          }
          const verified = expectedStatuses[record.state].includes(observed.status)
          const visualReceipts = [] as {
            hostSlideId: string
            status: 'not_captured' | 'matched' | 'different' | 'unavailable'
          }[]
          if (verified && record.captures?.length) {
            const before = observed.slideIds
            for (const capture of record.captures) {
              let status: 'matched' | 'different' | 'unavailable' = 'unavailable'
              try {
                const shot = await options.inspectPage(capture.hostSlideId, signal)
                await current()
                saved()
                if (
                  shot.slideId !== capture.hostSlideId ||
                  shot.shapesTruncated ||
                  shot.screenshot.mime !== 'image/png'
                )
                  throw new Error('office_read_failed')
                const png = validatePowerPointPageScreenshot(shot.screenshot.base64)
                const currentDigest = await sha(bytes(png))
                const after = await inspect()
                if (after.status !== observed.status || !same(after.slideIds, before))
                  throw new Error('office_state_uncertain')
                status = currentDigest === capture.screenshotDigest ? 'matched' : 'different'
              } catch {
                await current()
                saved()
              }
              visualReceipts.push({ hostSlideId: capture.hostSlideId, status })
            }
          }
          return {
            output: output({
              changeId: record.changeId,
              state: record.state,
              inspection: observed,
              currentHostVerified: verified,
              manualReview: !verified,
              ...(record.state === 'pending'
                ? { nextTool: 'reconcile_pending_existing_presentation_page_change' }
                : {}),
              visualReceipts,
              qaPassed: false,
            }),
            mutated: false,
            summary: '已检查现稿单页变更',
          }
        }
        if (action === 'capture' || action === 'record') {
          const targets =
            record.state === 'staged'
              ? [record.oldSlideId, record.newSlideId!]
              : record.state === 'applied'
                ? [record.newSlideId!]
                : record.state === 'discarded'
                  ? [record.oldSlideId]
                  : record.state === 'undone'
                    ? [record.restoredSlideId!]
                    : []
          const slideId = call.input.slide_id
          if (
            !host(slideId) ||
            !targets.includes(slideId) ||
            record.reviews?.some((review) => review.hostSlideId === slideId)
          )
            throw new Error('presentation_existing_page_state_invalid')
          const status =
            record.state === 'staged'
              ? 'staged'
              : record.state === 'applied'
                ? 'applied'
                : record.state === 'discarded'
                  ? 'baseline'
                  : 'undone'
          const before = await inspect()
          if (before.status !== status || !targets.every((id) => before.slideIds.includes(id)))
            throw new Error('presentation_existing_page_conflict')
          const shot = await options.inspectPage(slideId, signal)
          await current()
          saved()
          if (
            shot.slideId !== slideId ||
            shot.shapesTruncated ||
            shot.screenshot.mime !== 'image/png'
          )
            throw new Error('office_read_failed')
          const png = validatePowerPointPageScreenshot(shot.screenshot.base64)
          const screenshotDigest = await sha(bytes(png))
          const after = await inspect()
          if (after.status !== status || !same(after.slideIds, before.slideIds))
            throw new Error('presentation_existing_page_conflict')
          if (action === 'capture') {
            const capture = {
              hostSlideId: slideId,
              screenshotDigest,
              capturedAt: new Date().toISOString(),
            }
            await store({
              ...record,
              captures: [
                ...(record.captures ?? []).filter((item) => item.hostSlideId !== slideId),
                capture,
              ],
            })
            reviewCapture = {
              changeId: record.changeId,
              slideId,
              record: JSON.stringify(record),
              digest: screenshotDigest,
              capturedAt: capture.capturedAt,
              epoch,
            }
            return {
              output: output({
                changeId: record.changeId,
                hostSlideId: slideId,
                screenshotDigest,
                qaPassed: false,
              }),
              display: { kind: 'images', items: [{ url: `data:image/png;base64,${png}` }] },
              mutated: false,
              summary: '已采集整页变更截图，等待视觉判断',
            }
          }
          const receipt = record.captures?.find((capture) => capture.hostSlideId === slideId)
          if (
            !reviewCapture ||
            reviewCapture.epoch !== epoch ||
            reviewCapture.changeId !== record.changeId ||
            reviewCapture.slideId !== slideId ||
            reviewCapture.record !== JSON.stringify(record) ||
            reviewCapture.digest !== screenshotDigest ||
            receipt?.screenshotDigest !== screenshotDigest ||
            receipt.capturedAt !== reviewCapture.capturedAt ||
            call.input.screenshot_digest !== screenshotDigest ||
            !['pass', 'fail'].includes(call.input.status as string) ||
            typeof call.input.notes !== 'string' ||
            call.input.notes.length > 2000
          )
            throw new Error('presentation_existing_page_review_stale')
          const review = {
            hostSlideId: slideId,
            screenshotDigest,
            capturedAt: reviewCapture.capturedAt,
            reviewedAt: new Date().toISOString(),
            status: call.input.status as 'pass' | 'fail',
            notes: call.input.notes,
          }
          await store({ ...record, reviews: [...(record.reviews ?? []), review] })
          return {
            output: output({
              changeId: record.changeId,
              historicalReview: review,
              qaPassed: false,
            }),
            mutated: false,
            summary: '已保存整页变更历史视觉判断',
          }
        }
        const observed = action === 'stage' ? undefined : await inspect()
        if (action === 'stage') {
          if (!(await baselineFresh())) throw new Error('presentation_baseline_changed')
        } else if (action === 'resume') {
          if (
            !['inserted', 'commit_pending', 'discard_pending', 'restore_inserted'].includes(
              record.state,
            )
          )
            throw new Error('presentation_existing_page_manual_review')
          if (!observed || observed.status === 'conflict')
            throw new Error('presentation_existing_page_manual_review')
        } else if (
          action === 'commit' &&
          ((record.state !== 'staged' && record.state !== 'commit_pending') ||
            !observed ||
            !['staged', 'applied'].includes(observed.status))
        )
          throw new Error('presentation_existing_page_conflict')
        else if (
          action === 'discard' &&
          ((record.state !== 'staged' && record.state !== 'discard_pending') ||
            !observed ||
            !['staged', 'baseline'].includes(observed.status))
        )
          throw new Error('presentation_existing_page_conflict')
        else if (
          action === 'undo' &&
          ((record.state !== 'applied' &&
            record.state !== 'undo_pending' &&
            record.state !== 'restore_inserted') ||
            !observed ||
            !['applied', 'restore_staged', 'undone'].includes(observed.status))
        )
          throw new Error('presentation_existing_page_conflict')
        const initial = observed?.status
        const loadBackup = () =>
          loadRetainedPackage(record, record.backup, record.originalPackageDigest)
        const proposal = options.proposals.propose({
          operation: call.name,
          toolName: call.name,
          title: reapplying
            ? '重新应用现稿单页重建（暂存）'
            : action === 'stage'
              ? (call.input.explanation as string) || '暂存现稿单页重建'
              : `${action} 现稿单页重建`,
          preview: {
            oldSlideId: record.oldSlideId,
            newSlideId: record.newSlideId,
            originalDigest: record.originalPackageDigest,
            replacementDigest: record.replacementPackageDigest,
            stageKeepsOriginal: action === 'stage',
            restores: record.restores,
            ...(record.reapplies ? { reapplies: record.reapplies } : {}),
            restoredSlideGetsNewId: !!record.restores,
            qaPassed: false,
          },
          impact: { host: 'powerpoint', targets: [record.oldSlideId], count: 1 },
          fingerprint: selectionFingerprint(output(record)),
          validate: async () => {
            try {
              await sourceFresh()
              if (action !== 'stage') return same((await inspect()).status, initial)
              if (!(await baselineFresh())) return false
              const page = await options.exportAdapter.exportPresentationPagePackage(
                record.oldSlideId,
                signal,
              )
              await current()
              return (
                page.slideId === record.oldSlideId &&
                same(page.slideIds, record.beforeSlideIds) &&
                (await presentationPackageDigest(page.base64, signal)) ===
                  record.originalPackageDigest
              )
            } catch {
              return false
            }
          },
          execute: async () => {
            await sourceFresh()
            if (action === 'stage') {
              if (!(await baselineFresh())) throw new Error('proposal_stale')
              const exported = await options.exportAdapter.exportPresentationPagePackage(
                record.oldSlideId,
                signal,
              )
              await current()
              if (
                exported.slideId !== record.oldSlideId ||
                !same(exported.slideIds, record.beforeSlideIds)
              )
                throw new Error('proposal_stale')
              const originalBytes = bytes(exported.base64)
              if (!originalBytes.length || originalBytes.length > MAX_PPTX_PACKAGE_BYTES)
                throw new Error('presentation_page_backup_invalid')
              const originalPackageDigest = await presentationPackageDigest(exported.base64, signal)
              if (originalPackageDigest !== record.originalPackageDigest)
                throw new Error('proposal_stale')
              record = {
                ...record,
                backup: {
                  backupId: record.backup.backupId,
                  sha256: await sha(originalBytes),
                  sizeBytes: originalBytes.length,
                },
                sourceBackup: {
                  backupId: record.sourceBackup?.backupId ?? crypto.randomUUID(),
                  sha256: source!.sha256,
                  sizeBytes: source!.sizeBytes,
                },
              }
              try {
                await savePackage(
                  record,
                  record.backup,
                  originalBytes,
                  record.originalPackageDigest,
                )
                await savePackage(
                  record,
                  record.sourceBackup!,
                  bytes(source!.base64),
                  record.replacementPackageDigest,
                )
                await sourceFresh()
                if (!(await baselineFresh())) throw new Error('proposal_stale')
                const beforeWrite = await options.exportAdapter.exportPresentationPagePackage(
                  record.oldSlideId,
                  signal,
                )
                if (
                  beforeWrite.slideId !== record.oldSlideId ||
                  !same(beforeWrite.slideIds, record.beforeSlideIds) ||
                  (await presentationPackageDigest(beforeWrite.base64, signal)) !==
                    record.originalPackageDigest
                )
                  throw new Error('proposal_stale')
                await store(record)
              } catch (error) {
                if (
                  (await options.documentId()) === documentId &&
                  !options.readExistingPageChange(record.changeId)
                ) {
                  for (const metadata of [record.backup, record.sourceBackup!]) {
                    try {
                      await releasePackage(record, metadata)
                    } catch {
                      /* Preserve uncertain backup receipts for manual recovery. */
                    }
                  }
                }
                throw error
              }
              await sourceFresh()
              if (!(await baselineFresh())) throw new Error('proposal_stale')
              await options.adapter.stage(
                projected(record),
                source!.base64,
                async (newSlideId) => {
                  await store({ ...record, state: 'inserted', newSlideId })
                },
                assertCurrent,
                signal,
              )
              await store({ ...record, state: 'staged' })
            } else if (action === 'resume') {
              if (record.state === 'inserted' && initial === 'staged')
                await store({ ...record, state: 'staged' })
              else if (record.state === 'commit_pending' && initial === 'applied')
                await store({ ...record, state: 'applied' })
              else if (record.state === 'discard_pending' && initial === 'baseline')
                await store({ ...record, state: 'discarded' })
              else if (record.state === 'restore_inserted' && initial === 'undone')
                await store({ ...record, state: 'undone' })
              else throw new Error('presentation_existing_page_manual_review')
            } else if (action === 'commit') {
              await loadBackup()
              if ((await inspect()).status !== initial) throw new Error('proposal_stale')
              if (record.state === 'staged')
                await store({
                  ...record,
                  state: 'commit_pending',
                  captures: undefined,
                  reviews: undefined,
                })
              await options.adapter.commit(projected(record), assertCurrent, signal)
              await store({ ...record, state: 'applied' })
            } else if (action === 'discard') {
              if (record.state === 'staged')
                await store({
                  ...record,
                  state: 'discard_pending',
                  captures: undefined,
                  reviews: undefined,
                })
              await options.adapter.discard(projected(record), assertCurrent, signal)
              await store({ ...record, state: 'discarded' })
            } else {
              const backup = await loadBackup()
              if ((await inspect()).status !== initial) throw new Error('proposal_stale')
              if (record.state === 'applied')
                await store({
                  ...record,
                  state: 'undo_pending',
                  captures: undefined,
                  reviews: undefined,
                })
              await options.adapter.undo(
                projected(record),
                backup,
                async (restoredSlideId) => {
                  await store({ ...record, state: 'restore_inserted', restoredSlideId })
                },
                assertCurrent,
                signal,
              )
              await store({ ...record, state: 'undone' })
            }
          },
          verify: async () => {
            saved()
            const status = (await options.adapter.inspect(projected(record), signal)).status
            if (
              status !==
              (action === 'stage'
                ? 'staged'
                : action === 'discard'
                  ? 'baseline'
                  : action === 'undo'
                    ? 'undone'
                    : action === 'resume'
                      ? initial
                      : 'applied')
            )
              throw new Error('office_state_uncertain')
          },
          postWrite: async (): Promise<ProposalPostWriteEvidence> => {
            const targets =
              record.state === 'staged'
                ? [record.oldSlideId, record.newSlideId!]
                : record.state === 'applied'
                  ? [record.newSlideId!]
                  : record.state === 'discarded'
                    ? [record.oldSlideId]
                    : record.state === 'undone'
                      ? [record.restoredSlideId!]
                      : []
            if (!targets.length) throw new Error('office_state_uncertain')
            const status =
              record.state === 'staged'
                ? 'staged'
                : record.state === 'applied'
                  ? 'applied'
                  : record.state === 'discarded'
                    ? 'baseline'
                    : 'undone'
            const check = async () => {
              await current()
              saved()
              const observed = await options.adapter.inspect(projected(record))
              await current()
              saved()
              if (
                observed.status !== status ||
                !targets.every((id) => observed.slideIds.includes(id))
              )
                throw new Error('office_state_uncertain')
              return observed.slideIds
            }
            const before = await check()
            if (record.pictureTarget) {
              try {
                const targetSlideId =
                  record.state === 'staged' || record.state === 'applied'
                    ? record.newSlideId!
                    : record.state === 'undone'
                      ? record.restoredSlideId!
                      : record.oldSlideId
                const exported =
                  await options.exportAdapter.exportPresentationPagePackage(targetSlideId)
                await current()
                saved()
                if (exported.slideId !== targetSlideId || !same(exported.slideIds, before))
                  throw new Error('office_state_uncertain')
                const picture = await resolvePowerPointPictureIdentity(exported.base64, {
                  name: record.pictureTarget.name,
                })
                const expectedDigest =
                  record.state === 'staged' || record.state === 'applied'
                    ? record.pictureTarget.afterDigest
                    : record.pictureTarget.beforeDigest
                if (picture.mediaDigest !== expectedDigest || !same(await check(), before))
                  throw new Error('office_state_uncertain')
              } catch {
                return { status: 'unavailable', reason: 'picture_readback_failed' }
              }
            }
            const pages = [] as { slideId: string; pngBase64: string; digest: string }[]
            for (const slideId of targets) {
              const shot = await options.inspectPage(slideId)
              if (
                shot.slideId !== slideId ||
                shot.shapesTruncated ||
                shot.screenshot.mime !== 'image/png'
              )
                throw new Error('office_read_failed')
              const pngBase64 = validatePowerPointPageScreenshot(shot.screenshot.base64)
              pages.push({ slideId, pngBase64, digest: await sha(bytes(pngBase64)) })
              if (!same(await check(), before)) throw new Error('office_state_uncertain')
            }
            const capturedAt = new Date().toISOString()
            await store({
              ...record,
              captures: pages.map((page) => ({
                hostSlideId: page.slideId,
                screenshotDigest: page.digest,
                capturedAt,
              })),
            })
            return { status: 'captured', pages }
          },
        })
        return {
          output: output({
            proposalId: proposal.id,
            status: 'awaiting_confirmation',
            changeId: record.changeId,
            ...(record.reapplies
              ? { reapplies: record.reapplies, sourceChangeId: record.reapplies }
              : {}),
            state: record.state,
            oldSlideId: record.oldSlideId,
            newSlideId: record.newSlideId,
          }),
          mutated: false,
          summary: '已准备现稿单页变更，等待确认',
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : ''
        return {
          output:
            /^(presentation_[a-z_]+|office_[a-z_]+|invalid_tool_input|proposal_stale|cancelled|vfs_[a-z_]+)$/.test(
              message,
            )
              ? message
              : 'presentation_existing_page_failed',
          isError: true,
          mutated: false,
          summary: '现稿单页操作未完成',
        }
      }
    },
  }
}
