import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import {
  selectionFingerprint,
  type ProposalPostWriteEvidence,
  type StructuredProposalController,
} from '../../agent/proposal-controller.js'
import { validatePowerPointPageScreenshot, type PowerPointPageInspection } from './browser-powerpoint-adapter.js'
import {
  MAX_IMPORT_BYTES,
  readBoundedImage,
  supportsBrowserMediaValidation,
} from '../shared/import-media.js'
import type { InMemoryVfs } from '../shared/vfs.js'
import type { PresentationBaselineSkill } from './presentation-baseline.js'
import type {
  BrowserPresentationImageAdapter,
  PictureSnapshot,
} from './browser-presentation-image-adapter.js'
import type {
  PresentationImageBackup,
  ImageReplacementRecord,
} from './presentation-image-replacement-record.js'
import {
  validatePresentationExistingImageChange,
  type PresentationExistingImageChange,
} from './presentation-existing-image.js'

interface Options {
  baseline: PresentationBaselineSkill
  imageAdapter: Pick<
    BrowserPresentationImageAdapter,
    'inspect' | 'captureOriginal' | 'replace' | 'inspectRecovery' | 'finishRecovery'
  >
  imageBackup: PresentationImageBackup
  inspectPage(slideId: string, signal?: AbortSignal): Promise<Pick<PowerPointPageInspection, 'slideId' | 'shapesTruncated' | 'screenshot'>>
  vfs: InMemoryVfs
  proposals: StructuredProposalController
  documentId(): Promise<string>
  readExistingImageChange(id: string): PresentationExistingImageChange | undefined
  writeExistingImageChange(
    record: PresentationExistingImageChange,
    expected: PresentationExistingImageChange | undefined,
  ): Promise<void>
}
const id = { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' }
const hostId = { type: 'string', minLength: 1, maxLength: 256 }
const tools: AgentToolDef[] = [
  {
    name: 'replace_existing_presentation_image',
    description:
      'Propose a confirmed replacement of one ordinary native picture in an existing deck from a fresh scoped baseline and validated VFS PNG/JPEG. Original bytes are durably backed up before the savepoint and host insertion.',
    inputSchema: {
      type: 'object',
      properties: {
        baseline_id: id,
        slide_id: hostId,
        shape_id: hostId,
        path: { type: 'string', minLength: 1, maxLength: 1024 },
        explanation: { type: 'string', maxLength: 300 },
      },
      required: ['baseline_id', 'slide_id', 'shape_id', 'path'],
      additionalProperties: false,
    },
  },
  ...(['inspect', 'resume', 'undo'] as const).map((action): AgentToolDef => ({
    name: `${action}_existing_presentation_image_change`,
    description:
      action === 'inspect'
        ? 'Read current native picture evidence for this saved existing-deck image change without writing.'
        : 'Propose confirmed recovery or undo of an exact saved native image change. Ambiguous candidates require manual review; insertion is never replayed.',
    inputSchema: {
      type: 'object',
      properties: { change_id: id },
      required: ['change_id'],
      additionalProperties: false,
    },
  })),
  ...(['capture', 'record'] as const).map((action): AgentToolDef => ({
    name: `${action}_existing_presentation_image_review`,
    description: action === 'capture' ? 'Capture the saved image replacement page for visual judgment; does not pass QA.' : 'Persist a visual judgment for an unchanged captured screenshot.',
    inputSchema: { type: 'object', properties: {
      change_id: id,
      ...(action === 'record' ? { screenshot_digest: { type: 'string' }, status: { type: 'string', enum: ['pass', 'fail'] }, notes: { type: 'string', maxLength: 2000 } } : {}),
    }, required: action === 'record' ? ['change_id', 'screenshot_digest', 'status', 'notes'] : ['change_id'], additionalProperties: false },
  })),
]
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const result = (v: unknown) => {
  const s = JSON.stringify(v)
  if (new TextEncoder().encode(s).byteLength > 256 * 1024)
    throw new Error('presentation_existing_image_output_limit')
  return s
}
const hash = async (bytes: Uint8Array) =>
  Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes).buffer)),
    (b) => b.toString(16).padStart(2, '0'),
  ).join('')
const decode = (base64: string) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
const pictureIds = (before: PictureSnapshot, oldId: string, newId: string) =>
  before.shapeIds.map((id) => (id === oldId ? newId : id))
function bridge(r: PresentationExistingImageChange, reverse = false): ImageReplacementRecord {
  return reverse
    ? {
        version: 1,
        documentId: r.documentId,
        projectId: 'existing',
        requestId: r.changeId,
        pageId: 'existing',
        hostSlideId: r.hostSlideId,
        oldShapeId: r.insertedShapeId!,
        assetDigest: r.original.mediaDigest,
        state: 'pending',
        baseline: r.undoBaseline!,
        newShapeId: r.restoredShapeId,
      }
    : {
        version: 1,
        documentId: r.documentId,
        projectId: 'existing',
        requestId: r.changeId,
        pageId: 'existing',
        hostSlideId: r.hostSlideId,
        oldShapeId: r.oldShapeId,
        assetDigest: r.assetDigest,
        state: 'pending',
        baseline: r.original,
        backup: r.backup,
        newShapeId: r.insertedShapeId,
      }
}
const afterPicture = (
  record: PresentationExistingImageChange,
  picture: PictureSnapshot,
  reverse = false,
) => {
  const before = reverse ? record.undoBaseline! : record.original
  const oldId = reverse ? record.insertedShapeId! : record.oldShapeId
  const newId = reverse ? record.restoredShapeId! : record.insertedShapeId!
  const digest = reverse ? record.original.mediaDigest : record.assetDigest
  if (
    picture.slideId !== record.hostSlideId ||
    picture.shapeId !== newId ||
    picture.mediaDigest !== digest ||
    picture.zOrderPosition !== before.zOrderPosition ||
    !same(picture.shapeIds, pictureIds(before, oldId, newId)) ||
    picture.name !== before.name ||
    picture.altTextTitle !== before.altTextTitle ||
    picture.altTextDescription !== before.altTextDescription ||
    Math.abs(picture.rotation - before.rotation) > 0.01 ||
    (['left', 'top', 'width', 'height'] as const).some(
      (key) => Math.abs(picture.geometry[key] - before.geometry[key]) > 0.01,
    )
  )
    throw new Error('office_state_uncertain')
  return picture
}

export function createPresentationExistingImageEditingSkill(
  options: Options,
): AgentSkill & { clear(): void } {
  let epoch = 0
  let reviewCapture: { changeId: string; record: string; digest: string; capturedAt: string; epoch: number } | undefined
  return {
    id: 'presentation-existing-image-editing',
    get tools() {
      return tools.filter(
        (tool) =>
          tool.name === 'inspect_existing_presentation_image_change' ||
          tool.name.endsWith('_existing_presentation_image_review') ||
          (options.imageBackup.available() && supportsBrowserMediaValidation()),
      )
    },
    systemPrompt:
      'For existing native pictures, read_presentation_baseline, then replace_existing_presentation_image with an exact native slide/shape ID and validated VFS PNG/JPEG. A confirmed proposal backs up original bytes before the host write. New native picture IDs differ. For interrupted writes inspect then resume only an identified candidate; never retry insertion automatically. Undo requires the original backup and exact after snapshot. After a confirmed write, capture_existing_presentation_image_review, inspect the displayed image, then record_existing_presentation_image_review with its screenshot_digest and pass/fail notes. Inspect compares a current screenshot with the historical capture when possible; a match never certifies current or whole-deck QA.',
    clear() {
      epoch++
      reviewCapture = undefined
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
        const schema = tool.inputSchema as {
          properties: Record<string, unknown>
          required: string[]
        }
        if (
          Object.keys(call.input).some((k) => !Object.hasOwn(schema.properties, k)) ||
          schema.required.some((k) => !Object.hasOwn(call.input, k))
        )
          throw new Error('invalid_tool_input')
        if (
          call.name !== 'inspect_existing_presentation_image_change' &&
          !call.name.endsWith('_existing_presentation_image_review') &&
          (!options.imageBackup.available() || !supportsBrowserMediaValidation())
        )
          throw new Error('presentation_image_backup_unavailable')
        const creating = call.name === 'replace_existing_presentation_image'
        const documentId = await options.documentId()
        active()
        const current = async () => {
          active()
          if ((await options.documentId()) !== documentId)
            throw new Error('presentation_document_changed')
          active()
        }
        let record: PresentationExistingImageChange
        let originalBaseline: ReturnType<PresentationBaselineSkill['snapshot']>
        let source: { path: string; base64: string; digest: string } | undefined
        if (creating) {
          const i = call.input
          if (
            typeof i.baseline_id !== 'string' ||
            !/^[A-Za-z0-9_-]{1,128}$/.test(i.baseline_id) ||
            typeof i.slide_id !== 'string' ||
            typeof i.shape_id !== 'string' ||
            typeof i.path !== 'string' ||
            !i.path.length ||
            i.path.length > 1024 ||
            (i.explanation !== undefined &&
              (typeof i.explanation !== 'string' || i.explanation.length > 300))
          )
            throw new Error('invalid_tool_input')
          const baseline = options.baseline.snapshot(i.baseline_id)
          if (
            !baseline ||
            baseline.documentId !== documentId ||
            !baseline.scope.slideIds.includes(i.slide_id) ||
            (baseline.scope.shapeIds && !baseline.scope.shapeIds.includes(i.shape_id))
          )
            throw new Error('presentation_existing_scope_mismatch')
          originalBaseline = structuredClone(baseline)
          const shape = baseline.pages
            .find((p) => p.slideId === i.slide_id)
            ?.shapes.find((s) => s.id === i.shape_id)
          if (!shape || shape.type !== 'Image')
            throw new Error('presentation_existing_target_unsupported')
          const image = await readBoundedImage(options.vfs, i.path)
          await current()
          const digest = await hash(decode(image.base64))
          await current()
          source = { path: i.path, base64: image.base64, digest }
          const original = await options.imageAdapter.inspect(i.slide_id, i.shape_id, signal)
          await current()
          if (
            original.slideId !== i.slide_id ||
            original.shapeId !== i.shape_id ||
            (['left', 'top', 'width', 'height'] as const).some(
              (key) => Math.abs(original.geometry[key] - shape[key]) > 0.01,
            )
          )
            throw new Error('presentation_existing_target_changed')
          // Backup metadata is acquired only on confirmation. The proposed preview contains no server write.
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
            hostSlideId: i.slide_id,
            oldShapeId: i.shape_id,
            assetDigest: digest,
            original,
            backup: { attachmentId: original.mediaDigest, sizeBytes: 1, mime: image.mime },
            state: 'pending',
          }
        } else {
          if (
            typeof call.input.change_id !== 'string' ||
            !/^[A-Za-z0-9_-]{1,128}$/.test(call.input.change_id)
          )
            throw new Error('invalid_tool_input')
          const saved = options.readExistingImageChange(call.input.change_id)
          if (
            !saved ||
            !validatePresentationExistingImageChange(saved) ||
            saved.documentId !== documentId
          )
            throw new Error('presentation_existing_image_missing')
          record = structuredClone(saved)
        }
        let expected = creating ? undefined : structuredClone(record)
        const saved = () => {
          if (!same(options.readExistingImageChange(record.changeId), expected))
            throw new Error('presentation_existing_image_stale')
        }
        const store = async (next: PresentationExistingImageChange) => {
          await current()
          saved()
          await options.writeExistingImageChange(next, expected)
          await current()
          if (!same(options.readExistingImageChange(next.changeId), next))
            throw new Error('office_state_uncertain')
          record = structuredClone(next)
          expected = structuredClone(next)
          reviewCapture = undefined
        }
        const freshBaseline = async () => {
          if (!creating || !originalBaseline) return true
          if (!same(originalBaseline, options.baseline.snapshot(originalBaseline.baselineId)))
            return false
          const check = await options.baseline.executeTool(
            {
              id: 'image-check',
              name: 'check_presentation_baseline',
              input: { baseline_id: originalBaseline.baselineId },
            },
            signal,
          )
          await current()
          return (
            !check.isError &&
            JSON.parse(check.output).unchanged &&
            same(originalBaseline, options.baseline.snapshot(originalBaseline.baselineId))
          )
        }
        const sourceUnchanged = async () => {
          if (!source) return
          const bytes = options.vfs.readBytes(source.path, { maxBytes: MAX_IMPORT_BYTES + 1 })
          if (bytes.length > MAX_IMPORT_BYTES || (await hash(bytes)) !== source.digest)
            throw new Error('proposal_stale')
          await current()
        }
        const hostOriginal = async () => {
          await current()
          saved()
          const picture = await options.imageAdapter.inspect(
            record.hostSlideId,
            record.oldShapeId,
            signal,
          )
          await current()
          saved()
          return same(picture, record.original)
        }
        if (creating && (!(await freshBaseline()) || !(await hostOriginal())))
          throw new Error('presentation_baseline_changed')
        const backupBytes = async () => {
          const base64 = await options.imageBackup.load(documentId, record.backup, signal)
          if ((await hash(decode(base64))) !== record.original.mediaDigest)
            throw new Error('presentation_image_backup_invalid')
          await current()
          return base64
        }
        if (call.name === 'inspect_existing_presentation_image_change') {
          if (record.state === 'pending' || record.state === 'undo_pending') {
            const inspection =
              record.state === 'pending'
                ? await options.imageAdapter.inspectRecovery(bridge(record), signal)
                : await options.imageAdapter.inspectRecovery(bridge(record, true), signal)
            await current()
            saved()
            return {
              output: result({
                changeId: record.changeId,
                state: record.state,
                inspection,
                currentHostVerified: true,
                qaPassed: false,
              }),
              mutated: false,
              summary: '已检查现稿图片恢复状态',
            }
          }
          const shapeId =
            record.state === 'complete' ? record.insertedShapeId! : record.restoredShapeId!
          const picture = await options.imageAdapter.inspect(record.hostSlideId, shapeId, signal)
          await current()
          saved()
          let status: 'not_pending' | 'manual_review' = 'manual_review'
          try {
            if (record.state === 'complete')
              status = same(picture, record.after) ? 'not_pending' : 'manual_review'
            else {
              afterPicture(record, picture, true)
              status = 'not_pending'
            }
          } catch {
            status = 'manual_review'
          }
          let visualReceipt: 'not_captured' | 'matched' | 'different' | 'unavailable' = record.capture ? 'unavailable' : 'not_captured'
          if (record.capture && status === 'not_pending') {
            try {
              const shot = await options.inspectPage(record.hostSlideId, signal)
              await current(); saved()
              if (shot.slideId !== record.hostSlideId || shot.shapesTruncated || shot.screenshot.mime !== 'image/png') throw new Error('office_read_failed')
              const png = validatePowerPointPageScreenshot(shot.screenshot.base64)
              const currentDigest = await hash(decode(png))
              const again = await options.imageAdapter.inspect(record.hostSlideId, shapeId, signal)
              await current(); saved()
              if (!same(again, picture)) throw new Error('office_state_uncertain')
              visualReceipt = currentDigest === record.capture.screenshotDigest ? 'matched' : 'different'
            } catch (error) {
              await current(); saved()
              visualReceipt = 'unavailable'
            }
          }
          return {
            output: result({
              changeId: record.changeId,
              state: record.state,
              status,
              visualReceipt,
              currentHostVerified: true,
              qaPassed: false,
            }),
            mutated: false,
            summary: '已检查现稿图片保存点',
          }
        }
        if (call.name.endsWith('_existing_presentation_image_review')) {
          if (!['complete', 'undone'].includes(record.state) || record.review)
            throw new Error('presentation_existing_image_state_invalid')
          const shapeId = record.state === 'complete' ? record.insertedShapeId! : record.restoredShapeId!
          const picture = await options.imageAdapter.inspect(record.hostSlideId, shapeId, signal)
          await current(); saved()
          if (record.state === 'complete') {
            if (!same(picture, record.after)) throw new Error('presentation_existing_image_conflict')
          } else afterPicture(record, picture, true)
          const shot = await options.inspectPage(record.hostSlideId, signal)
          await current(); saved()
          if (shot.slideId !== record.hostSlideId || shot.shapesTruncated || shot.screenshot.mime !== 'image/png') throw new Error('office_read_failed')
          const png = validatePowerPointPageScreenshot(shot.screenshot.base64)
          const screenshotDigest = await hash(decode(png))
          const again = await options.imageAdapter.inspect(record.hostSlideId, shapeId, signal)
          await current(); saved()
          if (!same(again, picture)) throw new Error('presentation_existing_image_conflict')
          if (call.name.startsWith('capture_')) {
            const capture = { hostSlideId: record.hostSlideId, screenshotDigest, capturedAt: new Date().toISOString() }
            await store({ ...record, capture })
            reviewCapture = { changeId: record.changeId, record: JSON.stringify(record), digest: screenshotDigest, capturedAt: capture.capturedAt, epoch }
            return { output: result({ changeId: record.changeId, hostSlideId: record.hostSlideId, screenshotDigest, qaPassed: false }), display: { kind: 'images', items: [{ url: `data:image/png;base64,${png}` }] }, mutated: false, summary: '已采集图片变更页，等待视觉判断' }
          }
          if (!reviewCapture || reviewCapture.epoch !== epoch || reviewCapture.changeId !== record.changeId || reviewCapture.record !== JSON.stringify(record) || reviewCapture.digest !== screenshotDigest || record.capture?.screenshotDigest !== screenshotDigest || record.capture.capturedAt !== reviewCapture.capturedAt || call.input.screenshot_digest !== screenshotDigest || !['pass', 'fail'].includes(call.input.status as string) || typeof call.input.notes !== 'string' || call.input.notes.length > 2000)
            throw new Error('presentation_existing_image_review_stale')
          const review = { hostSlideId: record.hostSlideId, screenshotDigest, capturedAt: reviewCapture.capturedAt, reviewedAt: new Date().toISOString(), status: call.input.status as 'pass' | 'fail', notes: call.input.notes }
          await store({ ...record, review })
          return { output: result({ changeId: record.changeId, historicalReview: review, qaPassed: false }), mutated: false, summary: '已保存图片变更页历史视觉判断' }
        }
        if (call.name === 'undo_existing_presentation_image_change' && record.state !== 'complete')
          throw new Error('presentation_existing_image_state_invalid')
        if (
          call.name === 'resume_existing_presentation_image_change' &&
          !['pending', 'undo_pending'].includes(record.state)
        )
          throw new Error('presentation_existing_image_state_invalid')
        let recoveryStatus: 'ready_to_finish' | 'already_applied' | undefined
        if (!creating && call.name === 'resume_existing_presentation_image_change') {
          const projected = bridge(record, record.state === 'undo_pending')
          const inspection = await options.imageAdapter.inspectRecovery(projected, signal)
          await current()
          saved()
          if (inspection.status === 'manual_review')
            throw new Error('presentation_existing_image_manual_review')
          recoveryStatus = inspection.status
        }
        if (!creating) await backupBytes()
        const proposal = options.proposals.propose({
          operation: call.name,
          toolName: call.name,
          title: creating
            ? (call.input.explanation as string) || '替换现稿图片'
            : call.name.startsWith('undo_')
              ? '撤销现稿图片替换'
              : '恢复现稿图片替换',
          preview: {
            hostSlideId: record.hostSlideId,
            oldShapeId: record.oldShapeId,
            originalDigest: record.original.mediaDigest,
            replacementDigest: record.assetDigest,
            scope: record.scope,
            createsNewNativeShapeId: true,
            recoveryStatus,
          },
          impact: { host: 'powerpoint', targets: [record.hostSlideId], count: 1 },
          fingerprint: selectionFingerprint(result(record)),
          validate: async () => {
            try {
              await sourceUnchanged()
              return (
                (await freshBaseline()) &&
                (creating
                  ? await hostOriginal()
                  : same(options.readExistingImageChange(record.changeId), expected) &&
                    (call.name.startsWith('undo_')
                      ? same(
                          await options.imageAdapter.inspect(
                            record.hostSlideId,
                            record.insertedShapeId!,
                            signal,
                          ),
                          record.after,
                        )
                      : (
                          await options.imageAdapter.inspectRecovery(
                            bridge(record, record.state === 'undo_pending'),
                            signal,
                          )
                        ).status === recoveryStatus))
              )
            } catch {
              return false
            }
          },
          execute: async () => {
            await sourceUnchanged()
            if (!(await freshBaseline())) throw new Error('proposal_stale')
            if (creating) {
              if (!(await hostOriginal())) throw new Error('proposal_stale')
              const captured = await options.imageAdapter.captureOriginal(
                record.hostSlideId,
                record.oldShapeId,
                signal,
              )
              if (
                !same(captured.snapshot, record.original) ||
                (await hash(decode(captured.base64))) !== record.original.mediaDigest
              )
                throw new Error('proposal_stale')
              const backup = await options.imageBackup.save(documentId, captured.base64, signal)
              if (backup.attachmentId !== record.original.mediaDigest)
                throw new Error('presentation_image_backup_invalid')
              await current()
              await sourceUnchanged()
              if (!(await freshBaseline()) || !(await hostOriginal()))
                throw new Error('proposal_stale')
              await store({ ...record, backup })
              await sourceUnchanged()
              if (!(await freshBaseline()) || !(await hostOriginal()))
                throw new Error('proposal_stale')
              const inserted = await options.imageAdapter.replace(
                record.hostSlideId,
                record.oldShapeId,
                source!.base64,
                record.original,
                async (id) => {
                  if (record.insertedShapeId || record.original.shapeIds.includes(id))
                    throw new Error('office_state_uncertain')
                  await store({ ...record, insertedShapeId: id })
                },
                signal,
              )
              if (inserted.shapeId !== record.insertedShapeId)
                throw new Error('office_state_uncertain')
              const picture = afterPicture(
                record,
                await options.imageAdapter.inspect(record.hostSlideId, inserted.shapeId, signal),
              )
              await store({ ...record, state: 'complete', after: picture })
            } else if (call.name === 'resume_existing_presentation_image_change') {
              const reverse = record.state === 'undo_pending'
              const projected = bridge(record, reverse)
              const inspection = await options.imageAdapter.inspectRecovery(projected, signal)
              if (inspection.status !== recoveryStatus) throw new Error('proposal_stale')
              const completed = await options.imageAdapter.finishRecovery(
                projected,
                recoveryStatus!,
                signal,
              )
              if (completed.shapeId !== (reverse ? record.restoredShapeId : record.insertedShapeId))
                throw new Error('office_state_uncertain')
              if (reverse) await store({ ...record, state: 'undone' })
              else {
                const picture = afterPicture(
                  record,
                  await options.imageAdapter.inspect(record.hostSlideId, completed.shapeId, signal),
                )
                await store({ ...record, state: 'complete', after: picture })
              }
            } else {
              if (
                !same(
                  await options.imageAdapter.inspect(
                    record.hostSlideId,
                    record.insertedShapeId!,
                    signal,
                  ),
                  record.after,
                )
              )
                throw new Error('proposal_stale')
              const bytes = await backupBytes()
              await store({ ...record, state: 'undo_pending', undoBaseline: record.after, capture: undefined, review: undefined })
              const restored = await options.imageAdapter.replace(
                record.hostSlideId,
                record.insertedShapeId!,
                bytes,
                record.after!,
                async (id) => {
                  if (record.restoredShapeId || record.after!.shapeIds.includes(id))
                    throw new Error('office_state_uncertain')
                  await store({ ...record, restoredShapeId: id })
                },
                signal,
              )
              if (restored.shapeId !== record.restoredShapeId)
                throw new Error('office_state_uncertain')
              afterPicture(
                record,
                await options.imageAdapter.inspect(record.hostSlideId, restored.shapeId, signal),
                true,
              )
              await store({ ...record, state: 'undone' })
            }
            await current()
          },
          verify: async () => {
            saved()
            if (!['complete', 'undone'].includes(record.state))
              throw new Error('office_verify_failed')
          },
          postWrite: async (): Promise<ProposalPostWriteEvidence> => {
            const check = async () => {
              await current()
              saved()
              const reverse = record.state === 'undone'
              if (!reverse && record.state !== 'complete') throw new Error('office_state_uncertain')
              const picture = await options.imageAdapter.inspect(
                record.hostSlideId,
                reverse ? record.restoredShapeId! : record.insertedShapeId!,
              )
              await current()
              saved()
              afterPicture(record, picture, reverse)
              if (!reverse && !same(picture, record.after)) throw new Error('office_state_uncertain')
            }
            await check()
            const shot = await options.inspectPage(record.hostSlideId)
            if (shot.slideId !== record.hostSlideId || shot.shapesTruncated || shot.screenshot.mime !== 'image/png')
              throw new Error('office_read_failed')
            const pngBase64 = validatePowerPointPageScreenshot(shot.screenshot.base64)
            await check()
            const digest = await hash(decode(pngBase64))
            await store({ ...record, capture: { hostSlideId: record.hostSlideId, screenshotDigest: digest, capturedAt: new Date().toISOString() } })
            return { status: 'captured', pages: [{ slideId: record.hostSlideId, pngBase64, digest }] }
          },
        })
        return {
          output: result({
            proposalId: proposal.id,
            changeId: record.changeId,
            status: 'awaiting_confirmation',
          }),
          mutated: false,
          summary: '已准备现稿图片修改提案，等待确认',
        }
      } catch (error) {
        const code = error instanceof Error ? error.message : ''
        return {
          output:
            /^(presentation_[a-z_]+|office_[a-z_]+|invalid_tool_input|cancelled|proposal_stale)$/.test(
              code,
            )
              ? code
              : 'presentation_existing_image_operation_failed',
          isError: true,
          mutated: false,
          summary: '现稿图片操作未完成；未自动重试',
        }
      }
    },
  }
}
