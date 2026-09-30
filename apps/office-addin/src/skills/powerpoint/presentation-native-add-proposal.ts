import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import type { PowerPointAdapter } from './browser-powerpoint-adapter.js'
import {
  validatePresentationExistingBatch,
  type PresentationExistingBatch,
  type PresentationNativeAddBatch,
  type NativeAddOperation,
} from './presentation-existing-batch.js'
import {
  describePagePackageBackup,
  saveChartPackageBackup,
  readChartPackageBackup,
  cleanupUncommittedChartPackageBackup,
} from './presentation-chart-backup.js'
import { observePowerPointNativeAdd } from './presentation-native-add-observation.js'
import { presentationPackageDigest } from './powerpoint-package.js'
import type { NativeAddInspection } from './presentation-native-add-execution.js'

interface Options {
  proposals: StructuredProposalController
  documentId(): Promise<string>
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  readExistingBatch(id: string): PresentationExistingBatch | undefined
  writeExistingBatch(
    next: PresentationExistingBatch,
    expected: PresentationExistingBatch | undefined,
  ): Promise<void>
  adapter: Pick<PowerPointAdapter, 'exportPresentationPagePackage' | 'listSlideShapes'>
  execution: {
    step(changeId: string, signal?: AbortSignal): Promise<PresentationNativeAddBatch>
    inspect(changeId: string, signal?: AbortSignal): Promise<NativeAddInspection>
  }
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const abort = (signal?: AbortSignal) => {
  if (signal?.aborted) throw Error('cancelled')
}
const stale = (): never => {
  throw Error('presentation_baseline_changed')
}

/** A fresh single-page proposal; backup and durable intent precede every native host write. */
export function createPresentationNativeAddProposal(options: Options) {
  return {
    async propose(
      input: readonly NativeAddOperation[],
      explanation?: string,
      signal?: AbortSignal,
      toolName: 'add_slide_ir_objects' | 'execute_office_js' = 'add_slide_ir_objects',
    ) {
      abort(signal)
      if (!['add_slide_ir_objects', 'execute_office_js'].includes(toolName))
        throw Error('invalid_tool_input')
      if (!options.adapter.exportPresentationPagePackage) throw Error('office_api_unsupported')
      if (!Array.isArray(input) || !input.length || input.length > 32)
        throw Error('invalid_tool_input')
      const operations = structuredClone(input) as NativeAddOperation[]
      if (
        operations.some(
          (operation) =>
            operation.op === 'add_text_box' &&
            (operation.fontFace === undefined ||
              operation.fontSize === undefined ||
              operation.color === undefined),
        )
      )
        throw Error('office_api_unsupported')
      const slideIndex = operations[0]!.slide_index
      if (
        !Number.isInteger(slideIndex) ||
        slideIndex < 0 ||
        slideIndex > 31 ||
        operations.some((op) => op.slide_index !== slideIndex)
      )
        throw Error('invalid_tool_input')
      if (
        explanation !== undefined &&
        (typeof explanation !== 'string' ||
          !explanation.trim() ||
          explanation.length > 300 ||
          Array.from(explanation).some((character) => character.charCodeAt(0) < 32))
      )
        throw Error('invalid_tool_input')
      const documentId = await options.documentId()
      abort(signal)
      const initial = await options.adapter.listSlideShapes(slideIndex, signal)
      abort(signal)
      if (initial.slideIndex !== slideIndex) stale()
      const hostSlideId = initial.slideId
      const page = await options.adapter.exportPresentationPagePackage(hostSlideId, signal)
      abort(signal)
      const beforeSlideIds = page.slideIds
      if (
        page.slideId !== hostSlideId ||
        !Array.isArray(beforeSlideIds) ||
        !beforeSlideIds.length ||
        beforeSlideIds.length > 500 ||
        beforeSlideIds[slideIndex] !== hostSlideId ||
        new Set(beforeSlideIds).size !== beforeSlideIds.length ||
        beforeSlideIds.some((id) => typeof id !== 'string' || !id || id.length > 256)
      )
        stale()
      const metadata = await describePagePackageBackup(page.base64, signal)
      const record: PresentationNativeAddBatch = {
        version: 2,
        kind: 'native_page_add',
        changeId: crypto.randomUUID(),
        documentId,
        baselineId: crypto.randomUUID(),
        baselineDigest: metadata.packageDigest,
        hostSlideId,
        slideIndex,
        beforeSlideIds,
        scope: { slideIds: [hostSlideId] },
        intent: explanation ?? '添加可编辑原生对象',
        preserved: ['保留原页对象及包依赖'],
        validation: ['包级结构与实际宿主对象读回'],
        risk: 'high',
        backups: [{ hostSlideId, backupId: crypto.randomUUID(), ...metadata }],
        operations,
        createdShapeIds: [],
        nextIndex: 0,
        state: 'applying',
      }
      if (!validatePresentationExistingBatch(record)) throw Error('invalid_tool_input')
      const current = async (checkSignal?: AbortSignal) => {
        abort(checkSignal)
        if ((await options.documentId()) !== documentId)
          throw Error('presentation_document_changed')
        abort(checkSignal)
        if (options.readExistingBatch(record.changeId) !== undefined)
          throw Error('presentation_existing_batch_stale')
      }
      const fresh = async (checkSignal?: AbortSignal) => {
        await current(checkSignal)
        const shapes = await options.adapter.listSlideShapes(slideIndex, checkSignal)
        await current(checkSignal)
        if (
          shapes.slideId !== hostSlideId ||
          shapes.slideIndex !== slideIndex ||
          operations.some((op) => shapes.shapes.some((s) => s.name === op.name))
        )
          stale()
        const exported = await options.adapter.exportPresentationPagePackage!(
          hostSlideId,
          checkSignal,
        )
        await current(checkSignal)
        if (
          exported.slideId !== hostSlideId ||
          !same(exported.slideIds, beforeSlideIds) ||
          (await presentationPackageDigest(exported.base64, checkSignal)) !== metadata.packageDigest
        )
          stale()
        const proof = await observePowerPointNativeAdd(
          page.base64,
          exported.base64,
          operations,
          checkSignal,
        )
        await current(checkSignal)
        if (proof.status !== 'none') stale()
      }
      await fresh(signal)
      return options.proposals.propose({
        operation: toolName,
        toolName,
        title: record.intent,
        preview: {
          changeId: record.changeId,
          slideId: hostSlideId,
          operations,
          savepoint: '原页包备份完成并读验后逐项添加；视觉与专业 QA 尚未完成',
        },
        impact: { host: 'powerpoint', targets: [hostSlideId], count: operations.length },
        fingerprint: selectionFingerprint(JSON.stringify(record)),
        validate: async (checkSignal) => {
          try {
            await fresh(checkSignal)
            return true
          } catch {
            return false
          }
        },
        execute: async (writeSignal) => {
          await fresh(writeSignal)
          const backup = record.backups[0]!
          const scope = {
            request: options.request,
            documentId,
            hostSlideId,
            slideIds: beforeSlideIds,
          }
          try {
            const saved = await saveChartPackageBackup(
              { ...scope, backupId: backup.backupId, base64: page.base64 },
              writeSignal,
            )
            await current(writeSignal)
            if (saved.sha256 !== backup.sha256 || saved.sizeBytes !== backup.sizeBytes)
              throw Error('presentation_chart_backup_invalid')
            await readChartPackageBackup(
              { ...scope, backup, expectedPackageDigest: metadata.packageDigest },
              writeSignal,
            )
            await fresh(writeSignal)
            await options.writeExistingBatch(record, undefined)
          } catch (error) {
            let absent = false
            try {
              absent = options.readExistingBatch(record.changeId) === undefined
            } catch {
              // An unreadable intent may still own the backup.
            }
            if (absent)
              await Promise.allSettled([cleanupUncommittedChartPackageBackup({ ...scope, backup })])
            throw error
          }
          abort(writeSignal)
          if (
            (await options.documentId()) !== documentId ||
            !same(options.readExistingBatch(record.changeId), record)
          )
            throw Error('presentation_existing_batch_stale')
          for (let index = 0; index < operations.length; index++) {
            abort(writeSignal)
            const next = await options.execution.step(record.changeId, writeSignal)
            if (
              next.nextIndex !== index + 1 ||
              next.changeId !== record.changeId ||
              (index + 1 === operations.length && next.state !== 'applied')
            )
              throw Error('office_state_uncertain')
          }
        },
        verify: async (verifySignal) => {
          abort(verifySignal)
          if ((await options.documentId()) !== documentId)
            throw Error('presentation_document_changed')
          const saved = options.readExistingBatch(record.changeId)
          if (
            !saved ||
            saved.version !== 2 ||
            saved.changeId !== record.changeId ||
            saved.documentId !== record.documentId ||
            saved.baselineDigest !== record.baselineDigest ||
            !same(saved.operations, operations) ||
            !same(saved.beforeSlideIds, beforeSlideIds) ||
            saved.state !== 'applied' ||
            saved.nextIndex !== operations.length ||
            !validatePresentationExistingBatch(saved)
          )
            throw Error('office_state_uncertain')
          const proof = await options.execution.inspect(record.changeId, verifySignal)
          abort(verifySignal)
          if (
            proof.observation.status !== 'complete' ||
            proof.observation.completedCount !== operations.length ||
            !same(proof.createdShapeIds, saved.createdShapeIds)
          )
            throw Error('office_verify_failed')
        },
      })
    },
  }
}
