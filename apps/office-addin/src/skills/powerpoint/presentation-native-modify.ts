import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import {
  validatePowerPointPageScreenshot,
  type PowerPointAdapter,
} from './browser-powerpoint-adapter.js'
import {
  validatePresentationExistingBatch,
  type PresentationExistingBatch,
  type PresentationNativeModifyBatch,
  type NativeModifyOperation,
} from './presentation-existing-batch.js'
import {
  describePagePackageBackup,
  saveChartPackageBackup,
  readChartPackageBackup,
  cleanupUncommittedChartPackageBackup,
} from './presentation-chart-backup.js'
import { presentationPackageDigest } from './powerpoint-package.js'
import { readUntilConverged } from '../shared/office-write-transaction.js'
interface Options {
  documentId(): Promise<string>
  available?(): boolean
  adapter: PowerPointAdapter
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  proposals: StructuredProposalController
  readExistingBatch(id: string): PresentationExistingBatch | undefined
  writeExistingBatch(
    next: PresentationExistingBatch,
    expected: PresentationExistingBatch | undefined,
  ): Promise<void>
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const tools: AgentToolDef[] = ['inspect', 'resume'].map((action) => ({
  name: `${action}_native_modify_batch`,
  description:
    action === 'inspect'
      ? 'Inspect durable generic native edit progress. An uncertain write requires original-page restoration.'
      : 'Propose continuation of a durable native edit only when no write receipt is uncertain; checks every original backup and current package before writing.',
  inputSchema: {
    type: 'object',
    properties: { change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' } },
    required: ['change_id'],
    additionalProperties: false,
  },
}))
tools.push(
  ...(['capture_native_modify_page', 'record_native_modify_page_review'] as const).map((name) => ({
    name,
    description: name.startsWith('capture')
      ? 'Capture one applied generic native edit page after fresh package checks. No whole-deck QA claim.'
      : 'Save a bounded historical assessment of this session capture only after fresh screenshot and package readback.',
    inputSchema: {
      type: 'object',
      properties: {
        change_id: { type: 'string' },
        slide_id: { type: 'string' },
        ...(name.startsWith('record')
          ? {
              screenshot_digest: { type: 'string' },
              status: { type: 'string', enum: ['pass', 'fail'] },
              notes: { type: 'string', maxLength: 2000 },
            }
          : {}),
      },
      required: name.startsWith('record')
        ? ['change_id', 'slide_id', 'screenshot_digest', 'status', 'notes']
        : ['change_id', 'slide_id'],
      additionalProperties: false,
    },
  })),
)
export function createPresentationNativeModifySkill(options: Options) {
  let epoch = 0
  let visualEpoch = 0
  const captures = new Map<string, { digest: string; capturedAt: string; visualEpoch: number }>()
  const screenshotDigest = async (base64: string) =>
    Array.from(
      new Uint8Array(
        await crypto.subtle.digest(
          'SHA-256',
          Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)),
        ),
      ),
      (b) => b.toString(16).padStart(2, '0'),
    ).join('')
  const cancelled = (signal?: AbortSignal, token = epoch) => {
    if (signal?.aborted || token !== epoch) throw Error('cancelled')
  }
  const guard = async (
    documentId: string,
    signal?: AbortSignal,
    token = epoch,
    newWrite = false,
  ) => {
    cancelled(signal, token)
    if (newWrite && options.available?.() === false)
      throw Error('presentation_existing_persistence_unavailable')
    if ((await options.documentId()) !== documentId) throw Error('presentation_document_changed')
    cancelled(signal, token)
    if (newWrite && options.available?.() === false)
      throw Error('presentation_existing_persistence_unavailable')
  }
  const record = (changeId: string) => {
    const saved = options.readExistingBatch(changeId)
    if (!saved || saved.version !== 3 || !validatePresentationExistingBatch(saved))
      throw Error('presentation_existing_batch_missing')
    return structuredClone(saved)
  }
  const check = async (
    r: PresentationNativeModifyBatch,
    signal?: AbortSignal,
    token = epoch,
    newWrite = false,
  ) => {
    await guard(r.documentId, signal, token, newWrite)
    if (!same(options.readExistingBatch(r.changeId), r))
      throw Error('presentation_existing_batch_stale')
    if (!options.adapter.exportPresentationPagePackage) throw Error('office_api_unsupported')
    for (const p of r.pages) {
      const page = await options.adapter.exportPresentationPagePackage(p.hostSlideId, signal)
      const digest = await presentationPackageDigest(page.base64, signal)
      await guard(r.documentId, signal, token, newWrite)
      if (
        page.slideId !== p.hostSlideId ||
        !same(page.slideIds, r.beforeSlideIds) ||
        digest !== p.expectedPackageDigest
      )
        throw Error('presentation_baseline_changed')
    }
    if (!same(options.readExistingBatch(r.changeId), r))
      throw Error('presentation_existing_batch_stale')
  }
  const preflight = async (
    r: PresentationNativeModifyBatch,
    signal?: AbortSignal,
    token = epoch,
    newWrite = false,
  ) => {
    for (const p of r.pages) {
      const pending = r.operations
        .slice(r.nextIndex)
        .filter((op) => op.slide_index === p.slideIndex)
      if (!pending.length) continue
      const listed = await options.adapter.listSlideShapes(p.slideIndex, signal)
      await guard(r.documentId, signal, token, newWrite)
      if (
        listed.slideId !== p.hostSlideId ||
        listed.slideIndex !== p.slideIndex ||
        listed.shapes.length > 1000 ||
        new Set(listed.shapes.map((shape) => shape.id)).size !== listed.shapes.length
      )
        throw Error('office_read_failed')
      for (const op of pending) {
        const shape = listed.shapes.find((shape) => shape.id === op.shape_id)
        if (
          !shape ||
          !(['left', 'top', 'width', 'height'] as const).every((key) => Number.isFinite(shape[key]))
        )
          throw Error('office_read_failed')
        if (op.op === 'set_shape_text') {
          const read = await options.adapter.readSlideText(p.slideIndex, op.shape_id, signal)
          await guard(r.documentId, signal, token, newWrite)
          if (
            read.slideId !== p.hostSlideId ||
            read.shapeId !== op.shape_id ||
            typeof read.text !== 'string'
          )
            throw Error('office_read_failed')
        }
      }
    }
  }
  const backupCheck = async (
    r: PresentationNativeModifyBatch,
    signal?: AbortSignal,
    token = epoch,
    newWrite = false,
  ) => {
    for (const backup of r.backups) {
      await readChartPackageBackup(
        {
          request: options.request,
          documentId: r.documentId,
          hostSlideId: backup.hostSlideId,
          slideIds: r.beforeSlideIds,
          backup,
          expectedPackageDigest: backup.packageDigest,
        },
        signal,
      )
      await guard(r.documentId, signal, token, newWrite)
    }
  }
  const verifyTarget = async (
    r: PresentationNativeModifyBatch,
    op: NativeModifyOperation,
    slideId: string,
    signal?: AbortSignal,
    token = epoch,
    newWrite = false,
  ) => {
    const current = async () => {
      await guard(r.documentId, signal, token, newWrite)
      if (!same(options.readExistingBatch(r.changeId), r))
        throw Error('presentation_existing_batch_stale')
    }
    const verified = await readUntilConverged({
      signal,
      accept: Boolean,
      read: async () => {
        await current()
        const shapes = await options.adapter.listSlideShapes(op.slide_index, signal)
        await current()
        if (shapes.slideId !== slideId || shapes.slideIndex !== op.slide_index)
          throw Error('presentation_document_changed')
        if (
          shapes.shapes.length > 1000 ||
          new Set(shapes.shapes.map((shape) => shape.id)).size !== shapes.shapes.length
        )
          throw Error('office_read_failed')
        const shape = shapes.shapes.find((shape) => shape.id === op.shape_id)
        if (op.op === 'delete_shape') return !shape
        if (!shape) return false
        if (op.op === 'set_shape_text') {
          const text = await options.adapter.readSlideText(op.slide_index, op.shape_id, signal)
          await current()
          if (text.slideId !== slideId || text.shapeId !== op.shape_id)
            throw Error('presentation_document_changed')
          return text.text === op.text
        }
        return (['left', 'top', 'width', 'height'] as const).every(
          (key) => Number.isFinite(shape[key]) && Math.abs(shape[key] - op[key]) <= 0.01,
        )
      },
    })
    await current()
    if (!verified) throw Error('office_verify_failed')
  }
  const run = async (
    r: PresentationNativeModifyBatch,
    signal?: AbortSignal,
    token = epoch,
    newWrite = false,
  ) => {
    if (r.inFlightIndex !== undefined) throw Error('presentation_native_modify_write_uncertain')
    await preflight(r, signal, token, newWrite)
    await backupCheck(r, signal, token, newWrite)
    await check(r, signal, token, newWrite)
    while (r.state === 'applying') {
      await preflight(r, signal, token, newWrite)
      await check(r, signal, token, newWrite)
      const op = r.operations[r.nextIndex],
        p = r.pages.find((p) => p.slideIndex === op.slide_index)!
      const beforeShapes = await options.adapter.listSlideShapes(p.slideIndex, signal)
      await guard(r.documentId, signal, token, newWrite)
      if (beforeShapes.slideId !== p.hostSlideId || beforeShapes.slideIndex !== p.slideIndex)
        throw Error('presentation_document_changed')
      if (
        beforeShapes.shapes.length > 1000 ||
        new Set(beforeShapes.shapes.map((shape) => shape.id)).size !== beforeShapes.shapes.length
      )
        throw Error('office_read_failed')
      await check(r, signal, token, newWrite)
      const next = { ...r, inFlightIndex: r.nextIndex }
      await options.writeExistingBatch(next, r)
      r = next
      await check(r, signal, token, newWrite)
      await guard(r.documentId, signal, token, newWrite)
      await options.adapter.executeDeclarative([structuredClone(op)], signal)
      await guard(r.documentId, signal, token, newWrite)
      await verifyTarget(r, op, p.hostSlideId, signal, token, newWrite)
      const afterShapes = await options.adapter.listSlideShapes(p.slideIndex, signal)
      await guard(r.documentId, signal, token, newWrite)
      const otherShapes = (shapes: typeof beforeShapes.shapes) =>
        shapes.filter((shape) => shape.id !== op.shape_id).sort((a, b) => a.id.localeCompare(b.id))
      const beforeTarget = beforeShapes.shapes.find((shape) => shape.id === op.shape_id)
      const afterTarget = afterShapes.shapes.find((shape) => shape.id === op.shape_id)
      if (
        afterShapes.slideId !== p.hostSlideId ||
        afterShapes.slideIndex !== p.slideIndex ||
        afterShapes.shapes.length > 1000 ||
        new Set(afterShapes.shapes.map((shape) => shape.id)).size !== afterShapes.shapes.length ||
        !same(otherShapes(beforeShapes.shapes), otherShapes(afterShapes.shapes)) ||
        (op.op === 'delete_shape'
          ? Boolean(afterTarget)
          : !beforeTarget ||
            !afterTarget ||
            (op.op === 'set_shape_text'
              ? !same(beforeTarget, afterTarget)
              : !same(
                  [beforeTarget.id, beforeTarget.name, beforeTarget.type],
                  [afterTarget.id, afterTarget.name, afterTarget.type],
                )))
      )
        throw Error('office_verify_failed')
      const exported = await options.adapter.exportPresentationPagePackage!(p.hostSlideId, signal)
      await guard(r.documentId, signal, token, newWrite)
      if (exported.slideId !== p.hostSlideId || !same(exported.slideIds, r.beforeSlideIds))
        throw Error('presentation_document_changed')
      const digest = await presentationPackageDigest(exported.base64, signal)
      const { inFlightIndex: _inflight, ...rest } = r
      const acknowledged: PresentationNativeModifyBatch = {
        ...rest,
        nextIndex: r.nextIndex + 1,
        state: r.nextIndex + 1 === r.operations.length ? 'applied' : 'applying',
        pages: r.pages.map((page) =>
          page.hostSlideId === p.hostSlideId ? { ...page, expectedPackageDigest: digest } : page,
        ),
      }
      await guard(r.documentId, signal, token, newWrite)
      await options.writeExistingBatch(acknowledged, r)
      r = acknowledged
    }
    await check(r, signal, token, newWrite)
    return r
  }
  const proposeResume = (r: PresentationNativeModifyBatch, signal?: AbortSignal, token = epoch) => {
    if (r.inFlightIndex !== undefined) throw Error('presentation_native_modify_write_uncertain')
    if (r.state !== 'applying') throw Error('presentation_existing_batch_state_invalid')
    return options.proposals.propose({
      operation: 'resume_native_modify_batch',
      toolName: 'resume_native_modify_batch',
      title: r.intent,
      preview: {
        changeId: r.changeId,
        qaScope: { basis: 'native_modify_savepoints', hostSlideIds: r.scope.slideIds },
        nextIndex: r.nextIndex,
        operations: r.operations,
        atomic: false,
      },
      impact: {
        host: 'powerpoint',
        targets: r.scope.slideIds,
        count: r.operations.length - r.nextIndex,
      },
      fingerprint: selectionFingerprint(JSON.stringify(r)),
      validate: async (s) => {
        try {
          await check(r, s, token, true)
          await backupCheck(r, s, token, true)
          return true
        } catch {
          return false
        }
      },
      execute: async (s) => {
        await run(r, s, token, true)
      },
      verify: async (s) => {
        const done = record(r.changeId)
        if (done.state !== 'applied') throw Error('office_verify_failed')
        await check(done, s, token, true)
      },
    })
  }
  return {
    id: 'presentation-native-modify',
    systemPrompt:
      'All generic native writes require durable original page backups; uncertain writes require explicit original-page restoration.',
    name: 'presentation-native-modify',
    description:
      'Durable bounded generic native edits with original page savepoints and explicit recovery.',
    tools,
    clear() {
      epoch++
      visualEpoch++
      captures.clear()
    },
    beginMutation() {
      visualEpoch++
      captures.clear()
    },
    endMutation() {
      visualEpoch++
      captures.clear()
    },
    async propose(
      input: readonly NativeModifyOperation[],
      explanation?: string,
      signal?: AbortSignal,
    ) {
      const token = epoch
      cancelled(signal, token)
      if (options.available?.() === false)
        throw Error('presentation_existing_persistence_unavailable')
      const operations = structuredClone(input) as NativeModifyOperation[]
      if (
        !operations.length ||
        operations.length > 32 ||
        new Set(operations.map((o) => o.slide_index)).size > 8
      )
        throw Error('invalid_tool_input')
      if (!options.adapter.exportPresentationPagePackage) throw Error('office_api_unsupported')
      const documentId = await options.documentId()
      await guard(documentId, signal, token, true)
      const pages: PresentationNativeModifyBatch['pages'] = [],
        backups: PresentationNativeModifyBatch['backups'] = [],
        packages = new Map<string, string>()
      let beforeSlideIds: string[] | undefined
      for (const slideIndex of new Set(operations.map((o) => o.slide_index))) {
        const snapshot = await options.adapter.snapshotSlide(slideIndex, signal)
        await guard(documentId, signal, token, true)
        const page = await options.adapter.exportPresentationPagePackage(snapshot.slideId, signal)
        await guard(documentId, signal, token, true)
        if (
          page.slideId !== snapshot.slideId ||
          page.slideIds[slideIndex] !== snapshot.slideId ||
          (beforeSlideIds && !same(beforeSlideIds, page.slideIds))
        )
          throw Error('presentation_baseline_changed')
        beforeSlideIds = page.slideIds
        const metadata = await describePagePackageBackup(page.base64, signal)
        pages.push({
          hostSlideId: page.slideId,
          slideIndex,
          expectedPackageDigest: metadata.packageDigest,
        })
        backups.push({ hostSlideId: page.slideId, backupId: crypto.randomUUID(), ...metadata })
        packages.set(page.slideId, page.base64)
      }
      const r: PresentationNativeModifyBatch = {
        version: 3,
        kind: 'native_page_modify',
        changeId: crypto.randomUUID(),
        documentId,
        baselineId: crypto.randomUUID(),
        baselineDigest: backups[0].packageDigest,
        beforeSlideIds: beforeSlideIds!,
        scope: { slideIds: pages.map((p) => p.hostSlideId) },
        intent: explanation || 'Modify existing native PowerPoint objects',
        preserved: ['Original affected page packages'],
        validation: ['Native target readback', 'Durable package receipts'],
        risk: 'high',
        backups,
        operations,
        pages,
        nextIndex: 0,
        state: 'applying',
      }
      if (!validatePresentationExistingBatch(r)) throw Error('invalid_tool_input')
      const fresh = async (s?: AbortSignal) => {
        await guard(documentId, s, token, true)
        for (const p of pages) {
          const exported = await options.adapter.exportPresentationPagePackage!(p.hostSlideId, s)
          const digest = await presentationPackageDigest(exported.base64, s)
          await guard(documentId, s, token, true)
          if (
            exported.slideId !== p.hostSlideId ||
            !same(exported.slideIds, beforeSlideIds) ||
            digest !== p.expectedPackageDigest
          )
            throw Error('presentation_baseline_changed')
        }
      }
      await preflight(r, signal, token, true)
      await fresh(signal)
      const proposal = options.proposals.propose({
        operation: 'execute_office_js',
        toolName: 'execute_office_js',
        title: r.intent,
        preview: {
          changeId: r.changeId,
          qaScope: { basis: 'native_modify_savepoints', hostSlideIds: r.scope.slideIds },
          operations,
          originalPageSavepoints: backups.map((b) => b.hostSlideId),
          atomic: false,
          uncertainWriteRecovery:
            'Restore the original page; uncertain operations are never replayed.',
        },
        impact: { host: 'powerpoint', targets: r.scope.slideIds, count: operations.length },
        fingerprint: selectionFingerprint(JSON.stringify(r)),
        validate: async (s) => {
          try {
            await fresh(s)
            return true
          } catch {
            return false
          }
        },
        execute: async (s) => {
          await preflight(r, s, token, true)
          await fresh(s)
          try {
            for (const backup of backups) {
              const saved = await saveChartPackageBackup(
                {
                  request: options.request,
                  documentId,
                  hostSlideId: backup.hostSlideId,
                  slideIds: r.beforeSlideIds,
                  backupId: backup.backupId,
                  base64: packages.get(backup.hostSlideId)!,
                },
                s,
              )
              await guard(documentId, s, token, true)
              if (saved.sha256 !== backup.sha256 || saved.sizeBytes !== backup.sizeBytes)
                throw Error('presentation_chart_backup_invalid')
            }
            await backupCheck(r, s, token, true)
            await fresh(s)
            await options.writeExistingBatch(r, undefined)
          } catch (error) {
            let absent = false
            try {
              absent = options.readExistingBatch(r.changeId) === undefined
            } catch {
              // An unreadable intent may still own the backups.
            }
            if (absent)
              await Promise.allSettled(
                backups.map((backup) =>
                  cleanupUncommittedChartPackageBackup({
                    request: options.request,
                    documentId,
                    hostSlideId: backup.hostSlideId,
                    slideIds: r.beforeSlideIds,
                    backup,
                  }),
                ),
              )
            throw error
          }
          await run(r, s, token, true)
        },
        verify: async (s) => {
          const done = record(r.changeId)
          if (done.state !== 'applied') throw Error('office_verify_failed')
          await check(done, s, token, true)
        },
      })
      return { proposalId: proposal.id, changeId: r.changeId, status: 'awaiting_confirmation' }
    },
    async executeTool(call: Parameters<AgentSkill['executeTool']>[0], signal?: AbortSignal) {
      const token = epoch
      try {
        call = { ...call, input: structuredClone(call.input) }
        cancelled(signal, token)
        if (
          !tools.some((t) => t.name === call.name) ||
          call.inputError ||
          call.truncated ||
          Object.keys(call.input).some(
            (k) =>
              !Object.hasOwn(
                (tools.find((t) => t.name === call.name)!.inputSchema as any).properties,
                k,
              ),
          ) ||
          (tools.find((t) => t.name === call.name)!.inputSchema as any).required.some(
            (k: string) => !Object.hasOwn(call.input, k),
          ) ||
          typeof call.input.change_id !== 'string' ||
          !/^[A-Za-z0-9_-]{1,128}$/.test(call.input.change_id)
        )
          throw Error('invalid_tool_input')
        const r = record(call.input.change_id)
        await guard(r.documentId, signal, token)
        if (
          call.name === 'capture_native_modify_page' ||
          call.name === 'record_native_modify_page_review'
        ) {
          const i = call.input
          if (
            r.state !== 'applied' ||
            typeof i.slide_id !== 'string' ||
            !r.scope.slideIds.includes(i.slide_id) ||
            !options.adapter.inspectPresentationPage
          )
            throw Error('presentation_existing_batch_state_invalid')
          const captureKey = JSON.stringify([r.changeId, i.slide_id]),
            visualToken = visualEpoch
          const existing = captures.get(captureKey)
          const reviewing = call.name === 'record_native_modify_page_review'
          if (
            reviewing &&
            (!existing ||
              existing.visualEpoch !== visualToken ||
              typeof i.screenshot_digest !== 'string' ||
              i.screenshot_digest !== existing.digest ||
              !['pass', 'fail'].includes(String(i.status)) ||
              typeof i.notes !== 'string' ||
              i.notes.length > 2000)
          )
            throw Error('presentation_native_modify_capture_stale')
          await check(r, signal, token)
          const shot = await options.adapter.inspectPresentationPage(i.slide_id, signal)
          if (
            shot.slideId !== i.slide_id ||
            shot.shapesTruncated ||
            shot.screenshot.mime !== 'image/png'
          )
            throw Error('office_read_failed')
          const pngBase64 = validatePowerPointPageScreenshot(shot.screenshot.base64)
          const digest = await screenshotDigest(pngBase64)
          await check(r, signal, token)
          if (visualEpoch !== visualToken) throw Error('presentation_native_modify_capture_stale')
          if (!reviewing) {
            const capturedAt = new Date().toISOString()
            captures.set(captureKey, { digest, capturedAt, visualEpoch: visualToken })
            return {
              output: JSON.stringify({
                changeId: r.changeId,
                slideId: i.slide_id,
                screenshotDigest: digest,
                capturedAt,
                qaPassed: false,
                screenshot: { mime: 'image/png', base64: pngBase64 },
              }),
              mutated: false,
              summary: 'Captured applied native edit page',
            }
          }
          if (digest !== existing!.digest) throw Error('presentation_native_modify_capture_stale')
          const review = {
            hostSlideId: i.slide_id,
            screenshotDigest: digest,
            capturedAt: existing!.capturedAt,
            reviewedAt: new Date().toISOString(),
            status: i.status as 'pass' | 'fail',
            notes: i.notes as string,
          }
          const updated = {
            ...r,
            reviews: [
              ...(r.reviews ?? []).filter((review) => review.hostSlideId !== i.slide_id),
              review,
            ],
          }
          await options.writeExistingBatch(updated, r)
          await check(updated, signal, token)
          return {
            output: JSON.stringify({
              changeId: r.changeId,
              review,
              historical: true,
              qaPassed: false,
            }),
            mutated: false,
            summary: 'Saved historical native page screenshot assessment',
          }
        }
        if (call.name === 'inspect_native_modify_batch') {
          const restored = r.state === 'undone'
          const order = r.beforeSlideIds.map((id) =>
            restored && Object.hasOwn(r.restoredSlideIds!, id) ? r.restoredSlideIds![id] : id,
          )
          const current = async () => {
            await guard(r.documentId, signal, token)
            if (!same(options.readExistingBatch(r.changeId), r))
              throw Error('presentation_existing_batch_stale')
          }
          const pages = []
          for (const page of r.pages) {
            const hostSlideId = restored ? r.restoredSlideIds![page.hostSlideId] : page.hostSlideId
            const expectedPackageDigest = restored
              ? r.backups.find((backup) => backup.hostSlideId === page.hostSlideId)!.packageDigest
              : page.expectedPackageDigest
            const scoped = {
              originalHostSlideId: page.hostSlideId,
              hostSlideId,
              slideIndex: page.slideIndex,
              expectedPackageDigest,
            }
            await current()
            try {
              if (!options.adapter.exportPresentationPagePackage)
                throw Error('office_api_unsupported')
              const exported = await options.adapter.exportPresentationPagePackage(
                hostSlideId,
                signal,
              )
              await current()
              const observedDigest = await presentationPackageDigest(exported.base64, signal)
              await current()
              const status =
                exported.slideId === hostSlideId &&
                same(exported.slideIds, order) &&
                observedDigest === expectedPackageDigest
                  ? 'matched'
                  : 'changed'
              pages.push({ ...scoped, status, observedDigest })
            } catch {
              await current()
              pages.push({ ...scoped, status: 'unavailable' })
            }
          }
          await current()
          const currentPackageMatches = pages.every((page) => page.status === 'matched')
          return {
            output: JSON.stringify({
              changeId: r.changeId,
              state: r.state,
              nextIndex: r.nextIndex,
              inFlightIndex: r.inFlightIndex,
              pages,
              currentPackageMatches,
              currentHostVerified: currentPackageMatches,
              qaPassed: false,
              recovery:
                r.inFlightIndex !== undefined
                  ? 'original_page_restoration_required'
                  : r.state === 'applying'
                    ? currentPackageMatches
                      ? 'resume_or_original_page_restoration'
                      : 'original_page_restoration_required'
                    : r.state === 'applied'
                      ? 'original_page_restoration_available'
                      : r.state === 'undoing'
                        ? 'complete_original_page_restoration'
                        : 'terminal_original_pages_restored',
            }),
            mutated: false,
            summary: 'Inspected current native page packages without replaying uncertain writes',
          }
        }
        if (r.inFlightIndex !== undefined) throw Error('presentation_native_modify_write_uncertain')
        await check(r, signal, token)
        const proposed = proposeResume(r, signal, token)
        return {
          output: JSON.stringify({
            proposalId: proposed.id,
            changeId: r.changeId,
            status: 'awaiting_confirmation',
          }),
          mutated: false,
          summary: 'Proposed durable native edit continuation',
        }
      } catch (error) {
        return {
          output: error instanceof Error ? error.message : 'presentation_native_modify_failed',
          isError: true,
          mutated: false,
          summary: 'Native edit recovery stopped without replay',
        }
      }
    },
  }
}
