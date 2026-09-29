import type { PowerPointAdapter } from './browser-powerpoint-adapter.js'
import {
  validatePresentationExistingBatch,
  type PresentationExistingBatch,
  type PresentationNativeAddBatch,
} from './presentation-existing-batch.js'
import { readChartPackageBackup } from './presentation-chart-backup.js'
import { presentationPackageDigest } from './powerpoint-package.js'
import {
  observePowerPointNativeAdd,
  type NativeAddObservation,
} from './presentation-native-add-observation.js'

export interface NativeAddInspection {
  observation: NativeAddObservation
  packageDigest: string
  /** Actual Office SDK identities, never package cNvPr IDs. */
  createdShapeIds: string[]
}
interface Dependencies {
  documentId(): Promise<string>
  readExistingBatch(id: string): PresentationExistingBatch | undefined
  writeExistingBatch(
    next: PresentationExistingBatch,
    expected: PresentationExistingBatch | undefined,
  ): Promise<void>
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  adapter: Pick<
    PowerPointAdapter,
    'exportPresentationPagePackage' | 'listSlideShapes' | 'executeDeclarative'
  >
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const conflict = (): never => {
  throw Error('presentation_native_add_conflict')
}
const abort = (signal?: AbortSignal) => {
  if (signal?.aborted) throw Error('cancelled')
}

/** Metadata and bounded native writes only; caller owns proposal confirmation and the initial backup. */
export function createPresentationNativeAddExecution(options: Dependencies) {
  const read = async (
    changeId: string,
    signal?: AbortSignal,
  ): Promise<PresentationNativeAddBatch> => {
    abort(signal)
    const documentId = await options.documentId()
    abort(signal)
    const record = options.readExistingBatch(changeId)
    if (
      !record ||
      record.version !== 2 ||
      record.changeId !== changeId ||
      !validatePresentationExistingBatch(record)
    )
      throw Error('presentation_existing_batch_missing')
    if (record.documentId !== documentId) throw Error('presentation_document_changed')
    if (!['applying', 'applied'].includes(record.state)) conflict()
    return structuredClone(record)
  }
  const guard = async (record: PresentationNativeAddBatch, signal?: AbortSignal) => {
    abort(signal)
    if ((await options.documentId()) !== record.documentId)
      throw Error('presentation_document_changed')
    abort(signal)
    if (!same(options.readExistingBatch(record.changeId), record))
      throw Error('presentation_existing_batch_stale')
  }
  const inspectRecord = async (
    record: PresentationNativeAddBatch,
    signal?: AbortSignal,
  ): Promise<NativeAddInspection> => {
    await guard(record, signal)
    const exportPage = options.adapter.exportPresentationPagePackage
    if (!exportPage) throw Error('office_api_unsupported')
    const backup = record.backups[0]!
    const original = await readChartPackageBackup(
      {
        request: options.request,
        documentId: record.documentId,
        hostSlideId: record.hostSlideId,
        slideIds: record.beforeSlideIds,
        backup,
        expectedPackageDigest: record.baselineDigest,
      },
      signal,
    )
    await guard(record, signal)
    const before = await exportPage.call(options.adapter, record.hostSlideId, signal)
    await guard(record, signal)
    if (before.slideId !== record.hostSlideId || !same(before.slideIds, record.beforeSlideIds))
      conflict()
    const packageDigest = await presentationPackageDigest(before.base64, signal)
    const observation = await observePowerPointNativeAdd(
      original,
      before.base64,
      record.operations,
      signal,
    )
    await guard(record, signal)
    if (
      observation.status === 'conflict' ||
      observation.completedCount < record.nextIndex ||
      observation.completedCount > record.nextIndex + (record.inFlightIndex === undefined ? 0 : 1)
    )
      conflict()
    const sdk = await options.adapter.listSlideShapes(record.slideIndex, signal)
    await guard(record, signal)
    if (
      sdk.slideId !== record.hostSlideId ||
      sdk.slideIndex !== record.slideIndex ||
      sdk.shapes.length > 1000 ||
      new Set(sdk.shapes.map((shape) => shape.id)).size !== sdk.shapes.length
    )
      conflict()
    const createdShapeIds: string[] = []
    for (let index = 0; index < observation.completedCount; index++) {
      const operation = record.operations[index]!
      const matches = sdk.shapes.filter((shape) => shape.name === operation.name)
      if (matches.length !== 1) conflict()
      const shape = matches[0]!
      const type = {
        add_text_box: 'TextBox',
        add_geometric_shape: 'GeometricShape',
        add_native_table: 'Table',
      }[operation.op]
      if (
        shape.type !== type ||
        typeof shape.id !== 'string' ||
        !shape.id ||
        shape.id.length > 256 ||
        ![shape.left, shape.top, shape.width, shape.height].every(
          (value, i) =>
            Number.isFinite(value) &&
            Math.abs(
              value - [operation.left, operation.top, operation.width, operation.height][i]!,
            ) <= 0.01,
        ) ||
        (index < record.nextIndex && shape.id !== record.createdShapeIds[index])
      )
        conflict()
      createdShapeIds.push(shape.id)
    }
    const { inFlightIndex: _pending, ...identity } = record
    if (
      !validatePresentationExistingBatch({
        ...identity,
        nextIndex: observation.completedCount,
        createdShapeIds,
        state: observation.completedCount === record.operations.length ? 'applied' : 'applying',
      })
    )
      conflict()
    // Planned names outside the verified package prefix may not be hidden in SDK readback.
    if (
      record.operations
        .slice(observation.completedCount)
        .some((op) => sdk.shapes.some((shape) => shape.name === op.name))
    )
      conflict()
    const after = await exportPage.call(options.adapter, record.hostSlideId, signal)
    await guard(record, signal)
    if (
      after.slideId !== record.hostSlideId ||
      !same(after.slideIds, record.beforeSlideIds) ||
      (await presentationPackageDigest(after.base64, signal)) !== packageDigest
    )
      conflict()
    return { observation, packageDigest, createdShapeIds }
  }
  const store = async (
    next: PresentationNativeAddBatch,
    expected: PresentationNativeAddBatch,
    signal?: AbortSignal,
  ) => {
    await guard(expected, signal)
    if (!validatePresentationExistingBatch(next))
      throw Error('presentation_existing_batch_state_invalid')
    await options.writeExistingBatch(next, expected)
    await guard(next, signal)
  }
  return {
    async inspect(changeId: string, signal?: AbortSignal): Promise<NativeAddInspection> {
      return inspectRecord(await read(changeId, signal), signal)
    },
    async step(changeId: string, signal?: AbortSignal): Promise<PresentationNativeAddBatch> {
      let record = await read(changeId, signal)
      let proof = await inspectRecord(record, signal)
      if (record.state === 'applied') return record
      if (
        record.inFlightIndex !== undefined &&
        proof.observation.completedCount === record.nextIndex
      )
        throw Error('presentation_native_add_pending')
      if (record.inFlightIndex === undefined) {
        const pending = { ...record, inFlightIndex: record.nextIndex }
        await store(pending, record, signal)
        record = pending
        // The persisted inFlight is checked again against host state before one operation is queued.
        proof = await inspectRecord(record, signal)
        if (proof.observation.completedCount !== record.nextIndex) conflict()
        await options.adapter.executeDeclarative([record.operations[record.nextIndex]!], signal)
        await guard(record, signal)
        proof = await inspectRecord(record, signal)
      }
      if (proof.observation.completedCount !== record.nextIndex + 1) conflict()
      const { inFlightIndex: _inFlight, ...stable } = record
      const next: PresentationNativeAddBatch = {
        ...stable,
        nextIndex: record.nextIndex + 1,
        createdShapeIds: proof.createdShapeIds,
        state: record.nextIndex + 1 === record.operations.length ? 'applied' : 'applying',
      }
      await store(next, record, signal)
      return structuredClone(next)
    },
  }
}
