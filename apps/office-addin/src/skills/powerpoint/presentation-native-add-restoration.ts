import {
  validatePresentationExistingBatch,
  type PresentationExistingBatch,
  type PresentationNativeAddBatch,
} from './presentation-existing-batch.js'
import {
  validatePresentationExistingPageChange,
  type PresentationExistingPageChange,
} from './presentation-existing-page.js'
import { readChartPackageBackup } from './presentation-chart-backup.js'
import { presentationPackageDigest } from './powerpoint-package.js'
interface Dependencies {
  documentId(): Promise<string>
  readExistingBatch(id: string): PresentationExistingBatch | undefined
  writeExistingBatch(
    next: PresentationExistingBatch,
    expected: PresentationExistingBatch | undefined,
  ): Promise<void>
  readExistingPageChange(id: string): PresentationExistingPageChange | undefined
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  exportPresentationPagePackage(
    slideId: string,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; slideIds: string[]; base64: string }>
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const conflict = (): never => {
  throw Error('presentation_native_add_conflict')
}
const abort = (signal?: AbortSignal) => {
  if (signal?.aborted) throw Error('cancelled')
}
/** Closes durable metadata after separately confirmed whole-page restoration; never writes the host. */
export function createPresentationNativeAddRestoration(options: Dependencies) {
  return {
    async finalize(
      changeId: string,
      restorationChangeId: string,
      signal?: AbortSignal,
    ): Promise<PresentationNativeAddBatch> {
      abort(signal)
      const documentId = await options.documentId()
      abort(signal)
      const saved = options.readExistingBatch(changeId),
        pageSaved = options.readExistingPageChange(restorationChangeId)
      if (
        !validatePresentationExistingBatch(saved) ||
        saved.version !== 2 ||
        saved.changeId !== changeId ||
        saved.documentId !== documentId ||
        !validatePresentationExistingPageChange(pageSaved) ||
        pageSaved.changeId !== restorationChangeId ||
        pageSaved.documentId !== documentId
      )
        return conflict()
      let record = structuredClone(saved)
      const page = structuredClone(pageSaved),
        backup = record.backups[0]!
      const expectedRestores = {
        sourceKind: 'batch',
        sourceChangeId: changeId,
        sourceHostSlideId: record.hostSlideId,
        originalBackupId: backup.backupId,
        originalPackageDigest: record.baselineDigest,
      }
      if (
        page.state !== 'applied' ||
        !page.newSlideId ||
        page.oldSlideId !== record.hostSlideId ||
        page.reapplies !== undefined ||
        !page.restores ||
        Object.entries(expectedRestores).some(
          ([key, value]) => page.restores![key as keyof typeof page.restores] !== value,
        ) ||
        page.replacementPackageDigest !== record.baselineDigest ||
        !same(page.beforeSlideIds, record.beforeSlideIds) ||
        !page.sourceBackup ||
        page.sourceBackup.sha256 !== backup.sha256 ||
        page.sourceBackup.sizeBytes !== backup.sizeBytes
      )
        conflict()
      const guard = async () => {
        abort(signal)
        if ((await options.documentId()) !== documentId) conflict()
        abort(signal)
        if (
          !same(options.readExistingBatch(changeId), record) ||
          !same(options.readExistingPageChange(restorationChangeId), page)
        )
          throw Error('presentation_existing_batch_stale')
      }
      await guard()
      if (record.state === 'undone') {
        if (record.restoredSlideId !== page.newSlideId) conflict()
        return structuredClone(record)
      }
      if (record.backupReleasedAt || page.backupReleasedAt) conflict()
      const guardedRequest: Dependencies['request'] = async (body, requestSignal) => {
        await guard()
        const response = await options.request(body, requestSignal)
        await guard()
        return response
      }
      const prove = async () => {
        await guard()
        await readChartPackageBackup(
          {
            request: guardedRequest,
            documentId,
            hostSlideId: record.hostSlideId,
            slideIds: record.beforeSlideIds,
            backup,
            expectedPackageDigest: record.baselineDigest,
          },
          signal,
        )
        await guard()
        await readChartPackageBackup(
          {
            request: guardedRequest,
            documentId,
            hostSlideId: page.oldSlideId,
            slideIds: page.beforeSlideIds,
            backup: page.sourceBackup!,
            expectedPackageDigest: record.baselineDigest,
          },
          signal,
        )
        await guard()
        const order = record.beforeSlideIds.map((id) =>
          id === record.hostSlideId ? page.newSlideId! : id,
        )
        for (let i = 0; i < 2; i++) {
          const exported = await options.exportPresentationPagePackage(page.newSlideId!, signal)
          await guard()
          if (exported.slideId !== page.newSlideId || !same(exported.slideIds, order)) conflict()
          const digest = await presentationPackageDigest(exported.base64, signal)
          await guard()
          if (digest !== record.baselineDigest) conflict()
        }
      }
      await prove()
      const store = async (next: PresentationNativeAddBatch) => {
        await guard()
        await options.writeExistingBatch(next, record)
        record = next
        await guard()
      }
      if (record.state !== 'undoing') {
        const { inFlightIndex: _flight, ...rest } = record
        await store({ ...rest, state: 'undoing' })
      }
      await prove()
      await store({
        ...record,
        state: 'undone',
        restoredSlideId: page.newSlideId,
      })
      return structuredClone(record)
    },
  }
}
