import {
  selectionFingerprint,
  type StructuredProposalController,
} from '../../agent/proposal-controller.js'
import {
  validatePresentationExistingBatch,
  type PresentationExistingBatch,
} from './presentation-existing-batch.js'
interface Options {
  proposals: StructuredProposalController
  documentId(): Promise<string>
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  readExistingBatch(id: string): PresentationExistingBatch | undefined
  writeExistingBatch(
    next: PresentationExistingBatch,
    expected: PresentationExistingBatch | undefined,
  ): Promise<void>
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const abort = (signal?: AbortSignal) => {
  if (signal?.aborted) throw Error('cancelled')
}
const fail = (): never => {
  throw Error('presentation_native_add_backup_release_failed')
}
/** Releases only this ended native-add record's backup after explicit confirmation. */
export function createPresentationNativeAddRelease(options: Options) {
  return {
    async propose(changeId: string, signal?: AbortSignal) {
      abort(signal)
      const documentId = await options.documentId()
      abort(signal)
      const saved = options.readExistingBatch(changeId)
      if (
        !validatePresentationExistingBatch(saved) ||
        saved.version !== 2 ||
        saved.changeId !== changeId ||
        saved.documentId !== documentId
      )
        throw Error('presentation_native_add_conflict')
      if (saved.state !== 'undone' || saved.backupReleasedAt)
        throw Error('presentation_native_add_state_invalid')
      let record = structuredClone(saved)
      const backup = record.backups[0]!
      const current = async (s?: AbortSignal) => {
        abort(s)
        if ((await options.documentId()) !== documentId)
          throw Error('presentation_native_add_conflict')
        abort(s)
        if (!same(options.readExistingBatch(changeId), record))
          throw Error('presentation_native_add_conflict')
      }
      await current(signal)
      return options.proposals.propose({
        operation: 'release_slide_ir_addition',
        toolName: 'release_slide_ir_addition',
        title: '释放已恢复原生添加记录的原页备份',
        preview: {
          changeId,
          state: record.state,
          backupId: backup.backupId,
          sizeBytes: backup.sizeBytes,
          backupCount: 1,
          irreversible: true,
        },
        impact: { host: 'powerpoint', targets: [`backup:${backup.backupId}`], count: 1 },
        fingerprint: selectionFingerprint(JSON.stringify(record)),
        validate: async (s) => {
          try {
            await current(s)
            return true
          } catch {
            return false
          }
        },
        execute: async (s) => {
          await current(s)
          let response: Response
          try {
            response = await options.request(
              {
                operation: 'existing_page_backup_release',
                documentId,
                backupId: backup.backupId,
                hostSlideId: record.hostSlideId,
                slideIds: record.beforeSlideIds,
                sha256: backup.sha256,
                sizeBytes: backup.sizeBytes,
              },
              s,
            )
          } catch {
            await current(s)
            return fail()
          }
          await current(s)
          let receipt: unknown
          try {
            const text = await response.text()
            await current(s)
            if (new TextEncoder().encode(text).length > 16 * 1024) return fail()
            receipt = JSON.parse(text)
          } catch {
            await current(s)
            return fail()
          }
          if (!response.ok || !receipt || typeof receipt !== 'object' || Array.isArray(receipt))
            return fail()
          const r = receipt as Record<string, unknown>
          if (
            Object.keys(r).sort().join(',') !==
              'backupId,documentId,hostSlideId,sha256,sizeBytes,slideIds,status' ||
            r.status !== 'released' ||
            r.backupId !== backup.backupId ||
            r.documentId !== documentId ||
            r.hostSlideId !== record.hostSlideId ||
            !same(r.slideIds, record.beforeSlideIds) ||
            r.sha256 !== backup.sha256 ||
            r.sizeBytes !== backup.sizeBytes
          )
            return fail()
          await current(s)
          const next = { ...record, backupReleasedAt: new Date().toISOString() }
          try {
            await options.writeExistingBatch(next, record)
          } catch {
            await current(s)
            return fail()
          }
          record = next
          await current(s)
        },
        verify: async (s) => {
          await current(s)
          if (!record.backupReleasedAt) return fail()
        },
      })
    },
  }
}
