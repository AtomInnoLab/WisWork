import {
  imageReplacementKey,
  imageReplacementReservedBytes,
  validateImageReplacementRecord,
  type ImageReplacementRecord,
} from './presentation-image-replacement-record.js'
import {
  validatePresentationQaRecord,
  presentationQaMutationScope,
  presentationQaRecheckBytes,
  PRESENTATION_QA_RECHECK_FIELD_BYTES,
  type PresentationQaRecord,
} from './presentation-qa.js'
import { validPresentationImportRecord } from './presentation-page-delivery.js'
import type { PresentationImportRecord } from './presentation-delivery.js'
const ID_KEY = 'wiswork.presentation.document.v1'
const IMPORT_KEY = 'wiswork.presentation.imports.v1'
const IMAGE_KEY = 'wiswork.presentation.image-replacements.v1'
const QA_KEY = 'wiswork.presentation.qa.v1'
const PROJECT_KEY = 'wiswork.presentation.project.v1'
const validId = (value: unknown): value is string =>
  typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)

interface DocumentSettings {
  get(key: string): unknown
  set(key: string, value: string): void
  save(): Promise<void>
  location(): string
}

function validImportKey(key: string, record?: PresentationImportRecord): boolean {
  if (!/^(?:production\/)?[A-Za-z0-9_-]{1,128}\/[A-Za-z0-9_-]{1,128}$/.test(key)) return false
  return (
    record === undefined ||
    (key.split('/').length === 3
      ? record?.checkpoint?.version === 2
      : record?.checkpoint?.version !== 2)
  )
}

export function createPresentationDocumentBinding(
  settings: DocumentSettings,
  randomUUID: () => string = () => crypto.randomUUID(),
) {
  let pending: Promise<string> | undefined
  let unsavedId: string | undefined
  const documentId = (): Promise<string> => {
    if (pending) return pending
    pending = (async () => {
      try {
        const location = settings.location()
        if (location.length > 1800) throw new Error('invalid_location')
        let id = settings.get(ID_KEY)
        if (!validId(id)) {
          id = randomUUID()
          if (!validId(id)) throw new Error('invalid_identity')
          settings.set(ID_KEY, id)
          unsavedId = id
        }
        if (id === unsavedId) {
          await settings.save()
          unsavedId = undefined
        }
        if (settings.location() !== location || settings.get(ID_KEY) !== id)
          throw new Error('presentation_document_changed')
        // A Save As copy has a different binding even if it retains custom settings.
        return JSON.stringify([id, location])
      } catch (error) {
        if (error instanceof Error && error.message === 'presentation_document_changed') throw error
        throw new Error('presentation_document_identity_unavailable', { cause: error })
      }
    })().finally(() => {
      pending = undefined
    })
    return pending
  }
  let receiptWriteFailed = false
  let receiptQueue: Promise<void> = Promise.resolve()
  const readImports = (): Record<string, PresentationImportRecord> => {
    if (receiptWriteFailed) throw new Error('presentation_import_state_invalid')
    const raw = settings.get(IMPORT_KEY)
    if (raw === undefined || raw === null) return {}
    if (typeof raw !== 'string' || raw.length > 100_000)
      throw new Error('presentation_import_state_invalid')
    const value = JSON.parse(raw)
    if (
      !value ||
      typeof value !== 'object' ||
      Array.isArray(value) ||
      Object.keys(value).length > 32
    )
      throw new Error('presentation_import_state_invalid')
    for (const [key, record] of Object.entries(value)) {
      if (!validPresentationImportRecord(record) || !validImportKey(key, record))
        throw new Error('presentation_import_state_invalid')
    }
    return value
  }
  let qaWriteFailed = false
  const readQaRecords = (): Record<string, PresentationQaRecord> => {
    if (qaWriteFailed) throw new Error('presentation_qa_state_invalid')
    const raw = settings.get(QA_KEY)
    if (raw === undefined || raw === null) return {}
    if (
      typeof raw !== 'string' ||
      new TextEncoder().encode(raw).byteLength >
        256 * 1024 + 8 * 32 * PRESENTATION_QA_RECHECK_FIELD_BYTES
    )
      throw new Error('presentation_qa_state_invalid')
    let records: Record<string, PresentationQaRecord>
    try {
      records = JSON.parse(raw)
    } catch {
      throw new Error('presentation_qa_state_invalid')
    }
    if (
      !records ||
      typeof records !== 'object' ||
      Array.isArray(records) ||
      Object.keys(records).length > 8
    )
      throw new Error('presentation_qa_state_invalid')
    for (const [key, value] of Object.entries(records)) {
      if (
        !validatePresentationQaRecord(value) ||
        key !==
          `${value.source === 'production' ? 'production/' : ''}${value.projectId}/${value.requestId}`
      )
        throw new Error('presentation_qa_state_invalid')
    }
    if (
      new TextEncoder().encode(raw).byteLength -
        Object.values(records).reduce(
          (sum, record) => sum + presentationQaRecheckBytes(record),
          0,
        ) >
      256 * 1024
    )
      throw new Error('presentation_qa_state_invalid')
    return records
  }
  let imageWriteFailed = false
  const imageBytes = (raw: string, records: Record<string, ImageReplacementRecord>) =>
    new TextEncoder().encode(raw).byteLength +
    Object.values(records).reduce((sum, record) => sum + imageReplacementReservedBytes(record), 0)
  const readImageRecords = (): Record<string, ImageReplacementRecord> => {
    const invalid = () => new Error('presentation_image_replacement_state_invalid')
    if (imageWriteFailed) throw invalid()
    const raw = settings.get(IMAGE_KEY)
    if (raw === undefined || raw === null) return {}
    if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 128 * 1024)
      throw invalid()
    let records: Record<string, ImageReplacementRecord>
    try {
      records = JSON.parse(raw)
    } catch {
      throw invalid()
    }
    if (
      !records ||
      typeof records !== 'object' ||
      Array.isArray(records) ||
      Object.keys(records).length > 32
    )
      throw invalid()
    if (
      Object.entries(records).some(
        ([key, value]) => !/^[a-f0-9]{64}$/.test(key) || !validateImageReplacementRecord(value),
      )
    )
      throw invalid()
    if (imageBytes(raw, records) > 128 * 1024) throw invalid()
    return records
  }
  const saveQaRecords = async (records: Record<string, PresentationQaRecord>) => {
    if (Object.values(records).some((record) => !validatePresentationQaRecord(record)))
      throw new Error('presentation_qa_state_invalid')
    const previous = settings.get(QA_KEY),
      location = settings.location()
    const serialized = JSON.stringify(records)
    if (
      Object.keys(records).length > 8 ||
      new TextEncoder().encode(serialized).byteLength -
        Object.values(records).reduce(
          (sum, record) => sum + presentationQaRecheckBytes(record),
          0,
        ) >
        256 * 1024
    )
      throw new Error('presentation_qa_history_full')
    try {
      settings.set(QA_KEY, serialized)
      await settings.save()
      if (settings.location() !== location || settings.get(QA_KEY) !== serialized)
        throw new Error('presentation_document_changed')
    } catch (error) {
      if (settings.location() === location && settings.get(QA_KEY) === serialized) {
        try {
          settings.set(QA_KEY, typeof previous === 'string' ? previous : '{}')
        } catch {
          qaWriteFailed = true
        }
      } else qaWriteFailed = true
      throw error
    }
  }
  return {
    documentId,
    readReceipt: (key: string) => readImports()[key],
    writeReceipt(key: string, record: PresentationImportRecord | undefined) {
      const write = async () => {
        if (
          !validImportKey(key, record) ||
          (record !== undefined && !validPresentationImportRecord(record))
        )
          throw new Error('presentation_import_state_invalid')
        const imports = readImports()
        const previousRaw = settings.get(IMPORT_KEY)
        const location = settings.location()
        if (record) imports[key] = record
        else delete imports[key]
        const serialized = JSON.stringify(imports)
        if (Object.keys(imports).length > 32 || serialized.length > 100_000)
          throw new Error('presentation_import_history_full')
        try {
          settings.set(IMPORT_KEY, serialized)
          await settings.save()
          if (settings.location() !== location || settings.get(IMPORT_KEY) !== serialized)
            throw new Error('presentation_document_changed')
        } catch (error) {
          // A failed completion save must leave the earlier in-flight marker visible.
          // Initial reservations fail before any Office write, so restoring their prior state is safe.
          if (settings.location() === location && settings.get(IMPORT_KEY) === serialized) {
            try {
              settings.set(IMPORT_KEY, typeof previousRaw === 'string' ? previousRaw : '{}')
            } catch {
              receiptWriteFailed = true
            }
          } else receiptWriteFailed = true
          throw error
        }
      }
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
    readImageReplacement: (key: string) => readImageRecords()[key],
    writeImageReplacement(key: string, record: ImageReplacementRecord) {
      // Copy before queueing so callers cannot change the reservation during another save.
      const snapshot = structuredClone(record)
      const write = async () => {
        const invalid = () => new Error('presentation_image_replacement_state_invalid')
        if (
          !validateImageReplacementRecord(snapshot) ||
          key !==
            (await imageReplacementKey(
              snapshot.projectId,
              snapshot.requestId,
              snapshot.pageId,
              snapshot.oldShapeId,
              snapshot.source,
            ))
        )
          throw invalid()
        if ((await documentId()) !== snapshot.documentId)
          throw new Error('presentation_document_changed')
        const records = readImageRecords(),
          prior = records[key]
        if (!prior) {
          if (snapshot.state !== 'pending' || snapshot.newShapeId !== undefined) throw invalid()
        } else {
          const identity = (r: ImageReplacementRecord) =>
            JSON.stringify([
              r.documentId,
              r.source,
              r.projectId,
              r.requestId,
              r.pageId,
              r.hostSlideId,
              r.oldShapeId,
              r.assetDigest,
              r.baseline,
            ])
          if (
            identity(prior) !== identity(snapshot) ||
            (prior.newShapeId !== undefined && prior.newShapeId !== snapshot.newShapeId) ||
            (snapshot.state === 'complete' && !prior.newShapeId) ||
            (prior.state === 'complete' && snapshot.state !== 'complete')
          )
            throw invalid()
          if (prior.state === 'complete') return
        }
        records[key] = snapshot
        const serialized = JSON.stringify(records)
        if (Object.keys(records).length > 32 || imageBytes(serialized, records) > 128 * 1024)
          throw new Error('presentation_image_replacement_history_full')
        const previous = settings.get(IMAGE_KEY),
          location = settings.location()
        try {
          settings.set(IMAGE_KEY, serialized)
          await settings.save()
          if (settings.location() !== location || settings.get(IMAGE_KEY) !== serialized)
            throw new Error('presentation_document_changed')
        } catch (error) {
          if (settings.location() === location && settings.get(IMAGE_KEY) === serialized) {
            try {
              settings.set(IMAGE_KEY, typeof previous === 'string' ? previous : '{}')
            } catch {
              imageWriteFailed = true
            }
          } else imageWriteFailed = true
          throw error
        }
      }
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
    readQa: (key: string) => readQaRecords()[key],
    writeQa(key: string, record: PresentationQaRecord) {
      const write = async () => {
        if (
          !validatePresentationQaRecord(record) ||
          key !==
            `${record.source === 'production' ? 'production/' : ''}${record.projectId}/${record.requestId}`
        )
          throw new Error('presentation_qa_state_invalid')
        if ((await documentId()) !== record.documentId)
          throw new Error('presentation_document_changed')
        const records = readQaRecords()
        await saveQaRecords({ ...records, [key]: record })
      }
      // Serialize settings saves with import checkpoints so the two journals cannot race.
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
    invalidateQa(hostSlideIds?: readonly string[]) {
      const scope = presentationQaMutationScope(hostSlideIds)
      const matches = (hostSlideId: string) => scope === undefined || scope.has(hostSlideId)
      const write = async () => {
        const records = readQaRecords()
        if (
          !Object.values(records).some((record) =>
            record.pages.some((page) => matches(page.hostSlideId) && !page.recheckRequired),
          )
        )
          return
        // Unknown mutation scope remains conservatively document-wide.
        const invalidated = Object.fromEntries(
          Object.entries(records).map(([key, record]) => [
            key,
            {
              ...record,
              pages: record.pages.map((page) =>
                matches(page.hostSlideId) ? { ...page, recheckRequired: true as const } : page,
              ),
            },
          ]),
        )
        await saveQaRecords(invalidated)
      }
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
    lastProject(): string | undefined {
      const id = settings.get(PROJECT_KEY)
      return validId(id) ? id : undefined
    },
    async rememberProject(projectId: string) {
      if (!validId(projectId)) throw new Error('invalid_tool_input')
      settings.set(PROJECT_KEY, projectId)
      await settings.save()
    },
  }
}

export function createBrowserPresentationDocumentBinding() {
  return createPresentationDocumentBinding({
    get: (key) => Office.context.document.settings.get(key),
    set: (key, value) => Office.context.document.settings.set(key, value),
    location: () => Office.context.document.url ?? '',
    save: () =>
      new Promise<void>((resolve, reject) => {
        Office.context.document.settings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve()
          else reject(new Error('presentation_document_identity_unavailable'))
        })
      }),
  })
}
