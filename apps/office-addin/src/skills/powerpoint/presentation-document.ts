import { validPresentationImportRecord } from './presentation-page-delivery.js'
import type { PresentationImportRecord } from './presentation-delivery.js'
const ID_KEY = 'wiswork.presentation.document.v1'
const IMPORT_KEY = 'wiswork.presentation.imports.v1'
const PROJECT_KEY = 'wiswork.presentation.project.v1'
const validId = (value: unknown): value is string =>
  typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)

interface DocumentSettings {
  get(key: string): unknown
  set(key: string, value: string): void
  save(): Promise<void>
  location(): string
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
      if (
        !/^[A-Za-z0-9_-]{1,128}\/[A-Za-z0-9_-]{1,128}$/.test(key) ||
        !validPresentationImportRecord(record)
      )
        throw new Error('presentation_import_state_invalid')
    }
    return value
  }
  return {
    documentId,
    readReceipt: (key: string) => readImports()[key],
    writeReceipt(key: string, record: PresentationImportRecord | undefined) {
      const write = async () => {
        if (
          !/^[A-Za-z0-9_-]{1,128}\/[A-Za-z0-9_-]{1,128}$/.test(key) ||
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
