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
  const readImports = (): Record<string, PresentationImportRecord> => {
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
      const r = record as PresentationImportRecord
      if (
        !/^[A-Za-z0-9_-]{1,128}\/[A-Za-z0-9_-]{1,128}$/.test(key) ||
        !r ||
        !['pending', 'complete'].includes(r.state) ||
        typeof r.documentId !== 'string' ||
        (r.state === 'complete' &&
          (!Array.isArray(r.slideIds) ||
            r.slideIds.length > 100 ||
            !r.slideIds.every((id) => typeof id === 'string' && id.length <= 256)))
      )
        throw new Error('presentation_import_state_invalid')
    }
    return value
  }
  return {
    documentId,
    readReceipt: (key: string) => readImports()[key],
    async writeReceipt(key: string, record: PresentationImportRecord | undefined) {
      if (!/^[A-Za-z0-9_-]{1,128}\/[A-Za-z0-9_-]{1,128}$/.test(key))
        throw new Error('presentation_import_state_invalid')
      const imports = readImports()
      if (record) imports[key] = record
      else delete imports[key]
      const serialized = JSON.stringify(imports)
      if (Object.keys(imports).length > 32 || serialized.length > 100_000)
        throw new Error('presentation_import_history_full')
      settings.set(IMPORT_KEY, serialized)
      await settings.save()
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
