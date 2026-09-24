import {
  historyEntryId,
  validatePresentationHistoryEntry,
  presentationHistoryBytes,
  type PresentationHistoryEntry,
  type PresentationHistoryEnvelope,
} from './presentation-change-history.js'
import {
  validatePresentationTextChange,
  type PresentationTextChange,
} from './presentation-text-change.js'
import {
  validatePresentationPageReplacement,
  type PresentationPageReplacement,
} from './presentation-page-replacement-record.js'
import {
  validatePresentationGeometryChange,
  type PresentationGeometryChange,
} from './presentation-geometry-change.js'
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
const PAGE_REPLACEMENT_KEY = 'wiswork.presentation.page-replacement.v1'
const HISTORY_KEY = 'wiswork.presentation.change-history.v1'
const TEXT_KEY = 'wiswork.presentation.text-change.v1'
const GEOMETRY_KEY = 'wiswork.presentation.geometry-change.v1'
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
  let geometryWriteFailed = false
  const readRawGeometryChange = (): PresentationGeometryChange | undefined => {
    const invalid = () => new Error('presentation_geometry_change_state_invalid')
    if (geometryWriteFailed) throw invalid()
    const raw = settings.get(GEOMETRY_KEY)
    // Empty string is the tombstone for a failed first save; the settings adapter has no delete.
    if (raw === undefined || raw === null || raw === '') return undefined
    if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 16 * 1024)
      throw invalid()
    let value: unknown
    try {
      value = JSON.parse(raw)
    } catch {
      throw invalid()
    }
    if (!validatePresentationGeometryChange(value)) throw invalid()
    return value
  }
  let textWriteFailed = false
  const readRawTextChange = (): PresentationTextChange | undefined => {
    const invalid = () => new Error('presentation_text_change_state_invalid')
    if (textWriteFailed) throw invalid()
    const raw = settings.get(TEXT_KEY)
    // Empty string is the tombstone for a failed first save; the settings adapter has no delete.
    if (raw === undefined || raw === null || raw === '') return undefined
    if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 192 * 1024)
      throw invalid()
    let value: unknown
    try {
      value = JSON.parse(raw)
    } catch {
      throw invalid()
    }
    if (!validatePresentationTextChange(value)) throw invalid()
    return value
  }
  let pageReplacementWriteFailed = false
  type Overrides = Record<string, PresentationImportRecord | null>
  const replacementKeys = (r: PresentationPageReplacement) => [
    `production/${r.projectId}/${r.parentRequestId}`,
    `production/${r.projectId}/${r.requestId}`,
  ]
  const replacementMappings = (
    r: PresentationPageReplacement,
    receipts: Overrides,
    undone = false,
    restoredId = r.restoredSlideId,
  ): Overrides => {
    const [parentKey, childKey] = replacementKeys(r)
    if (!undone) return { ...receipts, [parentKey]: null, [childKey]: r.childReceipt! }
    const parent = structuredClone(r.parentReceipt!)
    const index = parent.checkpoint!.pageIds!.indexOf(r.pageId)
    parent.slideIds![index] = restoredId!
    parent.checkpoint!.completed[index].slideId = restoredId!
    return { ...receipts, [parentKey]: parent, [childKey]: null }
  }
  const validOverrides = (receipts: unknown, document: string): receipts is Overrides => {
    if (
      !receipts ||
      typeof receipts !== 'object' ||
      Array.isArray(receipts) ||
      Object.keys(receipts).length > 32 ||
      new TextEncoder().encode(JSON.stringify(receipts)).byteLength > 100_000
    )
      return false
    return Object.entries(receipts).every(
      ([key, receipt]) =>
        key.startsWith('production/') &&
        validImportKey(key) &&
        (receipt === null ||
          (validPresentationImportRecord(receipt) &&
            receipt.state === 'complete' &&
            receipt.checkpoint?.version === 2 &&
            receipt.documentId === document)),
    )
  }
  const readReplacementEnvelope = (): {
    change?: PresentationPageReplacement
    receipts: Overrides
  } => {
    const invalid = () => new Error('presentation_page_replacement_state_invalid')
    if (pageReplacementWriteFailed) throw invalid()
    const raw = settings.get(PAGE_REPLACEMENT_KEY)
    if (raw === undefined || raw === null || raw === '') return { receipts: {} }
    if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 192 * 1024 + 100_128)
      throw invalid()
    let value
    try {
      value = JSON.parse(raw)
    } catch {
      throw invalid()
    }
    const envelope = value?.version === 2
    const change = envelope ? value.change : value
    if (!validatePresentationPageReplacement(change)) throw invalid()
    if (change.documentId !== JSON.stringify([settings.get(ID_KEY), settings.location()]))
      throw new Error('presentation_document_changed')
    const receipts = envelope ? value.receipts : {}
    if (
      (envelope &&
        Object.keys(value).some((k) => !['version', 'change', 'receipts'].includes(k))) ||
      !validOverrides(receipts, change.documentId)
    )
      throw invalid()
    if (['applied', 'undo_pending', 'restore_inserted', 'undone'].includes(change.state)) {
      const expected = replacementMappings(change, receipts, change.state === 'undone')
      for (const key of replacementKeys(change))
        if (JSON.stringify(receipts[key]) !== JSON.stringify(expected[key])) throw invalid()
    }
    if (change.state === 'commit_pending') {
      const [parentKey, childKey] = replacementKeys(change)
      const parent = Object.hasOwn(receipts, parentKey)
        ? receipts[parentKey]
        : readImports()[parentKey]
      if (
        JSON.stringify(parent) !== JSON.stringify(change.parentReceipt) ||
        Object.hasOwn(receipts, childKey) ||
        readImports()[childKey] !== undefined
      )
        throw invalid()
    }
    if (
      change.parentReceipt &&
      (!validOverrides(replacementMappings(change, receipts), change.documentId) ||
        !validOverrides(
          replacementMappings(change, receipts, true, '\uffff'.repeat(256)),
          change.documentId,
        ))
    )
      throw invalid()
    return { change, receipts }
  }
  const readRawPageReplacement = () => readReplacementEnvelope().change
  let historyWriteFailed = false
  const invalidHistory = () => new Error('presentation_change_history_state_invalid')
  const rawHeads = () => ({
    text: readRawTextChange(),
    geometry: readRawGeometryChange(),
    page: readRawPageReplacement(),
  })
  const readHistory = (): PresentationHistoryEnvelope => {
    if (historyWriteFailed) throw invalidHistory()
    const heads = rawHeads(),
      images = Object.values(readImageRecords()),
      raw = settings.get(HISTORY_KEY)
    if (raw === undefined || raw === null || raw === '') {
      const history: PresentationHistoryEnvelope = { version: 1, entries: [], heads: {} }
      for (const kind of ['text', 'geometry', 'page'] as const) {
        const record = heads[kind]
        if (!record) continue
        const id = historyEntryId(kind, record)
        history.entries.push({
          id,
          kind,
          record,
          legacy: true,
          sequence: history.entries.length + 1,
        } as PresentationHistoryEntry)
        history.heads[kind] = id
      }
      for (const record of images)
        history.entries.push({
          id: historyEntryId('image', record),
          kind: 'image',
          record,
          legacy: true,
          sequence: history.entries.length + 1,
        })
      return history
    }
    if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 1024 * 1024)
      throw invalidHistory()
    let h: PresentationHistoryEnvelope
    try {
      h = JSON.parse(raw)
    } catch {
      throw invalidHistory()
    }
    if (
      !h ||
      h.version !== 1 ||
      Object.keys(h).length !== 3 ||
      !Array.isArray(h.entries) ||
      h.entries.length > 64 ||
      !h.heads ||
      typeof h.heads !== 'object' ||
      Array.isArray(h.heads) ||
      Object.keys(h.heads).some((k) => !['text', 'geometry', 'page'].includes(k))
    )
      throw invalidHistory()
    const ids = new Set<string>()
    let sequence = 0
    for (const e of h.entries) {
      if (!validatePresentationHistoryEntry(e) || ids.has(e.id) || e.sequence <= sequence)
        throw invalidHistory()
      ids.add(e.id)
      sequence = e.sequence
    }
    if (presentationHistoryBytes(h) > 1024 * 1024) throw invalidHistory()
    for (const kind of ['text', 'geometry', 'page'] as const) {
      const head = h.entries.find((e) => e.id === h.heads[kind])
      if ((head && head.kind !== kind) || (!head && h.entries.some((e) => e.kind === kind)))
        throw invalidHistory()
      if (
        (h.heads[kind] !== undefined && !head) ||
        JSON.stringify(head?.record) !== JSON.stringify(heads[kind])
      )
        throw invalidHistory()
    }
    const historicalImages = h.entries.filter((e) => e.kind === 'image')
    if (
      historicalImages.length !== images.length ||
      images.some(
        (r) =>
          !historicalImages.some(
            (e) =>
              e.id === historyEntryId('image', r) && JSON.stringify(e.record) === JSON.stringify(r),
          ),
      )
    )
      throw invalidHistory()
    return h
  }
  const readTextChange = (changeId?: string): PresentationTextChange | undefined => {
    const raw = readRawTextChange()
    const history = readHistory()
    return changeId === undefined
      ? raw
      : history.entries.find(
          (e): e is Extract<PresentationHistoryEntry, { kind: 'text' }> =>
            e.kind === 'text' && e.record.changeId === changeId,
        )?.record
  }
  const readGeometryChange = (changeId?: string): PresentationGeometryChange | undefined => {
    const raw = readRawGeometryChange()
    const history = readHistory()
    return changeId === undefined
      ? raw
      : history.entries.find(
          (e): e is Extract<PresentationHistoryEntry, { kind: 'geometry' }> =>
            e.kind === 'geometry' && e.record.changeId === changeId,
        )?.record
  }
  const readPageReplacement = () => {
    const raw = readRawPageReplacement()
    readHistory()
    return raw
  }
  const saveWithHistory = async (
    key: string,
    serialized: string,
    entry: PresentationHistoryEntry,
    fail: () => void,
  ) => {
    const history = readHistory(),
      index = history.entries.findIndex((e) => e.id === entry.id)
    const unresolved = (e: PresentationHistoryEntry) =>
      !['applied', 'undone', 'discarded', 'complete'].includes(e.record.state)
    if (
      (index < 0 || !unresolved(history.entries[index])) &&
      history.entries.some((e) => e.id !== entry.id && unresolved(e))
    )
      throw new Error('presentation_change_history_pending')
    if (index < 0)
      history.entries.push({ ...entry, sequence: (history.entries.at(-1)?.sequence ?? 0) + 1 })
    else
      history.entries[index] = {
        ...entry,
        sequence: history.entries[index].sequence,
        legacy: history.entries[index].legacy,
      } as PresentationHistoryEntry
    if (entry.kind !== 'image') history.heads[entry.kind] = entry.id
    if (
      history.entries.length > 64 ||
      !Number.isSafeInteger(history.entries.at(-1)?.sequence) ||
      presentationHistoryBytes(history) > 1024 * 1024
    )
      throw new Error('presentation_change_history_full')
    const historySerialized = JSON.stringify(history),
      previous = settings.get(key),
      previousHistory = settings.get(HISTORY_KEY),
      location = settings.location(),
      identity = settings.get(ID_KEY)
    try {
      settings.set(key, serialized)
      settings.set(HISTORY_KEY, historySerialized)
      await settings.save()
      if (
        settings.location() !== location ||
        settings.get(ID_KEY) !== identity ||
        settings.get(key) !== serialized ||
        settings.get(HISTORY_KEY) !== historySerialized
      )
        throw new Error('presentation_document_changed')
    } catch (error) {
      if (
        settings.location() === location &&
        settings.get(ID_KEY) === identity &&
        [previous, serialized].includes(settings.get(key)) &&
        [previousHistory, historySerialized].includes(settings.get(HISTORY_KEY))
      ) {
        try {
          settings.set(key, typeof previous === 'string' ? previous : key === IMAGE_KEY ? '{}' : '')
          settings.set(HISTORY_KEY, typeof previousHistory === 'string' ? previousHistory : '')
          if (
            settings.get(key) !==
              (typeof previous === 'string' ? previous : key === IMAGE_KEY ? '{}' : '') ||
            settings.get(HISTORY_KEY) !==
              (typeof previousHistory === 'string' ? previousHistory : '')
          )
            throw invalidHistory()
        } catch {
          historyWriteFailed = true
          fail()
        }
      } else {
        historyWriteFailed = true
        fail()
      }
      throw error
    }
  }
  const readReceipt = (key: string): PresentationImportRecord | undefined => {
    const { receipts } = readReplacementEnvelope()
    if (Object.hasOwn(receipts, key)) {
      if (receipts[key] === null) throw new Error('presentation_import_superseded')
      return receipts[key]
    }
    return readImports()[key]
  }
  return {
    documentId,
    listChangeHistory: () => structuredClone(readHistory().entries),
    readPageReplacement,
    writePageReplacement(
      record: PresentationPageReplacement,
      expectedChange: PresentationPageReplacement | undefined,
    ) {
      const snapshot = structuredClone(record),
        expected = structuredClone(expectedChange)
      const write = async () => {
        const invalid = () => new Error('presentation_page_replacement_state_invalid')
        if (
          !validatePresentationPageReplacement(snapshot) ||
          (expected !== undefined && !validatePresentationPageReplacement(expected))
        )
          throw invalid()
        if ((await documentId()) !== snapshot.documentId)
          throw new Error('presentation_document_changed')
        const prior = readPageReplacement()
        if (
          prior?.changeId !== snapshot.changeId &&
          readHistory().entries.some((e) => e.id === historyEntryId('page', snapshot))
        )
          throw invalid()
        if (JSON.stringify(prior) !== JSON.stringify(expected))
          throw new Error('presentation_page_replacement_stale')
        if (JSON.stringify(prior) === JSON.stringify(snapshot)) return
        if (prior?.changeId === snapshot.changeId) {
          const identity = (r: PresentationPageReplacement) =>
            JSON.stringify({
              ...r,
              state: undefined,
              newSlideId: undefined,
              parentReceipt: undefined,
              childReceipt: undefined,
              restoredSlideId: undefined,
            })
          const transitions: Record<PresentationPageReplacement['state'], string[]> = {
            pending: ['inserted'],
            inserted: ['staged'],
            staged: ['discard_pending', 'commit_pending'],
            discard_pending: ['discarded'],
            discarded: [],
            commit_pending: ['applied'],
            applied: ['undo_pending'],
            undo_pending: ['restore_inserted'],
            restore_inserted: ['undone'],
            undone: [],
          }
          if (
            identity(prior) !== identity(snapshot) ||
            !transitions[prior.state].includes(snapshot.state) ||
            (prior.newSlideId !== undefined && prior.newSlideId !== snapshot.newSlideId) ||
            (prior.parentReceipt !== undefined &&
              JSON.stringify(prior.parentReceipt) !== JSON.stringify(snapshot.parentReceipt)) ||
            (prior.childReceipt !== undefined &&
              JSON.stringify(prior.childReceipt) !== JSON.stringify(snapshot.childReceipt)) ||
            (prior.restoredSlideId !== undefined &&
              prior.restoredSlideId !== snapshot.restoredSlideId)
          )
            throw invalid()
        } else if (
          snapshot.state !== 'pending' ||
          (prior && !['discarded', 'undone'].includes(prior.state))
        )
          throw invalid()
        let { receipts } = readReplacementEnvelope()
        if (snapshot.state === 'commit_pending') {
          const [parentKey, childKey] = replacementKeys(snapshot)
          if (
            JSON.stringify(readReceipt(parentKey)) !== JSON.stringify(snapshot.parentReceipt) ||
            Object.hasOwn(receipts, childKey) ||
            readImports()[childKey] !== undefined
          )
            throw new Error('presentation_page_replacement_stale')
          // Reserve both terminal mappings before any destructive host write, including a maximal restored ID.
          if (
            !validOverrides(replacementMappings(snapshot, receipts), snapshot.documentId) ||
            !validOverrides(
              replacementMappings(snapshot, receipts, true, '\uffff'.repeat(256)),
              snapshot.documentId,
            )
          )
            throw new Error('presentation_import_history_full')
        }
        if (snapshot.state === 'applied' || snapshot.state === 'undone')
          receipts = replacementMappings(snapshot, receipts, snapshot.state === 'undone')
        if (!validOverrides(receipts, snapshot.documentId))
          throw new Error('presentation_import_history_full')
        await saveWithHistory(
          PAGE_REPLACEMENT_KEY,
          JSON.stringify({ version: 2, change: snapshot, receipts }),
          {
            id: historyEntryId('page', snapshot),
            kind: 'page',
            record: snapshot,
            legacy: false,
            sequence: 1,
          },
          () => {
            pageReplacementWriteFailed = true
          },
        )
      }
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
    readGeometryChange,
    writeGeometryChange(
      record: PresentationGeometryChange,
      expectedChange: PresentationGeometryChange | undefined,
    ) {
      const snapshot = structuredClone(record),
        expected = structuredClone(expectedChange)
      const write = async () => {
        const invalid = () => new Error('presentation_geometry_change_state_invalid')
        if (
          !validatePresentationGeometryChange(snapshot) ||
          (expected !== undefined && !validatePresentationGeometryChange(expected))
        )
          throw invalid()
        if ((await documentId()) !== snapshot.documentId)
          throw new Error('presentation_document_changed')
        const prior = readGeometryChange(
          expected?.changeId === snapshot.changeId ||
            readHistory().entries.some((e) => e.id === historyEntryId('geometry', snapshot))
            ? snapshot.changeId
            : undefined,
        )
        if (JSON.stringify(prior) !== JSON.stringify(expected))
          throw new Error('presentation_geometry_change_stale')
        if (JSON.stringify(prior) === JSON.stringify(snapshot)) return
        if (prior?.changeId === snapshot.changeId) {
          const identity = (r: PresentationGeometryChange) =>
            JSON.stringify({ ...r, state: undefined })
          const transitions = {
            pending: 'applied',
            applied: 'undo_pending',
            undo_pending: 'undone',
            undone: undefined,
          }
          if (identity(prior) !== identity(snapshot) || transitions[prior.state] !== snapshot.state)
            throw invalid()
        } else if (
          snapshot.state !== 'pending' ||
          (prior && !['applied', 'undone'].includes(prior.state))
        )
          throw invalid()
        await saveWithHistory(
          GEOMETRY_KEY,
          JSON.stringify(snapshot),
          {
            id: historyEntryId('geometry', snapshot),
            kind: 'geometry',
            record: snapshot,
            legacy: false,
            sequence: 1,
          },
          () => {
            geometryWriteFailed = true
          },
        )
      }
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
    readTextChange,
    writeTextChange(
      record: PresentationTextChange,
      expectedChange: PresentationTextChange | undefined,
    ) {
      const snapshot = structuredClone(record),
        expected = structuredClone(expectedChange)
      const write = async () => {
        const invalid = () => new Error('presentation_text_change_state_invalid')
        if (
          !validatePresentationTextChange(snapshot) ||
          (expected !== undefined && !validatePresentationTextChange(expected))
        )
          throw invalid()
        if ((await documentId()) !== snapshot.documentId)
          throw new Error('presentation_document_changed')
        const prior = readTextChange(
          expected?.changeId === snapshot.changeId ||
            readHistory().entries.some((e) => e.id === historyEntryId('text', snapshot))
            ? snapshot.changeId
            : undefined,
        )
        if (JSON.stringify(prior) !== JSON.stringify(expected))
          throw new Error('presentation_text_change_stale')
        if (JSON.stringify(prior) === JSON.stringify(snapshot)) return
        if (prior?.changeId === snapshot.changeId) {
          const identity = (r: PresentationTextChange) => JSON.stringify({ ...r, state: undefined })
          const transitions = {
            pending: 'applied',
            applied: 'undo_pending',
            undo_pending: 'undone',
            undone: undefined,
          }
          if (identity(prior) !== identity(snapshot) || transitions[prior.state] !== snapshot.state)
            throw invalid()
        } else if (
          snapshot.state !== 'pending' ||
          (prior && !['applied', 'undone'].includes(prior.state))
        )
          throw invalid()
        await saveWithHistory(
          TEXT_KEY,
          JSON.stringify(snapshot),
          {
            id: historyEntryId('text', snapshot),
            kind: 'text',
            record: snapshot,
            legacy: false,
            sequence: 1,
          },
          () => {
            textWriteFailed = true
          },
        )
      }
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
    readReceipt,
    writeReceipt(key: string, record: PresentationImportRecord | undefined) {
      const write = async () => {
        if (
          !validImportKey(key, record) ||
          (record !== undefined && !validPresentationImportRecord(record))
        )
          throw new Error('presentation_import_state_invalid')
        const { change, receipts } = readReplacementEnvelope()
        if (change?.state === 'commit_pending' && replacementKeys(change).includes(key)) {
          if (
            key === replacementKeys(change)[0] &&
            JSON.stringify(record) === JSON.stringify(change.parentReceipt)
          )
            return
          throw new Error('presentation_import_superseded')
        }
        if (Object.hasOwn(receipts, key)) {
          if (receipts[key] !== null && JSON.stringify(receipts[key]) === JSON.stringify(record))
            return
          throw new Error('presentation_import_superseded')
        }
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
    listImageReplacements: () => {
      const records = readImageRecords()
      readHistory()
      return Object.values(records).map((record) => structuredClone(record))
    },
    readImageReplacement: (key: string) => {
      const records = readImageRecords()
      readHistory()
      return records[key]
    },
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
              r.backup,
            ])
          if (
            identity(prior) !== identity(snapshot) ||
            (prior.newShapeId !== undefined && prior.newShapeId !== snapshot.newShapeId) ||
            (snapshot.state === 'complete' && !prior.newShapeId) ||
            (snapshot.state === 'undone' && !prior.restoredShapeId) ||
            (prior.after !== undefined &&
              JSON.stringify(prior.after) !== JSON.stringify(snapshot.after)) ||
            (prior.undoBaseline !== undefined &&
              JSON.stringify(prior.undoBaseline) !== JSON.stringify(snapshot.undoBaseline)) ||
            (prior.restoredShapeId !== undefined &&
              prior.restoredShapeId !== snapshot.restoredShapeId) ||
            !{
              pending: ['pending', 'complete'],
              complete: ['complete', 'undo_pending'],
              undo_pending: ['undo_pending', 'undone'],
              undone: ['undone'],
            }[prior.state].includes(snapshot.state)
          )
            throw invalid()
          if (JSON.stringify(prior) === JSON.stringify(snapshot)) return
        }
        records[key] = snapshot
        const serialized = JSON.stringify(records)
        if (Object.keys(records).length > 32 || imageBytes(serialized, records) > 128 * 1024)
          throw new Error('presentation_image_replacement_history_full')
        await saveWithHistory(
          IMAGE_KEY,
          serialized,
          {
            id: historyEntryId('image', snapshot),
            kind: 'image',
            record: snapshot,
            legacy: false,
            sequence: 1,
          },
          () => {
            imageWriteFailed = true
          },
        )
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
