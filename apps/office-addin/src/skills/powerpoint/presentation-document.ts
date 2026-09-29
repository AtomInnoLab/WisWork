import {
  validatePresentationPackageChange,
  validPackageTransition,
  type PresentationPackageChange,
} from './presentation-package-change.js'
import {
  validatePresentationNativeMasterChange,
  validNativeMasterTransition,
  type PresentationNativeMasterChange,
} from './presentation-native-master-change.js'
import {
  parsePresentationQaAttempt,
  presentationQaAttemptIdentity,
  PRESENTATION_QA_ATTEMPT_TERMINAL_RESERVE_BYTES,
  type PresentationQaAttempt,
} from './presentation-qa-attempts.js'
import { parseAgentResumeMessages, type AgentMessage } from '@wiswork/agent-core'
import {
  validatePresentationExistingChange,
  type PresentationExistingChange,
} from './presentation-existing-change.js'
import {
  validatePresentationExistingImageChange,
  validExistingImageTransition,
  type PresentationExistingImageChange,
} from './presentation-existing-image.js'
import {
  validatePresentationExistingPageChange,
  validExistingPageTransition,
  type PresentationExistingPageChange,
} from './presentation-existing-page.js'
import {
  validatePresentationExistingChartChange,
  validExistingChartTransition,
  type PresentationExistingChartChange,
} from './presentation-existing-chart.js'
import {
  validatePresentationExistingBatch,
  validExistingBatchTransition,
  type PresentationExistingBatch,
} from './presentation-existing-batch.js'
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
  PRESENTATION_QA_INVALIDATED_FIELD_BYTES,
  type PresentationQaRecord,
} from './presentation-qa.js'
import { validPresentationImportRecord } from './presentation-page-delivery.js'
import type { PresentationImportRecord } from './presentation-delivery.js'
const ID_KEY = 'wiswork.presentation.document.v1'
const IMPORT_KEY = 'wiswork.presentation.imports.v1'
const IMAGE_KEY = 'wiswork.presentation.image-replacements.v1'
const PAGE_REPLACEMENT_KEY = 'wiswork.presentation.page-replacement.v1'
const EXISTING_KEY = 'wiswork.presentation.existing-change.v1'
const EXISTING_BATCH_KEY = 'wiswork.presentation.existing-batch.v1'
const EXISTING_IMAGE_KEY = 'wiswork.presentation.existing-image.v1'
const EXISTING_PAGE_KEY = 'wiswork.presentation.existing-page.v1'
const EXISTING_CHART_KEY = 'wiswork.presentation.existing-chart.v1'
const PACKAGE_XML_KEY = 'wiswork.presentation.package-xml.v1'
const NATIVE_MASTER_KEY = 'wiswork.presentation.native-master.v1'
const HISTORY_KEY = 'wiswork.presentation.change-history.v1'
const TEXT_KEY = 'wiswork.presentation.text-change.v1'
const GEOMETRY_KEY = 'wiswork.presentation.geometry-change.v1'
const QA_KEY = 'wiswork.presentation.qa.v1'
const QA_ATTEMPTS_KEY = 'wiswork.presentation.qa-attempts.v1'
const PROJECT_KEY = 'wiswork.presentation.project.v1'
const SELECTED_PRODUCTION_KEY = 'wiswork.presentation.selected-production.v1'
const AGENT_RUN_KEY = 'wiswork.presentation.agent-run.v1'
export interface PresentationAgentRunRecovery {
  runId: string
  instruction: string
  phase: 'running' | 'tool_pending' | 'tool_completed'
  toolName?: string
  toolCallId?: string
  restartSafe?: boolean
  messages?: AgentMessage[]
  importReceipt?: { state: 'complete' | 'partial' | 'uncertain'; completed: number; total?: number }
  changeReceipt?: { total: number; unresolved: number }
}
// Only audited reads may restart the original instruction after a Taskpane interruption.
// A write or an unknown tool permanently closes this path for the run.
const restartSafeTools = new Set([
  'read_presentation_plan',
  'read_presentation_import_status',
  'read_presentation_page',
  'read_presentation_page_geometry',
  'read_presentation_baseline',
  'read_presentation_baseline_page',
  'read_presentation_baseline_complex_page',
  'read_presentation_baseline_chart_source',
  'read_presentation_baseline_notes',
  'read_presentation_baseline_source_links',
  'read_presentation_baseline_rich_text',
  'read_presentation_qa',
  'read_presentation_attachment',
  'list_presentation_attachments',
  'list_presentation_changes',
  'list_presentation_brand_kits',
  'read_presentation_brand_kit',
  'list_presentation_preferences',
  'read_presentation_domain_skill',
  'compare_presentation_page_structure',
  'read_presentation_page_reviews',
  'read_presentation_claim_review',
  'read_presentation_claim_evidence',
  'check_presentation_page_content',
  'read_presentation_production',
  'audit_presentation_sources',
  'list_presentation_review_comments',
])
const AGENT_RUN_LOCAL_PREFIX = 'wiswork.presentation.agent-run.prompt.v1.'
const AGENT_RUN_LOCAL_TTL_MS = 7 * 24 * 60 * 60 * 1000
const AGENT_RUN_LOCAL_RECORD_LIMIT = 144 * 1024
const AGENT_RUN_INSTRUCTION_LIMIT = 8 * 1024
function recoveryMessages(
  value: unknown,
  record: PresentationAgentRunRecovery,
): AgentMessage[] | undefined {
  if (record.phase !== 'tool_completed' || record.restartSafe !== true || !record.toolCallId)
    return undefined
  const messages = parseAgentResumeMessages(value)
  const user = messages
    ? [...messages].reverse().find((message) => message.role === 'user')
    : undefined
  const currentMessages = messages && user ? messages.slice(messages.indexOf(user)) : undefined
  if (
    !messages ||
    currentMessages?.some(
      (message) =>
        message.role === 'assistant' &&
        message.toolCalls?.some((call) => !restartSafeTools.has(call.name)),
    )
  )
    return undefined
  if (
    !record.instruction ||
    user?.role !== 'user' ||
    (user.text !== record.instruction && !user.text.startsWith(`${record.instruction}\n\n`))
  )
    return undefined
  const last = messages.at(-1)
  const result = last?.role === 'tool' ? last.results.at(-1) : undefined
  return result?.id === record.toolCallId && result.name === record.toolName ? messages : undefined
}
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
  let runCheckpointQueue: Promise<void> = Promise.resolve()
  const queueRunCheckpoint = (write: () => Promise<void>) => {
    const result = runCheckpointQueue.then(write)
    runCheckpointQueue = result.catch(() => undefined)
    return result
  }
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
  let qaAttemptWriteFailed = false
  const qaAttemptKey = (a: PresentationQaAttempt) =>
    `${a.source === 'production' ? 'production/' : ''}${a.projectId}/${a.requestId}`
  const qaAttemptValue = (a: PresentationQaAttempt) =>
    JSON.stringify([presentationQaAttemptIdentity(a), a.status, a.finishedAt, a.errorCode])
  const qaAttemptBytes = (raw: string, records: Record<string, PresentationQaAttempt>) =>
    new TextEncoder().encode(raw).byteLength +
    Object.values(records).filter((a) => a.status === 'started').length *
      PRESENTATION_QA_ATTEMPT_TERMINAL_RESERVE_BYTES
  const readRawQaAttempts = (): Record<string, PresentationQaAttempt> => {
    if (qaAttemptWriteFailed) throw Error('presentation_qa_attempt_state_invalid')
    const raw = settings.get(QA_ATTEMPTS_KEY)
    if (raw === undefined || raw === null) return {}
    try {
      if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 128 * 1024)
        throw Error()
      const map = JSON.parse(raw)
      if (!map || typeof map !== 'object' || Array.isArray(map) || Object.keys(map).length > 64)
        throw Error()
      const result: Record<string, PresentationQaAttempt> = {}
      for (const [id, value] of Object.entries(map)) {
        const attempt = parsePresentationQaAttempt(value)
        if (id !== attempt.id) throw Error()
        result[id] = attempt
      }
      if (qaAttemptBytes(raw, result) > 128 * 1024) throw Error()
      return result
    } catch {
      throw Error('presentation_qa_attempt_state_invalid')
    }
  }
  let qaWriteFailed = false
  const readQaRecords = (): Record<string, PresentationQaRecord> => {
    if (qaWriteFailed) throw new Error('presentation_qa_state_invalid')
    const raw = settings.get(QA_KEY)
    if (raw === undefined || raw === null) return {}
    if (
      typeof raw !== 'string' ||
      new TextEncoder().encode(raw).byteLength >
        256 * 1024 +
          8 * 32 * (PRESENTATION_QA_RECHECK_FIELD_BYTES + PRESENTATION_QA_INVALIDATED_FIELD_BYTES)
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
  let existingWriteFailed = false
  let existingBatchWriteFailed = false
  let existingImageWriteFailed = false
  const readRawExistingChange = (): PresentationExistingChange | undefined => {
    const invalid = () => new Error('presentation_existing_change_state_invalid')
    if (existingWriteFailed) throw invalid()
    const raw = settings.get(EXISTING_KEY)
    if (raw === undefined || raw === null || raw === '') return undefined
    if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 192 * 1024)
      throw invalid()
    let value: unknown
    try {
      value = JSON.parse(raw)
    } catch {
      throw invalid()
    }
    if (!validatePresentationExistingChange(value)) throw invalid()
    return value
  }
  const readRawExistingBatch = (): PresentationExistingBatch | undefined => {
    if (existingBatchWriteFailed) throw new Error('presentation_existing_batch_state_invalid')
    const raw = settings.get(EXISTING_BATCH_KEY)
    if (raw === undefined || raw === null || raw === '') return undefined
    if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 192 * 1024)
      throw new Error('presentation_existing_batch_state_invalid')
    let value: unknown
    try {
      value = JSON.parse(raw)
    } catch {
      throw new Error('presentation_existing_batch_state_invalid')
    }
    if (!validatePresentationExistingBatch(value))
      throw new Error('presentation_existing_batch_state_invalid')
    return value
  }
  const readRawExistingImage = (): PresentationExistingImageChange | undefined => {
    if (existingImageWriteFailed) throw new Error('presentation_existing_image_state_invalid')
    const raw = settings.get(EXISTING_IMAGE_KEY)
    if (raw === undefined || raw === null || raw === '') return undefined
    if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 192 * 1024)
      throw new Error('presentation_existing_image_state_invalid')
    let value: unknown
    try {
      value = JSON.parse(raw)
    } catch {
      throw new Error('presentation_existing_image_state_invalid')
    }
    if (!validatePresentationExistingImageChange(value))
      throw new Error('presentation_existing_image_state_invalid')
    return value
  }
  let existingPageWriteFailed = false
  const readRawExistingPage = (): PresentationExistingPageChange | undefined => {
    if (existingPageWriteFailed) throw new Error('presentation_existing_page_state_invalid')
    const raw = settings.get(EXISTING_PAGE_KEY)
    if (raw === undefined || raw === null || raw === '') return undefined
    if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 192 * 1024)
      throw new Error('presentation_existing_page_state_invalid')
    let value: unknown
    try {
      value = JSON.parse(raw)
    } catch {
      throw new Error('presentation_existing_page_state_invalid')
    }
    if (!validatePresentationExistingPageChange(value))
      throw new Error('presentation_existing_page_state_invalid')
    return value
  }
  let existingChartWriteFailed = false
  const readRawExistingChart = (): PresentationExistingChartChange | undefined => {
    if (existingChartWriteFailed) throw new Error('presentation_existing_chart_state_invalid')
    const raw = settings.get(EXISTING_CHART_KEY)
    if (raw === undefined || raw === null || raw === '') return undefined
    if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 32 * 1024)
      throw new Error('presentation_existing_chart_state_invalid')
    let value: unknown
    try {
      value = JSON.parse(raw)
    } catch {
      throw new Error('presentation_existing_chart_state_invalid')
    }
    if (!validatePresentationExistingChartChange(value))
      throw new Error('presentation_existing_chart_state_invalid')
    return value
  }
  let packageWriteFailed = false
  const readRawPackageChange = (): PresentationPackageChange | undefined => {
    if (packageWriteFailed) throw new Error('presentation_package_state_invalid')
    const raw = settings.get(PACKAGE_XML_KEY)
    if (raw === undefined || raw === null || raw === '') return undefined
    if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 192 * 1024)
      throw new Error('presentation_package_state_invalid')
    let value: unknown
    try {
      value = JSON.parse(raw)
    } catch {
      throw new Error('presentation_package_state_invalid')
    }
    if (!validatePresentationPackageChange(value))
      throw new Error('presentation_package_state_invalid')
    return value
  }
  let nativeMasterWriteFailed = false
  const readRawNativeMasterChange = (): PresentationNativeMasterChange | undefined => {
    if (nativeMasterWriteFailed) throw new Error('presentation_native_master_state_invalid')
    const raw = settings.get(NATIVE_MASTER_KEY)
    if (raw === undefined || raw === null || raw === '') return undefined
    if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 192 * 1024)
      throw new Error('presentation_native_master_state_invalid')
    let value: unknown
    try {
      value = JSON.parse(raw)
    } catch {
      throw new Error('presentation_native_master_state_invalid')
    }
    if (!validatePresentationNativeMasterChange(value))
      throw new Error('presentation_native_master_state_invalid')
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
    existing: readRawExistingChange(),
    existing_batch: readRawExistingBatch(),
    existing_image: readRawExistingImage(),
    existing_page: readRawExistingPage(),
    existing_chart: readRawExistingChart(),
    native_master: readRawNativeMasterChange(),
    package_xml: readRawPackageChange(),
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
      for (const kind of [
        'text',
        'geometry',
        'page',
        'existing',
        'existing_batch',
        'existing_image',
        'existing_page',
        'existing_chart',
        'native_master',
        'package_xml',
      ] as const) {
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
    if (
      typeof raw !== 'string' ||
      new TextEncoder().encode(raw).byteLength > 1024 * 1024 + 64 * 120
    )
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
      Object.keys(h.heads).some(
        (k) =>
          ![
            'text',
            'geometry',
            'page',
            'existing',
            'existing_batch',
            'existing_image',
            'existing_page',
            'existing_chart',
            'native_master',
            'package_xml',
          ].includes(k),
      )
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
    for (const kind of [
      'text',
      'geometry',
      'page',
      'existing',
      'existing_batch',
      'existing_image',
      'existing_page',
      'existing_chart',
      'native_master',
      'package_xml',
    ] as const) {
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
  const readExistingChange = (changeId: string): PresentationExistingChange | undefined =>
    readHistory().entries.find(
      (e): e is Extract<PresentationHistoryEntry, { kind: 'existing' }> =>
        e.kind === 'existing' && e.record.changeId === changeId,
    )?.record
  const readExistingBatch = (changeId: string): PresentationExistingBatch | undefined =>
    readHistory().entries.find(
      (e): e is Extract<PresentationHistoryEntry, { kind: 'existing_batch' }> =>
        e.kind === 'existing_batch' && e.record.changeId === changeId,
    )?.record
  const readExistingImageChange = (changeId: string): PresentationExistingImageChange | undefined =>
    readHistory().entries.find(
      (e): e is Extract<PresentationHistoryEntry, { kind: 'existing_image' }> =>
        e.kind === 'existing_image' && e.record.changeId === changeId,
    )?.record
  const readExistingPageChange = (changeId: string): PresentationExistingPageChange | undefined =>
    readHistory().entries.find(
      (e): e is Extract<PresentationHistoryEntry, { kind: 'existing_page' }> =>
        e.kind === 'existing_page' && e.record.changeId === changeId,
    )?.record
  const readExistingChartChange = (changeId: string): PresentationExistingChartChange | undefined =>
    readHistory().entries.find(
      (e): e is Extract<PresentationHistoryEntry, { kind: 'existing_chart' }> =>
        e.kind === 'existing_chart' && e.record.changeId === changeId,
    )?.record
  const readNativeMasterChange = (changeId: string): PresentationNativeMasterChange | undefined =>
    readHistory().entries.find(
      (entry): entry is Extract<PresentationHistoryEntry, { kind: 'native_master' }> =>
        entry.kind === 'native_master' && entry.record.changeId === changeId,
    )?.record
  const readPackageChange = (changeId: string): PresentationPackageChange | undefined =>
    readHistory().entries.find(
      (entry): entry is Extract<PresentationHistoryEntry, { kind: 'package_xml' }> =>
        entry.kind === 'package_xml' && entry.record.changeId === changeId,
    )?.record
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
    const previousEntry = history.entries[index]
    const latestAt = history.entries.reduce((latest, value) => {
      const at = value.checkpointRestoredAt ?? value.checkpointCreatedAt ?? ''
      return at > latest ? at : latest
    }, '')
    const now = new Date().toISOString()
    const at = now < latestAt ? latestAt : now
    const events = {
      ...(previousEntry?.checkpointCreatedAt
        ? { checkpointCreatedAt: previousEntry.checkpointCreatedAt }
        : index < 0 && !entry.legacy
          ? { checkpointCreatedAt: at }
          : {}),
      ...(previousEntry?.checkpointRestoredAt
        ? { checkpointRestoredAt: previousEntry.checkpointRestoredAt }
        : (entry.record.state === 'undone' ||
              (entry.kind === 'package_xml' && entry.record.state === 'discarded')) &&
            previousEntry &&
            previousEntry.record.state !== entry.record.state
          ? { checkpointRestoredAt: at }
          : {}),
    }
    const unresolved = (e: PresentationHistoryEntry) =>
      (e.kind === 'package_xml' && Boolean(e.record.pending)) ||
      !['applied', 'undone', 'discarded', 'complete', 'cancelled'].includes(e.record.state)
    // A recovery may journal its own replacement while the source write remains uncertain.
    // The exception is tied to the exact original backup; unrelated pending work still blocks.
    const restoringSource = (source: PresentationHistoryEntry) => {
      if (
        entry.kind !== 'existing_page' ||
        source.kind !== 'existing_batch' ||
        source.record.version === 1 ||
        source.record.version === 4 ||
        !entry.record.restores ||
        entry.record.restores.sourceKind !== 'batch' ||
        source.record.changeId !== entry.record.restores.sourceChangeId ||
        source.record.documentId !== entry.record.documentId ||
        source.record.backupReleasedAt
      )
        return false
      const original = source.record.backups.find(
        (backup) => backup.hostSlideId === entry.record.oldSlideId,
      )
      return Boolean(
        original &&
        entry.record.restores.sourceHostSlideId === original.hostSlideId &&
        entry.record.restores.originalBackupId === original.backupId &&
        entry.record.restores.originalPackageDigest === original.packageDigest &&
        entry.record.replacementPackageDigest === original.packageDigest &&
        entry.record.sourceBackup?.sha256 === original.sha256 &&
        entry.record.sourceBackup.sizeBytes === original.sizeBytes,
      )
    }
    if (
      (index < 0 || !unresolved(history.entries[index])) &&
      history.entries.some((e) => e.id !== entry.id && unresolved(e) && !restoringSource(e))
    )
      throw new Error('presentation_change_history_pending')
    if (index < 0) {
      let attribution: { agentRunId: string; toolCallId: string } | undefined
      try {
        const raw = settings.get(AGENT_RUN_KEY)
        const run =
          typeof raw === 'string' ? (JSON.parse(raw) as Record<string, unknown>) : undefined
        if (
          run?.documentId === entry.record.documentId &&
          run.phase === 'tool_pending' &&
          typeof run.toolName === 'string' &&
          !restartSafeTools.has(run.toolName) &&
          !['import_generated_presentation', 'import_presentation_production'].includes(
            run.toolName,
          ) &&
          validId(run.runId) &&
          typeof run.toolCallId === 'string' &&
          run.toolCallId.length > 0 &&
          run.toolCallId.length <= 256
        )
          attribution = { agentRunId: run.runId, toolCallId: run.toolCallId }
      } catch {
        /* Legacy checkpoints cannot establish an originating tool. */
      }
      history.entries.push({
        ...entry,
        ...attribution,
        ...events,
        sequence: (history.entries.at(-1)?.sequence ?? 0) + 1,
      })
    } else
      history.entries[index] = {
        ...entry,
        ...events,
        sequence: history.entries[index].sequence,
        legacy: history.entries[index].legacy,
        ...(history.entries[index].agentRunId
          ? {
              agentRunId: history.entries[index].agentRunId,
              toolCallId: history.entries[index].toolCallId,
            }
          : {}),
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
    assertDocumentId(expected: string): void {
      const id = settings.get(ID_KEY)
      if (
        unsavedId !== undefined ||
        !validId(id) ||
        JSON.stringify([id, settings.location()]) !== expected
      )
        throw new Error('presentation_document_changed')
    },
    listChangeHistory: () => structuredClone(readHistory().entries),
    readExistingChange,
    readExistingBatch,
    readExistingImageChange,
    readExistingPageChange,
    readExistingChartChange,
    readPackageChange,
    writePackageChange(
      record: PresentationPackageChange,
      expectedRecord: PresentationPackageChange | undefined,
    ) {
      const snapshot = structuredClone(record),
        expected = structuredClone(expectedRecord)
      const write = async () => {
        if (
          !validatePresentationPackageChange(snapshot) ||
          (expected !== undefined && !validatePresentationPackageChange(expected))
        )
          throw new Error('presentation_package_state_invalid')
        if ((await documentId()) !== snapshot.documentId)
          throw new Error('presentation_document_changed')
        const prior = readPackageChange(snapshot.changeId)
        if (JSON.stringify(prior) !== JSON.stringify(expected))
          throw new Error('presentation_package_stale')
        if (JSON.stringify(prior) === JSON.stringify(snapshot)) return
        if (!validPackageTransition(prior, snapshot))
          throw new Error('presentation_package_state_invalid')
        await saveWithHistory(
          PACKAGE_XML_KEY,
          JSON.stringify(snapshot),
          {
            id: historyEntryId('package_xml', snapshot),
            kind: 'package_xml',
            record: snapshot,
            legacy: false,
            sequence: 1,
          },
          () => {
            packageWriteFailed = true
          },
        )
      }
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
    readNativeMasterChange,
    writeNativeMasterChange(
      record: PresentationNativeMasterChange,
      expectedRecord: PresentationNativeMasterChange | undefined,
    ) {
      const snapshot = structuredClone(record),
        expected = structuredClone(expectedRecord)
      const write = async () => {
        if (
          !validatePresentationNativeMasterChange(snapshot) ||
          (expected !== undefined && !validatePresentationNativeMasterChange(expected))
        )
          throw new Error('presentation_native_master_state_invalid')
        if ((await documentId()) !== snapshot.documentId)
          throw new Error('presentation_document_changed')
        const prior = readNativeMasterChange(snapshot.changeId)
        if (JSON.stringify(prior) !== JSON.stringify(expected))
          throw new Error('presentation_native_master_stale')
        if (JSON.stringify(prior) === JSON.stringify(snapshot)) return
        if (!validNativeMasterTransition(prior, snapshot))
          throw new Error('presentation_native_master_state_invalid')
        await saveWithHistory(
          NATIVE_MASTER_KEY,
          JSON.stringify(snapshot),
          {
            id: historyEntryId('native_master', snapshot),
            kind: 'native_master',
            record: snapshot,
            legacy: false,
            sequence: 1,
          },
          () => {
            nativeMasterWriteFailed = true
          },
        )
      }
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
    writeExistingChartChange(
      record: PresentationExistingChartChange,
      expectedChange: PresentationExistingChartChange | undefined,
    ) {
      const snapshot = structuredClone(record),
        expected = structuredClone(expectedChange)
      const write = async () => {
        if (
          !validatePresentationExistingChartChange(snapshot) ||
          (expected !== undefined && !validatePresentationExistingChartChange(expected))
        )
          throw new Error('presentation_existing_chart_state_invalid')
        if ((await documentId()) !== snapshot.documentId)
          throw new Error('presentation_document_changed')
        const prior = readExistingChartChange(snapshot.changeId)
        if (JSON.stringify(prior) !== JSON.stringify(expected))
          throw new Error('presentation_existing_chart_stale')
        if (JSON.stringify(prior) === JSON.stringify(snapshot)) return
        if (!validExistingChartTransition(prior, snapshot))
          throw new Error('presentation_existing_chart_state_invalid')
        await saveWithHistory(
          EXISTING_CHART_KEY,
          JSON.stringify(snapshot),
          {
            id: historyEntryId('existing_chart', snapshot),
            kind: 'existing_chart',
            record: snapshot,
            legacy: false,
            sequence: 1,
          },
          () => {
            existingChartWriteFailed = true
          },
        )
      }
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
    writeExistingPageChange(
      record: PresentationExistingPageChange,
      expectedChange: PresentationExistingPageChange | undefined,
    ) {
      const snapshot = structuredClone(record),
        expected = structuredClone(expectedChange)
      const write = async () => {
        if (
          !validatePresentationExistingPageChange(snapshot) ||
          (expected !== undefined && !validatePresentationExistingPageChange(expected))
        )
          throw new Error('presentation_existing_page_state_invalid')
        if ((await documentId()) !== snapshot.documentId)
          throw new Error('presentation_document_changed')
        const prior = readExistingPageChange(snapshot.changeId)
        if (JSON.stringify(prior) !== JSON.stringify(expected))
          throw new Error('presentation_existing_page_stale')
        if (JSON.stringify(prior) === JSON.stringify(snapshot)) return
        if (!validExistingPageTransition(prior, snapshot))
          throw new Error('presentation_existing_page_state_invalid')
        await saveWithHistory(
          EXISTING_PAGE_KEY,
          JSON.stringify(snapshot),
          {
            id: historyEntryId('existing_page', snapshot),
            kind: 'existing_page',
            record: snapshot,
            legacy: false,
            sequence: 1,
          },
          () => {
            existingPageWriteFailed = true
          },
        )
      }
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
    writeExistingImageChange(
      record: PresentationExistingImageChange,
      expectedChange: PresentationExistingImageChange | undefined,
    ) {
      const snapshot = structuredClone(record),
        expected = structuredClone(expectedChange)
      const write = async () => {
        if (
          !validatePresentationExistingImageChange(snapshot) ||
          (expected !== undefined && !validatePresentationExistingImageChange(expected))
        )
          throw new Error('presentation_existing_image_state_invalid')
        if ((await documentId()) !== snapshot.documentId)
          throw new Error('presentation_document_changed')
        const prior = readExistingImageChange(snapshot.changeId)
        if (JSON.stringify(prior) !== JSON.stringify(expected))
          throw new Error('presentation_existing_image_stale')
        if (JSON.stringify(prior) === JSON.stringify(snapshot)) return
        if (!validExistingImageTransition(prior, snapshot))
          throw new Error('presentation_existing_image_state_invalid')
        await saveWithHistory(
          EXISTING_IMAGE_KEY,
          JSON.stringify(snapshot),
          {
            id: historyEntryId('existing_image', snapshot),
            kind: 'existing_image',
            record: snapshot,
            legacy: false,
            sequence: 1,
          },
          () => {
            existingImageWriteFailed = true
          },
        )
      }
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
    writeExistingBatch(
      record: PresentationExistingBatch,
      expectedBatch: PresentationExistingBatch | undefined,
    ) {
      const snapshot = structuredClone(record),
        expected = structuredClone(expectedBatch)
      const write = async () => {
        if (
          !validatePresentationExistingBatch(snapshot) ||
          (expected !== undefined && !validatePresentationExistingBatch(expected))
        )
          throw new Error('presentation_existing_batch_state_invalid')
        if ((await documentId()) !== snapshot.documentId)
          throw new Error('presentation_document_changed')
        const prior = readExistingBatch(snapshot.changeId)
        if (JSON.stringify(prior) !== JSON.stringify(expected))
          throw new Error('presentation_existing_batch_stale')
        if (JSON.stringify(prior) === JSON.stringify(snapshot)) return
        if (!validExistingBatchTransition(prior, snapshot))
          throw new Error('presentation_existing_batch_state_invalid')
        await saveWithHistory(
          EXISTING_BATCH_KEY,
          JSON.stringify(snapshot),
          {
            id: historyEntryId('existing_batch', snapshot),
            kind: 'existing_batch',
            record: snapshot,
            legacy: false,
            sequence: 1,
          },
          () => {
            existingBatchWriteFailed = true
          },
        )
      }
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
    writeExistingChange(
      record: PresentationExistingChange,
      expectedChange: PresentationExistingChange | undefined,
    ) {
      const snapshot = structuredClone(record),
        expected = structuredClone(expectedChange)
      const write = async () => {
        const invalid = () => new Error('presentation_existing_change_state_invalid')
        if (
          !validatePresentationExistingChange(snapshot) ||
          (expected !== undefined && !validatePresentationExistingChange(expected))
        )
          throw invalid()
        if ((await documentId()) !== snapshot.documentId)
          throw new Error('presentation_document_changed')
        const prior = readExistingChange(snapshot.changeId)
        if (JSON.stringify(prior) !== JSON.stringify(expected))
          throw new Error('presentation_existing_change_stale')
        if (JSON.stringify(prior) === JSON.stringify(snapshot)) return
        if (prior) {
          const core = (r: PresentationExistingChange) =>
            JSON.stringify({
              ...r,
              state: undefined,
              review: undefined,
              backupReleasedAt: undefined,
            })
          const transitions = {
            pending: 'applied',
            applied: 'undo_pending',
            undo_pending: 'undone',
            undone:
              prior.backup && prior.beforeSlideIds && !prior.backupReleasedAt
                ? 'pending'
                : undefined,
          }
          const backupRejournal =
            prior.state === 'pending' &&
            snapshot.state === 'pending' &&
            prior.backup !== undefined &&
            snapshot.backup !== undefined &&
            prior.backup.backupId !== snapshot.backup.backupId &&
            (prior.backup.sha256 !== snapshot.backup.sha256 ||
              prior.backup.sizeBytes !== snapshot.backup.sizeBytes) &&
            prior.backupReleasedAt === undefined &&
            snapshot.backupReleasedAt === undefined &&
            JSON.stringify(prior.review) === JSON.stringify(snapshot.review) &&
            JSON.stringify({
              ...prior,
              backup: {
                hostSlideId: prior.backup.hostSlideId,
                packageDigest: prior.backup.packageDigest,
              },
            }) ===
              JSON.stringify({
                ...snapshot,
                backup: {
                  hostSlideId: snapshot.backup.hostSlideId,
                  packageDigest: snapshot.backup.packageDigest,
                },
              })
          if (
            (core(prior) !== core(snapshot) && !backupRejournal) ||
            (prior.backupReleasedAt !== snapshot.backupReleasedAt &&
              !(
                prior.state === 'undone' &&
                snapshot.state === 'undone' &&
                prior.backupReleasedAt === undefined &&
                snapshot.backupReleasedAt !== undefined &&
                JSON.stringify(prior.review) === JSON.stringify(snapshot.review)
              )) ||
            (prior.state !== snapshot.state &&
              (transitions[prior.state] !== snapshot.state || snapshot.review !== undefined))
          )
            throw invalid()
        } else if (snapshot.state !== 'pending' || snapshot.review !== undefined) throw invalid()
        await saveWithHistory(
          EXISTING_KEY,
          JSON.stringify(snapshot),
          {
            id: historyEntryId('existing', snapshot),
            kind: 'existing',
            record: snapshot,
            legacy: false,
            sequence: 1,
          },
          () => {
            existingWriteFailed = true
          },
        )
      }
      const result = receiptQueue.then(write)
      receiptQueue = result.catch(() => {})
      return result
    },
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
    listReceipts() {
      const { receipts } = readReplacementEnvelope()
      return structuredClone(
        Object.entries({ ...readImports(), ...receipts })
          .filter((entry): entry is [string, PresentationImportRecord] => entry[1] !== null)
          .sort(([a], [b]) => a.localeCompare(b))
          .map(([key, record]) => ({ key, record })),
      )
    },
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
        if (record) {
          const storedRecord = { ...record }
          delete storedRecord.agentRunId
          if (record.toolCallId) {
            let activeRun: Record<string, unknown> | undefined
            try {
              const raw = settings.get(AGENT_RUN_KEY)
              if (typeof raw === 'string') activeRun = JSON.parse(raw) as Record<string, unknown>
            } catch {
              /* An unreadable checkpoint cannot establish the originating run. */
            }
            const runId =
              activeRun?.documentId === record.documentId &&
              activeRun?.phase === 'tool_pending' &&
              activeRun?.toolCallId === record.toolCallId &&
              ['import_generated_presentation', 'import_presentation_production'].includes(
                String(activeRun?.toolName),
              ) &&
              validId(activeRun?.runId)
                ? activeRun.runId
                : imports[key]?.toolCallId === record.toolCallId
                  ? imports[key]?.agentRunId
                  : undefined
            if (runId) storedRecord.agentRunId = runId
          }
          imports[key] = storedRecord
        } else delete imports[key]
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
    readQaAttempts(key: string): PresentationQaAttempt[] {
      const current = JSON.stringify([settings.get(ID_KEY), settings.location()])
      return Object.values(readRawQaAttempts())
        .filter((a) => qaAttemptKey(a) === key && a.documentId === current)
        .sort((a, b) => a.startedAt.localeCompare(b.startedAt) || a.id.localeCompare(b.id))
    },
    writeQaAttempt(key: string, supplied: PresentationQaAttempt): Promise<void> {
      const snapshot = structuredClone(supplied)
      const write = async () => {
        const attempt = parsePresentationQaAttempt(snapshot)
        if (qaAttemptKey(attempt) !== key) throw Error('presentation_qa_attempt_state_invalid')
        if ((await documentId()) !== attempt.documentId)
          throw Error('presentation_document_changed')
        const records = readRawQaAttempts(),
          previousAttempt = records[attempt.id]
        if (previousAttempt) {
          if (
            presentationQaAttemptIdentity(attempt) !==
            presentationQaAttemptIdentity(previousAttempt)
          )
            throw Error('presentation_qa_attempt_state_invalid')
          if (qaAttemptValue(attempt) === qaAttemptValue(previousAttempt)) return
          if (previousAttempt.status !== 'started' || attempt.status === 'started')
            throw Error('presentation_qa_attempt_state_invalid')
        } else if (attempt.status !== 'started')
          throw Error('presentation_qa_attempt_state_invalid')
        records[attempt.id] = attempt
        const oldestFinished = Object.values(records)
          .filter((a) => a.status !== 'started' && a.id !== attempt.id)
          .sort(
            (a, b) =>
              a.finishedAt!.localeCompare(b.finishedAt!) ||
              a.startedAt.localeCompare(b.startedAt) ||
              a.id.localeCompare(b.id),
          )
        let serialized = JSON.stringify(records)
        while (
          Object.keys(records).length > 64 ||
          qaAttemptBytes(serialized, records) > 128 * 1024
        ) {
          const oldest = oldestFinished.shift()
          if (!oldest) throw Error('presentation_qa_attempt_history_full')
          delete records[oldest.id]
          serialized = JSON.stringify(records)
        }
        const previous = settings.get(QA_ATTEMPTS_KEY),
          location = settings.location()
        try {
          settings.set(QA_ATTEMPTS_KEY, serialized)
          await settings.save()
          if (settings.location() !== location || settings.get(QA_ATTEMPTS_KEY) !== serialized)
            throw Error('presentation_document_changed')
        } catch (error) {
          if (settings.location() === location && settings.get(QA_ATTEMPTS_KEY) === serialized) {
            try {
              settings.set(QA_ATTEMPTS_KEY, typeof previous === 'string' ? previous : '{}')
            } catch {
              qaAttemptWriteFailed = true
            }
          } else qaAttemptWriteFailed = true
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
                matches(page.hostSlideId) && !page.recheckRequired
                  ? {
                      ...page,
                      recheckRequired: true as const,
                      invalidatedAt: new Date(
                        Math.max(
                          Date.now(),
                          Date.parse(page.capturedAt),
                          Date.parse(page.visual.reviewedAt ?? '') || 0,
                        ),
                      ).toISOString(),
                    }
                  : page,
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
    agentRunRecovery(boundDocumentId: string): PresentationAgentRunRecovery | undefined {
      const raw = settings.get(AGENT_RUN_KEY)
      if (typeof raw !== 'string' || new TextEncoder().encode(raw).byteLength > 4096)
        return undefined
      try {
        const value = JSON.parse(raw) as Record<string, unknown>
        const valid =
          value &&
          !Array.isArray(value) &&
          value.documentId === boundDocumentId &&
          validId(value.runId) &&
          typeof value.startedAt === 'number' &&
          Number.isSafeInteger(value.startedAt) &&
          value.startedAt > 0 &&
          value.startedAt <= Date.now()
        if (!valid) return undefined
        const keys = Object.keys(value).sort().join(',')
        if (keys === 'documentId,runId,startedAt')
          return { runId: value.runId as string, instruction: '', phase: 'tool_pending' }
        if (
          Object.keys(value).some(
            (key) =>
              ![
                'documentId',
                'runId',
                'startedAt',
                'phase',
                'toolName',
                'toolCallId',
                'instruction',
                'restartSafe',
              ].includes(key),
          ) ||
          !Object.hasOwn(value, 'phase')
        )
          return undefined
        if (
          !['running', 'tool_pending', 'tool_completed'].includes(value.phase as string) ||
          (value.restartSafe !== undefined && typeof value.restartSafe !== 'boolean') ||
          (value.toolName !== undefined &&
            (typeof value.toolName !== 'string' || value.toolName.length > 128)) ||
          (value.toolCallId !== undefined &&
            (typeof value.toolCallId !== 'string' ||
              value.toolCallId.length < 1 ||
              value.toolCallId.length > 256))
        )
          return undefined
        return {
          runId: value.runId as string,
          instruction: '',
          phase: value.phase as PresentationAgentRunRecovery['phase'],
          ...(typeof value.toolName === 'string' ? { toolName: value.toolName } : {}),
          ...(typeof value.toolCallId === 'string' ? { toolCallId: value.toolCallId } : {}),
          ...(typeof value.restartSafe === 'boolean'
            ? {
                restartSafe:
                  value.restartSafe &&
                  (value.phase === 'running' || restartSafeTools.has(String(value.toolName))),
              }
            : {}),
        }
      } catch {
        return undefined
      }
    },
    agentImportReceipt(boundDocumentId: string, runId: string, toolCallId: string) {
      if (
        !validId(runId) ||
        typeof toolCallId !== 'string' ||
        toolCallId.length < 1 ||
        toolCallId.length > 256
      )
        return undefined
      const matching = Object.values(readImports()).filter(
        (record) =>
          record.documentId === boundDocumentId &&
          record.agentRunId === runId &&
          record.toolCallId === toolCallId,
      )
      if (matching.length !== 1) return undefined
      const record = matching[0]!
      const completed = record.checkpoint?.completed.length ?? record.slideIds?.length ?? 0
      return {
        state:
          record.state === 'complete'
            ? ('complete' as const)
            : !record.checkpoint || record.checkpoint.inFlight
              ? ('uncertain' as const)
              : ('partial' as const),
        completed,
        ...(record.checkpoint ? { total: record.checkpoint.sourceSlideIds.length } : {}),
      }
    },
    agentChangeReceipt(boundDocumentId: string, runId: string, toolCallId: string) {
      if (
        !validId(runId) ||
        typeof toolCallId !== 'string' ||
        toolCallId.length < 1 ||
        toolCallId.length > 256
      )
        return undefined
      const entries = readHistory().entries.filter(
        (entry) =>
          entry.record.documentId === boundDocumentId &&
          entry.agentRunId === runId &&
          entry.toolCallId === toolCallId,
      )
      if (!entries.length) return undefined
      return {
        total: entries.length,
        unresolved: entries.filter(
          (entry) =>
            (entry.kind === 'package_xml' && Boolean(entry.record.pending)) ||
            !['applied', 'undone', 'discarded', 'complete', 'cancelled'].includes(
              entry.record.state,
            ),
        ).length,
      }
    },
    interruptedAgentRun(boundDocumentId: string): boolean {
      return Boolean(this.agentRunRecovery(boundDocumentId))
    },
    async scrubAgentRunPrompt(boundDocumentId: string) {
      return queueRunCheckpoint(async () => {
        if ((await documentId()) !== boundDocumentId)
          throw new Error('presentation_document_changed')
        const previous = settings.get(AGENT_RUN_KEY)
        if (typeof previous !== 'string' || !this.agentRunRecovery(boundDocumentId)) return
        const record = JSON.parse(previous) as Record<string, unknown>
        if (!Object.hasOwn(record, 'instruction')) return
        delete record.instruction
        const raw = JSON.stringify(record)
        settings.set(AGENT_RUN_KEY, raw)
        try {
          await settings.save()
        } catch (error) {
          settings.set(AGENT_RUN_KEY, previous)
          throw error
        }
        if ((await documentId()) !== boundDocumentId || settings.get(AGENT_RUN_KEY) !== raw)
          throw new Error('presentation_document_changed')
      })
    },
    async rememberAgentRun(boundDocumentId: string, runId: string) {
      return queueRunCheckpoint(async () => {
        if (!validId(runId) || (await documentId()) !== boundDocumentId)
          throw new Error('presentation_document_changed')
        const previous = settings.get(AGENT_RUN_KEY)
        const raw = JSON.stringify({
          documentId: boundDocumentId,
          runId,
          startedAt: Date.now(),
          phase: 'running',
          restartSafe: true,
        })
        if (new TextEncoder().encode(raw).byteLength > 4096)
          throw new Error('presentation_run_checkpoint_unavailable')
        settings.set(AGENT_RUN_KEY, raw)
        try {
          await settings.save()
        } catch (error) {
          settings.set(AGENT_RUN_KEY, typeof previous === 'string' ? previous : '')
          throw error
        }
        if ((await documentId()) !== boundDocumentId || settings.get(AGENT_RUN_KEY) !== raw)
          throw new Error('presentation_document_changed')
      })
    },
    async updateAgentRun(
      boundDocumentId: string,
      runId: string,
      phase: 'tool_pending' | 'tool_completed',
      toolName: string,
      mutated = false,
      toolCallId?: string,
    ) {
      return queueRunCheckpoint(async () => {
        if (
          (await documentId()) !== boundDocumentId ||
          toolName.length > 128 ||
          (toolCallId !== undefined && (toolCallId.length < 1 || toolCallId.length > 256))
        )
          throw new Error('presentation_document_changed')
        const previous = settings.get(AGENT_RUN_KEY)
        const current = this.agentRunRecovery(boundDocumentId)
        if (!current || current.runId !== runId || typeof previous !== 'string')
          throw new Error('presentation_run_checkpoint_unavailable')
        if (
          (phase === 'tool_pending' && current.phase === 'tool_pending') ||
          (phase === 'tool_completed' &&
            (current.phase !== 'tool_pending' ||
              current.toolName !== toolName ||
              current.toolCallId !== toolCallId))
        )
          throw new Error('presentation_run_checkpoint_unavailable')
        const next = JSON.parse(previous) as Record<string, unknown>
        delete next.instruction
        delete next.toolCallId
        const raw = JSON.stringify({
          ...next,
          phase,
          toolName,
          ...(toolCallId === undefined ? {} : { toolCallId }),
          restartSafe: current.restartSafe === true && restartSafeTools.has(toolName) && !mutated,
        })
        if (new TextEncoder().encode(raw).byteLength > 4096)
          throw new Error('presentation_run_checkpoint_unavailable')
        settings.set(AGENT_RUN_KEY, raw)
        try {
          await settings.save()
        } catch (error) {
          settings.set(AGENT_RUN_KEY, previous)
          throw error
        }
        if ((await documentId()) !== boundDocumentId || settings.get(AGENT_RUN_KEY) !== raw)
          throw new Error('presentation_document_changed')
      })
    },
    async finishAgentRun(boundDocumentId: string, runId: string) {
      return queueRunCheckpoint(async () => {
        if ((await documentId()) !== boundDocumentId)
          throw new Error('presentation_document_changed')
        const raw = settings.get(AGENT_RUN_KEY)
        if (typeof raw !== 'string') return
        let current: { documentId?: unknown; runId?: unknown }
        try {
          current = JSON.parse(raw)
        } catch {
          return
        }
        if (current.documentId !== boundDocumentId || current.runId !== runId) return
        settings.set(AGENT_RUN_KEY, '')
        try {
          await settings.save()
        } catch (error) {
          settings.set(AGENT_RUN_KEY, raw)
          throw error
        }
        if ((await documentId()) !== boundDocumentId)
          throw new Error('presentation_document_changed')
      })
    },
    selectedProduction(projectId: string, boundDocumentId: string): string | undefined {
      const raw = settings.get(SELECTED_PRODUCTION_KEY)
      if (typeof raw !== 'string' || raw.length > 4600) return undefined
      try {
        const value = JSON.parse(raw) as Record<string, unknown>
        if (
          !value ||
          typeof value !== 'object' ||
          Array.isArray(value) ||
          Object.keys(value).sort().join(',') !== 'documentId,projectId,requestId' ||
          value.documentId !== boundDocumentId ||
          value.projectId !== projectId ||
          !validId(value.requestId)
        )
          return undefined
        return value.requestId
      } catch {
        return undefined
      }
    },
    async rememberSelectedProduction(
      projectId: string,
      boundDocumentId: string,
      requestId: string,
    ) {
      if (!validId(projectId) || !validId(requestId)) throw new Error('invalid_tool_input')
      if ((await documentId()) !== boundDocumentId) throw new Error('presentation_document_changed')
      const previous = settings.get(SELECTED_PRODUCTION_KEY)
      const raw = JSON.stringify({ documentId: boundDocumentId, projectId, requestId })
      settings.set(SELECTED_PRODUCTION_KEY, raw)
      try {
        await settings.save()
      } catch (error) {
        settings.set(SELECTED_PRODUCTION_KEY, typeof previous === 'string' ? previous : '')
        throw error
      }
      if ((await documentId()) !== boundDocumentId || settings.get(SELECTED_PRODUCTION_KEY) !== raw)
        throw new Error('presentation_document_changed')
    },
    async rememberProject(projectId: string) {
      if (!validId(projectId)) throw new Error('invalid_tool_input')
      settings.set(PROJECT_KEY, projectId)
      await settings.save()
    },
  }
}

export function createPresentationAgentRunCheckpoint(
  binding: Pick<
    ReturnType<typeof createPresentationDocumentBinding>,
    | 'documentId'
    | 'agentRunRecovery'
    | 'agentImportReceipt'
    | 'agentChangeReceipt'
    | 'rememberAgentRun'
    | 'updateAgentRun'
    | 'finishAgentRun'
  >,
  boundDocumentId: string,
  storage?: Pick<Storage, 'getItem' | 'setItem' | 'removeItem'> &
    Partial<Pick<Storage, 'key' | 'length'>>,
) {
  const runDocuments = new Map<string, string>()
  const mirroredRuns = new Set<string>()
  const localKey = (runId: string) => `${AGENT_RUN_LOCAL_PREFIX}${runId}`
  const sweepExpiredPrompts = () => {
    if (!storage?.key || typeof storage.length !== 'number') return
    try {
      const now = Date.now()
      for (let index = Math.min(storage.length, 4096) - 1; index >= 0; index -= 1) {
        const key = storage.key(index)
        if (!key?.startsWith(AGENT_RUN_LOCAL_PREFIX)) continue
        const raw = storage.getItem(key)
        let expiresAt: unknown
        try {
          expiresAt =
            raw && new TextEncoder().encode(raw).byteLength <= AGENT_RUN_LOCAL_RECORD_LIMIT
              ? (JSON.parse(raw) as Record<string, unknown>).expiresAt
              : undefined
        } catch {
          /* malformed local entry */
        }
        if (
          typeof expiresAt !== 'number' ||
          !Number.isSafeInteger(expiresAt) ||
          expiresAt <= now ||
          expiresAt > now + AGENT_RUN_LOCAL_TTL_MS
        )
          storage.removeItem(key)
      }
    } catch {
      /* local storage may be disabled */
    }
  }
  sweepExpiredPrompts()
  const recovery = (): PresentationAgentRunRecovery | undefined => {
    const saved = binding.agentRunRecovery(boundDocumentId)
    if (!saved) return undefined
    let importReceipt: PresentationAgentRunRecovery['importReceipt']
    let changeReceipt: PresentationAgentRunRecovery['changeReceipt']
    if (
      saved.toolCallId &&
      ['import_generated_presentation', 'import_presentation_production'].includes(
        saved.toolName ?? '',
      )
    )
      try {
        importReceipt = binding.agentImportReceipt(boundDocumentId, saved.runId, saved.toolCallId)
      } catch {
        /* A damaged import journal cannot justify resuming a write. */
      }
    if (saved.toolCallId)
      try {
        changeReceipt = binding.agentChangeReceipt(boundDocumentId, saved.runId, saved.toolCallId)
      } catch {
        /* A damaged change journal cannot justify resuming a write. */
      }
    const record = {
      ...saved,
      ...(importReceipt ? { importReceipt } : {}),
      ...(changeReceipt ? { changeReceipt, restartSafe: false, instruction: '' } : {}),
    }
    if (!storage) return record
    try {
      const raw = storage.getItem(localKey(record.runId))
      if (!raw) return record
      if (new TextEncoder().encode(raw).byteLength > AGENT_RUN_LOCAL_RECORD_LIMIT) {
        storage.removeItem(localKey(record.runId))
        return record
      }
      const value = JSON.parse(raw) as Record<string, unknown>
      if (
        !value ||
        Array.isArray(value) ||
        Object.keys(value).sort().join(',') !==
          [
            'documentId',
            'expiresAt',
            'instruction',
            'phase',
            'restartSafe',
            'runId',
            ...(value.toolCallId === undefined ? [] : ['toolCallId']),
            ...(value.toolName === undefined ? [] : ['toolName']),
            ...(value.messages === undefined ? [] : ['messages']),
          ]
            .sort()
            .join(',') ||
        value.documentId !== boundDocumentId ||
        value.runId !== record.runId ||
        typeof value.expiresAt !== 'number' ||
        !Number.isSafeInteger(value.expiresAt) ||
        value.expiresAt <= Date.now() ||
        value.expiresAt > Date.now() + AGENT_RUN_LOCAL_TTL_MS ||
        typeof value.instruction !== 'string' ||
        new TextEncoder().encode(value.instruction).byteLength > AGENT_RUN_INSTRUCTION_LIMIT ||
        !['running', 'tool_pending', 'tool_completed'].includes(String(value.phase)) ||
        typeof value.restartSafe !== 'boolean' ||
        (value.toolCallId !== undefined &&
          (typeof value.toolCallId !== 'string' ||
            value.toolCallId.length < 1 ||
            value.toolCallId.length > 256)) ||
        (value.toolName !== undefined && typeof value.toolName !== 'string')
      ) {
        storage.removeItem(localKey(record.runId))
        return record
      }
      if (
        value.phase !== record.phase ||
        value.toolName !== record.toolName ||
        value.toolCallId !== record.toolCallId ||
        value.restartSafe !== record.restartSafe
      )
        return { ...record, instruction: '', restartSafe: false }
      const messages = !changeReceipt
        ? recoveryMessages(value.messages, { ...record, instruction: value.instruction })
        : undefined
      return {
        ...record,
        instruction: changeReceipt ? '' : value.instruction,
        ...(messages ? { messages } : {}),
      }
    } catch {
      return record
    }
  }
  return {
    recovery,
    async adopt(runId: string, messages: readonly AgentMessage[]) {
      const current = recovery()
      if (
        !current?.messages ||
        current.runId !== runId ||
        JSON.stringify(current.messages) !== JSON.stringify(messages)
      )
        throw new Error('presentation_run_checkpoint_unavailable')
      const captured = JSON.stringify(current)
      if ((await binding.documentId()) !== boundDocumentId)
        throw new Error('presentation_document_changed')
      if (JSON.stringify(recovery()) !== captured)
        throw new Error('presentation_run_checkpoint_unavailable')
      runDocuments.set(runId, boundDocumentId)
      mirroredRuns.add(runId)
    },
    async begin(runId: string, instruction = '') {
      sweepExpiredPrompts()
      if ((await binding.documentId()) !== boundDocumentId)
        throw new Error('presentation_document_changed')
      const recoverableInstruction =
        new TextEncoder().encode(instruction).byteLength <= AGENT_RUN_INSTRUCTION_LIMIT
          ? instruction
          : ''
      const previous = binding.agentRunRecovery(boundDocumentId)
      await binding.rememberAgentRun(boundDocumentId, runId)
      try {
        if (storage && recoverableInstruction) {
          const localRecord = JSON.stringify({
            documentId: boundDocumentId,
            runId,
            instruction: recoverableInstruction,
            expiresAt: Date.now() + AGENT_RUN_LOCAL_TTL_MS,
            phase: 'running',
            restartSafe: true,
          })
          if (new TextEncoder().encode(localRecord).byteLength <= AGENT_RUN_LOCAL_RECORD_LIMIT) {
            storage.setItem(localKey(runId), localRecord)
            if (storage.getItem(localKey(runId)) !== localRecord)
              throw new Error('presentation_run_checkpoint_unavailable')
            mirroredRuns.add(runId)
          }
        }
      } catch {
        await binding.finishAgentRun(boundDocumentId, runId)
        try {
          storage?.removeItem(localKey(runId))
        } catch {
          /* browser storage unavailable */
        }
        throw new Error('presentation_run_checkpoint_unavailable')
      }
      if ((await binding.documentId()) !== boundDocumentId) {
        await binding.finishAgentRun(boundDocumentId, runId).catch(() => undefined)
        try {
          storage?.removeItem(localKey(runId))
        } catch {
          /* browser storage unavailable */
        }
        throw new Error('presentation_document_changed')
      }
      if (previous && previous.runId !== runId) {
        try {
          storage?.removeItem(localKey(previous.runId))
        } catch {
          /* stale local cache */
        }
      }
      runDocuments.set(runId, boundDocumentId)
    },
    async tool(
      runId: string,
      phase: 'tool_pending' | 'tool_completed',
      toolName: string,
      mutated = false,
      toolCallId?: string,
    ) {
      const id = runDocuments.get(runId)
      if (!id) throw new Error('presentation_run_checkpoint_unavailable')
      let localRecord: Record<string, unknown> | undefined
      if (mirroredRuns.has(runId)) {
        try {
          const raw = storage!.getItem(localKey(runId))
          const current = binding.agentRunRecovery(id)
          if (
            !raw ||
            !current ||
            new TextEncoder().encode(raw).byteLength > AGENT_RUN_LOCAL_RECORD_LIMIT
          )
            throw new Error('presentation_run_checkpoint_unavailable')
          localRecord = JSON.parse(raw) as Record<string, unknown>
          if (
            !localRecord ||
            localRecord.documentId !== id ||
            localRecord.runId !== runId ||
            localRecord.phase !== current.phase ||
            localRecord.toolName !== current.toolName ||
            localRecord.toolCallId !== current.toolCallId ||
            localRecord.restartSafe !== current.restartSafe
          )
            throw new Error('presentation_run_checkpoint_unavailable')
        } catch {
          throw new Error('presentation_run_checkpoint_unavailable')
        }
      }
      await binding.updateAgentRun(id, runId, phase, toolName, mutated, toolCallId)
      if (localRecord) {
        const current = binding.agentRunRecovery(id)
        if (!current || current.runId !== runId)
          throw new Error('presentation_run_checkpoint_unavailable')
        try {
          const nextLocal = { ...localRecord }
          delete nextLocal.toolCallId
          if (phase === 'tool_pending') delete nextLocal.messages
          const raw = JSON.stringify({
            ...nextLocal,
            phase: current.phase,
            toolName: current.toolName,
            ...(current.toolCallId === undefined ? {} : { toolCallId: current.toolCallId }),
            restartSafe: current.restartSafe === true,
          })
          storage!.setItem(localKey(runId), raw)
          if (storage!.getItem(localKey(runId)) !== raw)
            throw new Error('presentation_run_checkpoint_unavailable')
        } catch {
          throw new Error('presentation_run_checkpoint_unavailable')
        }
      }
    },
    async conversation(runId: string, value: readonly AgentMessage[]) {
      const id = runDocuments.get(runId)
      if (!id || (await binding.documentId()) !== id)
        throw new Error('presentation_document_changed')
      if (!mirroredRuns.has(runId)) return
      const current = binding.agentRunRecovery(id)
      if (!current || current.runId !== runId)
        throw new Error('presentation_run_checkpoint_unavailable')
      try {
        const raw = storage!.getItem(localKey(runId))
        if (!raw || new TextEncoder().encode(raw).byteLength > AGENT_RUN_LOCAL_RECORD_LIMIT)
          throw new Error('presentation_run_checkpoint_unavailable')
        const localRecord = JSON.parse(raw) as Record<string, unknown>
        if (
          localRecord.documentId !== id ||
          localRecord.runId !== runId ||
          localRecord.phase !== current.phase ||
          localRecord.toolName !== current.toolName ||
          localRecord.toolCallId !== current.toolCallId ||
          localRecord.restartSafe !== current.restartSafe
        )
          throw new Error('presentation_run_checkpoint_unavailable')
        const messages = recoveryMessages(value, {
          ...current,
          instruction: typeof localRecord.instruction === 'string' ? localRecord.instruction : '',
        })
        delete localRecord.messages
        const next = JSON.stringify({ ...localRecord, ...(messages ? { messages } : {}) })
        if (new TextEncoder().encode(next).byteLength > AGENT_RUN_LOCAL_RECORD_LIMIT)
          throw new Error('presentation_run_checkpoint_unavailable')
        storage!.setItem(localKey(runId), next)
        if (storage!.getItem(localKey(runId)) !== next)
          throw new Error('presentation_run_checkpoint_unavailable')
      } catch {
        throw new Error('presentation_run_checkpoint_unavailable')
      }
      if ((await binding.documentId()) !== id || binding.agentRunRecovery(id)?.runId !== runId)
        throw new Error('presentation_document_changed')
    },
    async finish(runId: string) {
      const id = runDocuments.get(runId)
      if (!id) return
      await binding.finishAgentRun(id, runId)
      try {
        storage?.removeItem(localKey(runId))
      } catch {
        /* checkpoint is already cleared */
      }
      runDocuments.delete(runId)
      mirroredRuns.delete(runId)
    },
  }
}

export async function preparePresentationAgentRunRecovery(
  binding: Pick<
    ReturnType<typeof createPresentationDocumentBinding>,
    'scrubAgentRunPrompt' | 'interruptedAgentRun'
  >,
  boundDocumentId: string,
): Promise<{ interrupted: boolean; scrubFailed: boolean }> {
  try {
    await binding.scrubAgentRunPrompt(boundDocumentId)
    return { interrupted: binding.interruptedAgentRun(boundDocumentId), scrubFailed: false }
  } catch (error) {
    if (error instanceof Error && error.message === 'presentation_document_changed') throw error
    return { interrupted: binding.interruptedAgentRun(boundDocumentId), scrubFailed: true }
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
