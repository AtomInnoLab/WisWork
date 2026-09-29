import {
  validatePresentationNativeMasterChange,
  nativeMasterReservedBytes,
  type PresentationNativeMasterChange,
} from './presentation-native-master-change.js'
import {
  validatePresentationExistingChange,
  existingChangeReservedBytes,
  type PresentationExistingChange,
} from './presentation-existing-change.js'
import {
  validatePresentationExistingImageChange,
  existingImageReservedBytes,
  type PresentationExistingImageChange,
} from './presentation-existing-image.js'
import {
  validatePresentationExistingPageChange,
  existingPageReservedBytes,
  type PresentationExistingPageChange,
} from './presentation-existing-page.js'
import {
  validatePresentationExistingChartChange,
  existingChartReservedBytes,
  type PresentationExistingChartChange,
} from './presentation-existing-chart.js'
import {
  validatePresentationExistingBatch,
  existingBatchReservedBytes,
  type PresentationExistingBatch,
} from './presentation-existing-batch.js'
import {
  validatePresentationTextChange,
  type PresentationTextChange,
} from './presentation-text-change.js'
import {
  validatePresentationGeometryChange,
  type PresentationGeometryChange,
} from './presentation-geometry-change.js'
import {
  validateImageReplacementRecord,
  imageReplacementReservedBytes,
  type ImageReplacementRecord,
} from './presentation-image-replacement-record.js'
import {
  validatePresentationPageReplacement,
  type PresentationPageReplacement,
} from './presentation-page-replacement-record.js'
type Records = {
  native_master: PresentationNativeMasterChange
  existing: PresentationExistingChange
  existing_image: PresentationExistingImageChange
  existing_page: PresentationExistingPageChange
  existing_chart: PresentationExistingChartChange
  existing_batch: PresentationExistingBatch
  text: PresentationTextChange
  geometry: PresentationGeometryChange
  image: ImageReplacementRecord
  page: PresentationPageReplacement
}
export type PresentationHistoryEntry = {
  [K in keyof Records]: {
    id: string
    sequence: number
    legacy: boolean
    kind: K
    record: Records[K]
  }
}[keyof Records] & {
  agentRunId?: string
  toolCallId?: string
  checkpointCreatedAt?: string
  checkpointRestoredAt?: string
}
export function historyEntryId<K extends keyof Records>(kind: K, record: Records[K]): string {
  if (kind === 'image') {
    const r = record as ImageReplacementRecord
    return `image:${JSON.stringify([r.documentId, r.source ?? null, r.projectId, r.requestId, r.pageId, r.oldShapeId])}`
  }
  return `${kind}:${(record as PresentationTextChange).changeId}`
}
export function validatePresentationHistoryEntry(
  value: unknown,
): value is PresentationHistoryEntry {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const e = value as PresentationHistoryEntry
  const eventKeys = ['checkpointCreatedAt', 'checkpointRestoredAt']
  const date = (value: unknown) =>
    typeof value === 'string' &&
    /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(value) &&
    Number.isFinite(Date.parse(value)) &&
    new Date(value).toISOString() === value
  if (
    ![
      'id,kind,legacy,record,sequence',
      'agentRunId,id,kind,legacy,record,sequence,toolCallId',
    ].includes(
      Object.keys(e)
        .filter((key) => !eventKeys.includes(key))
        .sort()
        .join(','),
    ) ||
    (e.checkpointCreatedAt !== undefined && (!date(e.checkpointCreatedAt) || e.legacy)) ||
    (e.checkpointRestoredAt !== undefined &&
      (!date(e.checkpointRestoredAt) ||
        (e.checkpointCreatedAt !== undefined && e.checkpointRestoredAt < e.checkpointCreatedAt))) ||
    (e.agentRunId !== undefined &&
      (typeof e.agentRunId !== 'string' ||
        !/^[A-Za-z0-9_-]{1,128}$/.test(e.agentRunId) ||
        typeof e.toolCallId !== 'string' ||
        e.toolCallId.length < 1 ||
        e.toolCallId.length > 256)) ||
    !Number.isSafeInteger(e.sequence) ||
    e.sequence < 1 ||
    typeof e.legacy !== 'boolean'
  )
    return false
  const valid =
    e.kind === 'native_master'
      ? validatePresentationNativeMasterChange(e.record)
      : e.kind === 'text'
        ? validatePresentationTextChange(e.record)
        : e.kind === 'geometry'
          ? validatePresentationGeometryChange(e.record)
          : e.kind === 'image'
            ? validateImageReplacementRecord(e.record)
            : e.kind === 'existing_batch'
              ? validatePresentationExistingBatch(e.record)
              : e.kind === 'existing_image'
                ? validatePresentationExistingImageChange(e.record)
                : e.kind === 'existing_page'
                  ? validatePresentationExistingPageChange(e.record)
                  : e.kind === 'existing_chart'
                    ? validatePresentationExistingChartChange(e.record)
                    : e.kind === 'existing'
                      ? validatePresentationExistingChange(e.record)
                      : e.kind === 'page'
                        ? validatePresentationPageReplacement(e.record)
                        : false
  return valid && e.id === historyEntryId(e.kind, e.record)
}
export interface PresentationHistoryEnvelope {
  version: 1
  entries: PresentationHistoryEntry[]
  heads: Partial<
    Record<
      | 'text'
      | 'geometry'
      | 'page'
      | 'existing'
      | 'existing_batch'
      | 'existing_image'
      | 'existing_page'
      | 'existing_chart'
      | 'native_master',
      string
    >
  >
}
export const presentationHistoryBytes = (history: PresentationHistoryEnvelope) =>
  new TextEncoder().encode(JSON.stringify(history)).byteLength +
  // Updating an older record can lengthen the current-head ID without adding an entry.
  (
    [
      'text',
      'geometry',
      'page',
      'existing',
      'existing_batch',
      'existing_image',
      'existing_page',
      'existing_chart',
      'native_master',
    ] as const
  ).reduce(
    (sum, kind) =>
      sum +
      (history.heads[kind] === undefined ? 0 : kind.length + 1 + 128 - history.heads[kind]!.length),
    0,
  ) +
  history.entries.reduce(
    (sum, e) =>
      sum +
      // Fixed timestamp metadata has a separate bounded allowance, preserving old full histories.
      -(e.checkpointCreatedAt || e.checkpointRestoredAt
        ? new TextEncoder().encode(
            JSON.stringify({
              checkpointCreatedAt: e.checkpointCreatedAt,
              checkpointRestoredAt: e.checkpointRestoredAt,
            }),
          ).byteLength - 1
        : 0) +
      (e.kind === 'native_master'
        ? nativeMasterReservedBytes(e.record)
        : e.kind === 'existing_batch'
          ? existingBatchReservedBytes(e.record)
          : e.kind === 'existing_image'
            ? existingImageReservedBytes(e.record)
            : e.kind === 'existing_page'
              ? existingPageReservedBytes(e.record)
              : e.kind === 'existing_chart'
                ? existingChartReservedBytes(e.record)
                : e.kind === 'existing'
                  ? existingChangeReservedBytes(e.record)
                  : e.kind === 'page'
                    ? Math.max(
                        0,
                        192 * 1024 - new TextEncoder().encode(JSON.stringify(e.record)).byteLength,
                      )
                    : e.kind === 'image'
                      ? imageReplacementReservedBytes(e.record)
                      : 'undo_pending'.length - e.record.state.length),
    0,
  )
