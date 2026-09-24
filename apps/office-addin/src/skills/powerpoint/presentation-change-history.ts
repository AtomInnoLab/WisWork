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
}[keyof Records]
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
  if (
    Object.keys(e).length !== 5 ||
    !Number.isSafeInteger(e.sequence) ||
    e.sequence < 1 ||
    typeof e.legacy !== 'boolean'
  )
    return false
  const valid =
    e.kind === 'text'
      ? validatePresentationTextChange(e.record)
      : e.kind === 'geometry'
        ? validatePresentationGeometryChange(e.record)
        : e.kind === 'image'
          ? validateImageReplacementRecord(e.record)
          : e.kind === 'page'
            ? validatePresentationPageReplacement(e.record)
            : false
  return valid && e.id === historyEntryId(e.kind, e.record)
}
export interface PresentationHistoryEnvelope {
  version: 1
  entries: PresentationHistoryEntry[]
  heads: Partial<Record<'text' | 'geometry' | 'page', string>>
}
export const presentationHistoryBytes = (history: PresentationHistoryEnvelope) =>
  new TextEncoder().encode(JSON.stringify(history)).byteLength +
  // Updating an older record can lengthen the current-head ID without adding an entry.
  (['text', 'geometry', 'page'] as const).reduce(
    (sum, kind) =>
      sum +
      (history.heads[kind] === undefined ? 0 : kind.length + 1 + 128 - history.heads[kind]!.length),
    0,
  ) +
  history.entries.reduce(
    (sum, e) =>
      sum +
      (e.kind === 'page'
        ? Math.max(0, 192 * 1024 - new TextEncoder().encode(JSON.stringify(e.record)).byteLength)
        : e.kind === 'image'
          ? imageReplacementReservedBytes(e.record)
          : 'undo_pending'.length - e.record.state.length),
    0,
  )
