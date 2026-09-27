import {
  validatePresentationHistoryEntry,
  type PresentationHistoryEntry,
} from './presentation-change-history.js'

export interface PresentationPreferenceCandidate {
  changeId: string
  kind: 'text' | 'geometry'
  pageId: string
  before: string
  after: string
  status: 'candidate'
}

/** Confirmed edits are observations only. They never amend a saved plan or brand kit. */
export function presentationPreferenceCandidates(
  entries: PresentationHistoryEntry[],
  documentId: string,
  projectId: string,
): PresentationPreferenceCandidate[] {
  if (
    !Array.isArray(entries) || entries.length > 64 ||
    entries.some((entry) => !validatePresentationHistoryEntry(entry)) ||
    new Set(entries.map((entry) => entry.id)).size !== entries.length ||
    new Set(entries.map((entry) => entry.sequence)).size !== entries.length ||
    new TextEncoder().encode(JSON.stringify(entries)).byteLength > 1024 * 1024
  ) throw new Error('presentation_change_history_invalid')
  return entries
    .filter((entry) =>
      (entry.kind === 'text' || entry.kind === 'geometry') &&
      entry.record.documentId === documentId &&
      entry.record.projectId === projectId &&
      entry.record.state === 'applied')
    .slice(-8)
    .map((entry) => {
      if (entry.kind !== 'text' && entry.kind !== 'geometry') throw new Error('unreachable')
      const short = (value: string) => value.slice(0, 240)
      return {
        changeId: entry.record.changeId,
        kind: entry.kind,
        pageId: entry.record.pageId,
        before: short(entry.kind === 'text' ? entry.record.before : JSON.stringify(entry.record.before)),
        after: short(entry.kind === 'text' ? entry.record.after : JSON.stringify(entry.record.after)),
        status: 'candidate' as const,
      }
    })
}
