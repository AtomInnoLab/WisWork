import type { PresentationExistingChange } from '../skills/powerpoint/presentation-existing-change.js'
import type { PresentationExistingBatch } from '../skills/powerpoint/presentation-existing-batch.js'
import type { PresentationExistingImageChange } from '../skills/powerpoint/presentation-existing-image.js'
import type { PresentationExistingPageChange } from '../skills/powerpoint/presentation-existing-page.js'
import type { PresentationExistingChartChange } from '../skills/powerpoint/presentation-existing-chart.js'
import {
  presentationChangeSetSummary,
  selectPresentationHistory,
  type PresentationChangeSetSummary,
} from '../skills/powerpoint/presentation-history.js'
import {
  validatePresentationHistoryEntry,
  type PresentationHistoryEntry,
} from '../skills/powerpoint/presentation-change-history.js'
import type { AgentSkill } from '@wiswork/agent-core'
import type { CompiledPresentationArtifact } from '../skills/powerpoint/presentation-delivery.js'
import { presentationArtifactContent } from '../skills/powerpoint/presentation-page-delivery.js'
import {
  validatePresentationTextChange,
  type PresentationTextChange,
} from '../skills/powerpoint/presentation-text-change.js'
import {
  validatePresentationGeometryChange,
  type PresentationGeometryChange,
} from '../skills/powerpoint/presentation-geometry-change.js'
import {
  validatePresentationPageReplacement,
  type PresentationPageReplacement,
} from '../skills/powerpoint/presentation-page-replacement-record.js'
import {
  validateImageReplacementRecord,
  type ImageReplacementRecord,
} from '../skills/powerpoint/presentation-image-replacement-record.js'

export type PresentationChangeAction =
  'inspect' | 'undo' | 'resume' | 'commit' | 'discard' | 'release'
export interface PresentationChangeEntry {
  source?: 'existing' | 'existing_batch' | 'existing_image' | 'existing_page' | 'existing_chart'
  review?: PresentationExistingChange['review']
  reviews?: PresentationExistingBatch['reviews']
  visualReviews?: PresentationExistingPageChange['reviews']
  visualCaptures?: PresentationExistingPageChange['captures']
  visualPageIds?: string[]
  affectedPageCount?: number
  sequence?: number
  legacy?: boolean
  changeSet?: PresentationChangeSetSummary
  id: string
  kind: 'text' | 'geometry' | 'table_cell' | 'image' | 'page' | 'chart'
  pageId: string
  state: string
  before: string
  after: string
  actions: PresentationChangeAction[]
}
export interface PresentationChangesSnapshot {
  phase: 'idle' | 'loading' | 'acting'
  projectId?: string
  requestId?: string
  entries: PresentationChangeEntry[]
  backupAudit?: { active: number; unmatched: number }
  notice?: string
  error?: string
}
export interface PresentationChangesController {
  snapshot(): PresentationChangesSnapshot
  subscribe(listener: () => void): () => void
  refresh(): Promise<void>
  run(entryId: string, action: PresentationChangeAction): Promise<void>
  clear(): void
}
type Read<T> = () => T | undefined | Promise<T | undefined>
export interface PresentationChangesOptions {
  listChangeHistory?: () => PresentationHistoryEntry[]
  listExistingPageBackups?: (documentId: string) => Promise<
    {
      backupId: string
      status: string
      hostSlideId: string
      slideIds: string[]
      sha256: string
      sizeBytes: number
    }[]
  >
  existingAvailable?: () => boolean
  available(): boolean
  artifact(): CompiledPresentationArtifact | undefined
  documentId(): Promise<string>
  readTextChange?: Read<PresentationTextChange>
  readGeometryChange?: Read<PresentationGeometryChange>
  readPageReplacement?: Read<PresentationPageReplacement>
  listImageReplacements?: () => ImageReplacementRecord[] | Promise<ImageReplacementRecord[]>
  executeTool: AgentSkill['executeTool']
}
type RecordValue =
  | PresentationTextChange
  | PresentationGeometryChange
  | PresentationPageReplacement
  | ImageReplacementRecord
interface SavedEntry {
  entry: PresentationChangeEntry
  record:
    | RecordValue
    | PresentationExistingChange
    | PresentationExistingBatch
    | PresentationExistingImageChange
    | PresentationExistingPageChange
    | PresentationExistingChartChange
  fingerprint: string
  historical?: boolean
}
const copy = <T>(value: T): T => structuredClone(value)
const pageActions: Record<PresentationPageReplacement['state'], PresentationChangeAction[]> = {
  pending: ['inspect'],
  inserted: ['inspect', 'resume'],
  staged: ['inspect', 'commit', 'discard'],
  discard_pending: ['inspect', 'discard'],
  commit_pending: ['inspect', 'commit'],
  applied: ['inspect', 'undo'],
  undo_pending: ['inspect', 'undo'],
  restore_inserted: ['inspect', 'undo'],
  discarded: [],
  undone: [],
}
function entry(kind: PresentationChangeEntry['kind'], record: RecordValue): SavedEntry {
  let before: string, after: string, actions: PresentationChangeAction[]
  if (kind === 'image') {
    const r = record as ImageReplacementRecord
    before = `图片对象：${r.oldShapeId}${r.baseline ? `\n资源摘要：${r.baseline.mediaDigest}` : ''}${r.backup ? '\n原图已持久备份' : '\n无可撤销原图备份'}`
    after = `图片对象：${r.newShapeId ?? '尚未记录'}\n资源摘要：${r.assetDigest}${r.restoredShapeId ? `\n恢复图片：${r.restoredShapeId}\n恢复资源摘要：${r.baseline!.mediaDigest}` : ''}`
    actions =
      r.state === 'pending'
        ? ['inspect', ...(r.baseline && r.newShapeId ? ['resume' as const] : [])]
        : r.state === 'complete' && r.backup && r.after
          ? ['undo']
          : r.state === 'undo_pending'
            ? ['inspect', ...(r.restoredShapeId ? ['resume' as const] : [])]
            : []
  } else if (kind === 'page') {
    const r = record as PresentationPageReplacement
    before = `页面：${r.oldSlideId}\n包摘要：${r.originalPackageDigest}`
    after = `页面：${r.newSlideId ?? '尚未记录'}\n包摘要：${r.replacementPackageDigest}${r.restoredSlideId ? `\n恢复页面：${r.restoredSlideId}\n恢复包摘要：${r.originalPackageDigest}` : ''}`
    actions = pageActions[r.state]
  } else {
    const r = record as PresentationTextChange | PresentationGeometryChange
    before = typeof r.before === 'string' ? r.before : JSON.stringify(r.before, null, 2)
    after = typeof r.after === 'string' ? r.after : JSON.stringify(r.after, null, 2)
    actions = r.state === 'applied' ? ['undo'] : r.state === 'undone' ? [] : ['inspect', 'resume']
  }
  const id =
    kind +
    ':' +
    ('changeId' in record ? record.changeId : JSON.stringify([record.pageId, record.oldShapeId]))
  return {
    entry: {
      id,
      kind,
      pageId: record.pageId,
      state: record.state,
      before,
      after,
      actions: [...actions],
    },
    record: copy(record),
    fingerprint: JSON.stringify(record),
  }
}
function inspectionNotice(output: string): string {
  const messages: Record<string, string> = {
    ready_to_apply: '可准备继续执行；仍需确认提案。',
    already_applied: '目标内容已存在；继续操作仅补齐保存回执，仍需确认提案。',
    ready_to_finish: '可准备完成图片替换；仍需确认提案。',
    manual_review: '需要人工检查，不能自动重放。',
    not_pending: '此保存点没有待恢复操作。',
    baseline: '原页面基线仍在，尚未确认替换完成。',
    staged: '新页面已暂存，原页面仍保留。',
    applied: '整页替换已应用。',
    restore_staged: '恢复页面已暂存，需继续完成撤销。',
    undone: '整页替换已撤销。',
    conflict: '需要人工检查：页面现状与保存点不一致。',
  }
  try {
    if (output.length <= 256 * 1024) {
      const value = JSON.parse(output)
      const status = value?.inspection?.status ?? value?.status
      if (typeof status === 'string' && Object.hasOwn(messages, status)) {
        const receipts =
          typeof value?.visualReceipt === 'string'
            ? [value.visualReceipt]
            : Array.isArray(value?.visualReceipts) && value.visualReceipts.length <= 2
              ? value.visualReceipts.map((item: { status?: unknown }) => item?.status)
              : []
        const visual = receipts.includes('different')
          ? '当前截图与历史回执不同，需重新复核。'
          : receipts.includes('unavailable')
            ? '当前截图暂无法与历史回执核对。'
            : receipts.length && receipts.every((item: unknown) => item === 'matched')
              ? '当前截图与历史回执一致。'
              : ''
        return messages[status] + (visual ? ` ${visual}` : '') + ' 此检查不代表页面 QA 通过。'
      }
    }
  } catch {
    /* Tool output is untrusted; keep unknown results out of the UI. */
  }
  return '检查已完成。操作资格仍以工具实时检查为准。'
}
export function createPresentationChangesController(
  options: PresentationChangesOptions,
): PresentationChangesController {
  let state: PresentationChangesSnapshot = { phase: 'idle', entries: [] }
  let saved: SavedEntry[] = [],
    bound: string | undefined,
    boundDocument: string | undefined,
    generation = 0
  let abort: AbortController | undefined
  const listeners = new Set<() => void>()
  const publish = () => {
    for (const listener of listeners) listener()
  }
  const scopeIdentity = () =>
    JSON.stringify({
      artifact: options.artifact(),
      generated: options.available(),
      existing: options.existingAvailable?.() ?? false,
    })
  const current = (scope: string | undefined, documentId: string) =>
    scopeIdentity() === scope && (!boundDocument || boundDocument === documentId)
  async function read(
    artifact: CompiledPresentationArtifact | undefined,
    scope: string,
    ticket: number,
  ) {
    const documentId = await options.documentId()
    if (ticket !== generation || !current(scope, documentId)) throw new Error('stale')
    if (options.listChangeHistory) {
      const history = copy(options.listChangeHistory())
      if (
        !Array.isArray(history) ||
        history.length > 64 ||
        history.some((e) => !validatePresentationHistoryEntry(e)) ||
        new Set(history.map((e) => e.id)).size !== history.length ||
        new Set(history.map((e) => e.sequence)).size !== history.length ||
        new TextEncoder().encode(JSON.stringify(history)).byteLength > 1024 * 1024
      )
        throw new Error('invalid')
      const selected: PresentationHistoryEntry[] =
        artifact && options.available()
          ? await selectPresentationHistory(history, artifact, documentId)
          : []
      if (options.existingAvailable?.())
        selected.push(
          ...history.filter(
            (e) =>
              (e.kind === 'existing' ||
                e.kind === 'existing_batch' ||
                e.kind === 'existing_image' ||
                e.kind === 'existing_page' ||
                e.kind === 'existing_chart') &&
              e.record.documentId === documentId,
          ),
        )
      const latestDocumentId = await options.documentId()
      if (
        ticket !== generation ||
        !current(scope, latestDocumentId) ||
        documentId !== latestDocumentId ||
        JSON.stringify(history) !== JSON.stringify(options.listChangeHistory())
      )
        throw new Error('stale')
      boundDocument = documentId
      return selected
        .sort((a, b) => Number(a.legacy) - Number(b.legacy) || b.sequence - a.sequence)
        .map((saved) => {
          const row: SavedEntry =
            saved.kind === 'existing_chart'
              ? {
                  entry: {
                    id: saved.id,
                    source: 'existing_chart',
                    kind: 'chart',
                    pageId: saved.record.newSlideId ?? saved.record.oldSlideId,
                    state: saved.record.state,
                    affectedPageCount: 1,
                    before: `原页：${saved.record.oldSlideId}\n图表：${saved.record.shapeId}\n包摘要：${saved.record.beforePackageDigest}\n${saved.record.backupReleasedAt ? `备份已释放：${saved.record.backupReleasedAt}` : '原页已持久备份'}`,
                    after: `新页：${saved.record.newSlideId ?? '尚未记录'}\n包摘要：${saved.record.afterPackageDigest}${saved.record.restoredSlideId ? `\n恢复页：${saved.record.restoredSlideId}` : ''}`,
                    actions:
                      saved.record.state === 'applied'
                        ? ['inspect', 'undo']
                        : ['undone', 'cancelled'].includes(saved.record.state)
                          ? [
                              'inspect',
                              ...(saved.record.backupReleasedAt ? [] : ['release' as const]),
                            ]
                          : ['inspect', 'resume'],
                  },
                  record: copy(saved.record),
                  fingerprint: JSON.stringify(saved),
                }
              : saved.kind === 'existing_page'
                ? {
                    entry: {
                      id: saved.id,
                      source: 'existing_page',
                      kind: 'page',
                      pageId: saved.record.oldSlideId,
                      state: saved.record.state,
                      visualReviews: copy(saved.record.reviews),
                      visualCaptures: copy(saved.record.captures),
                      visualPageIds:
                        saved.record.state === 'staged'
                          ? [saved.record.oldSlideId, saved.record.newSlideId!]
                          : saved.record.state === 'applied'
                            ? [saved.record.newSlideId!]
                            : saved.record.state === 'discarded'
                              ? [saved.record.oldSlideId]
                              : saved.record.state === 'undone'
                                ? [saved.record.restoredSlideId!]
                                : [],
                      affectedPageCount: saved.record.state === 'staged' ? 2 : 1,
                      before: `原页：${saved.record.oldSlideId}\n包摘要：${saved.record.originalPackageDigest}\n${saved.record.backupReleasedAt ? `备份已释放：${saved.record.backupReleasedAt}` : '原页已持久备份'}`,
                      after: `新页：${saved.record.newSlideId ?? '尚未记录'}\n包摘要：${saved.record.replacementPackageDigest}${saved.record.restoredSlideId ? `\n恢复页面：${saved.record.restoredSlideId}` : ''}`,
                      actions: [
                        ...pageActions[saved.record.state],
                        ...(['discarded', 'undone'].includes(saved.record.state) &&
                        !saved.record.backupReleasedAt
                          ? ['release' as const]
                          : []),
                      ],
                    },
                    record: copy(saved.record),
                    fingerprint: JSON.stringify(saved),
                  }
                : saved.kind === 'existing_image'
                  ? {
                      entry: {
                        id: saved.id,
                        source: 'existing_image',
                        kind: 'image',
                        pageId: saved.record.hostSlideId,
                        state: saved.record.state,
                        visualReviews: saved.record.review ? [copy(saved.record.review)] : [],
                        visualCaptures: saved.record.capture ? [copy(saved.record.capture)] : [],
                        visualPageIds: ['complete', 'undone'].includes(saved.record.state)
                          ? [saved.record.hostSlideId]
                          : [],
                        affectedPageCount: 1,
                        before: `原图：${saved.record.oldShapeId}\n媒体摘要：${saved.record.original.mediaDigest}\n原图已持久备份`,
                        after: `新图：${saved.record.insertedShapeId ?? '尚未记录'}\n媒体摘要：${saved.record.assetDigest}${saved.record.restoredShapeId ? `\n恢复图片：${saved.record.restoredShapeId}` : ''}`,
                        actions:
                          saved.record.state === 'complete'
                            ? ['inspect', 'undo']
                            : saved.record.state === 'undone'
                              ? ['inspect']
                              : ['inspect', 'resume'],
                      },
                      record: copy(saved.record),
                      fingerprint: JSON.stringify(saved),
                    }
                  : saved.kind === 'existing_batch'
                    ? {
                        entry: {
                          id: saved.id,
                          source: 'existing_batch',
                          kind: saved.record.operations[0].kind,
                          pageId: saved.record.operations[0].hostSlideId,
                          state: saved.record.state,
                          reviews: copy(saved.record.reviews),
                          affectedPageCount: new Set(
                            saved.record.operations.map((op) => op.hostSlideId),
                          ).size,
                          before: saved.record.operations
                            .map(
                              (op) =>
                                `${op.hostSlideId}/${op.shapeId}${op.kind === 'table_cell' ? `[${op.rowIndex},${op.columnIndex}]` : ''}: ${JSON.stringify(op.before)}`,
                            )
                            .join('\n'),
                          after: saved.record.operations
                            .map(
                              (op) =>
                                `${op.hostSlideId}/${op.shapeId}${op.kind === 'table_cell' ? `[${op.rowIndex},${op.columnIndex}]` : ''}: ${JSON.stringify(op.after)}`,
                            )
                            .join('\n'),
                          actions:
                            saved.record.state === 'applied'
                              ? ['inspect', 'undo']
                              : saved.record.state === 'undone'
                                ? saved.record.backups?.length && !saved.record.backupReleasedAt
                                  ? ['inspect', 'release']
                                  : ['inspect']
                                : ['inspect', 'resume'],
                        },
                        record: copy(saved.record),
                        fingerprint: JSON.stringify(saved),
                      }
                    : saved.kind === 'existing'
                      ? {
                          entry: {
                            id: saved.id,
                            source: 'existing',
                            kind: saved.record.kind,
                            pageId: saved.record.hostSlideId,
                            state: saved.record.state,
                            before:
                              saved.record.kind === 'table_cell'
                                ? `单元格 (${saved.record.rowIndex}, ${saved.record.columnIndex}): ${saved.record.before}`
                                : typeof saved.record.before === 'string'
                                  ? saved.record.before
                                  : JSON.stringify(saved.record.before, null, 2),
                            after:
                              saved.record.kind === 'table_cell'
                                ? `单元格 (${saved.record.rowIndex}, ${saved.record.columnIndex}): ${saved.record.after}`
                                : typeof saved.record.after === 'string'
                                  ? saved.record.after
                                  : JSON.stringify(saved.record.after, null, 2),
                            review: copy(saved.record.review),
                            actions:
                              saved.record.state === 'applied'
                                ? ['inspect', 'undo']
                                : saved.record.state === 'undone'
                                  ? ['inspect']
                                  : ['inspect', 'resume'],
                          },
                          record: copy(saved.record),
                          fingerprint: JSON.stringify(saved),
                        }
                      : entry(saved.kind, saved.record)
          row.entry = {
            ...row.entry,
            id: saved.id,
            sequence: saved.sequence,
            legacy: saved.legacy,
            changeSet: presentationChangeSetSummary(saved),
          }
          row.historical = true
          row.fingerprint = JSON.stringify(saved)
          return row
        })
    }
    if (!artifact || !options.available() || artifact.documentId !== documentId) return []
    const [text, geometry, page, images] = await Promise.all([
      options.readTextChange?.(),
      options.readGeometryChange?.(),
      options.readPageReplacement?.(),
      options.listImageReplacements?.() ?? [],
    ])
    if (
      (text !== undefined && !validatePresentationTextChange(text)) ||
      (geometry !== undefined && !validatePresentationGeometryChange(geometry)) ||
      (page !== undefined && !validatePresentationPageReplacement(page)) ||
      !Array.isArray(images) ||
      images.length > 32 ||
      images.some((r) => !validateImageReplacementRecord(r))
    )
      throw new Error('invalid')
    const digest = Array.from(
      new Uint8Array(
        await crypto.subtle.digest(
          'SHA-256',
          new TextEncoder().encode(presentationArtifactContent(artifact)),
        ),
      ),
      (b) => b.toString(16).padStart(2, '0'),
    ).join('')
    const source = artifact.pagePptxBase64 === undefined ? undefined : 'production'
    const matches = (r: RecordValue) =>
      r.documentId === documentId &&
      r.projectId === artifact.projectId &&
      artifact.pages?.some((p) => p.id === r.pageId)
    const rows: SavedEntry[] = []
    for (const [kind, r] of [
      ['text', text],
      ['geometry', geometry],
    ] as const)
      if (
        r &&
        matches(r) &&
        r.requestId === artifact.requestId &&
        r.source === source &&
        r.artifactDigest === digest
      )
        rows.push(entry(kind, r))
    for (const r of images)
      if (matches(r) && r.requestId === artifact.requestId && r.source === source)
        rows.push(entry('image', r))
    if (
      page &&
      source === 'production' &&
      matches(page) &&
      [page.parentRequestId, page.requestId].includes(artifact.requestId)
    )
      rows.push(entry('page', page))
    if (new Set(rows.map((r) => r.entry.id)).size !== rows.length) throw new Error('invalid')
    if (ticket !== generation || !current(scope, await options.documentId()))
      throw new Error('stale')
    boundDocument = documentId
    return rows
  }
  async function refresh() {
    // Storage notifications during a tool action are collected by the final refresh.
    if (state.phase === 'acting') {
      if (scopeIdentity() === bound) return
      ++generation
      abort?.abort()
      abort = undefined
      state = { phase: 'idle', entries: [] }
      saved = []
      bound = undefined
      boundDocument = undefined
    }
    const ticket = ++generation
    const artifact = options.artifact(),
      scope = scopeIdentity()
    if ((!options.available() || !artifact) && !options.existingAvailable?.()) {
      saved = []
      bound = undefined
      boundDocument = undefined
      state = { phase: 'idle', entries: [], notice: '恢复当前任务后可查看最近保存点。' }
      publish()
      return
    }
    boundDocument = undefined
    state = {
      phase: 'loading',
      projectId: artifact?.projectId,
      requestId: artifact?.requestId,
      entries: [],
    }
    publish()
    try {
      const rows = await read(copy(artifact), scope!, ticket)
      if (ticket !== generation) return
      let backupAudit: PresentationChangesSnapshot['backupAudit']
      if (options.listExistingPageBackups && boundDocument) {
        try {
          const backups = await options.listExistingPageBackups(boundDocument)
          if (
            !Array.isArray(backups) ||
            backups.length > 8 ||
            backups.some(
              (b) =>
                !b ||
                typeof b.backupId !== 'string' ||
                !['ready', 'uploading'].includes(b.status) ||
                typeof b.hostSlideId !== 'string' ||
                !Array.isArray(b.slideIds) ||
                !b.slideIds.every((id) => typeof id === 'string') ||
                typeof b.sha256 !== 'string' ||
                !Number.isSafeInteger(b.sizeBytes),
            )
          )
            throw new Error('invalid')
          const known: {
            backupId: string
            sha256: string
            sizeBytes: number
            hostSlideId: string
            slideIds?: string[]
          }[] = []
          for (const row of rows) {
            if (row.entry.source === 'existing_batch') {
              const record = row.record as PresentationExistingBatch
              known.push(
                ...(record.backups ?? []).map((backup) => ({
                  backupId: backup.backupId,
                  sha256: backup.sha256,
                  sizeBytes: backup.sizeBytes,
                  hostSlideId: backup.hostSlideId,
                  slideIds: record.beforeSlideIds,
                })),
              )
            }
            if (row.entry.source === 'existing_chart' || row.entry.source === 'existing_page') {
              const record = row.record as
                PresentationExistingChartChange | PresentationExistingPageChange
              known.push({
                backupId: record.backup.backupId,
                sha256: record.backup.sha256,
                sizeBytes: record.backup.sizeBytes,
                hostSlideId: record.oldSlideId,
                slideIds: record.beforeSlideIds,
              })
            }
          }
          backupAudit = {
            active: backups.length,
            unmatched: backups.filter(
              (backup) =>
                !known.some(
                  (item) =>
                    item.backupId === backup.backupId &&
                    item.sha256 === backup.sha256 &&
                    item.sizeBytes === backup.sizeBytes &&
                    item.hostSlideId === backup.hostSlideId &&
                    JSON.stringify(item.slideIds) === JSON.stringify(backup.slideIds),
                ),
            ).length,
          }
        } catch {
          /* Backup inventory is advisory; history remains available. */
        }
      }
      if (ticket !== generation || !current(scope, await options.documentId())) return
      saved = rows
      bound = scope
      state = {
        phase: 'idle',
        projectId: artifact?.projectId,
        requestId: artifact?.requestId,
        entries: rows.map((r) => copy(r.entry)),
        backupAudit,
      }
    } catch {
      if (ticket !== generation) return
      saved = []
      bound = undefined
      boundDocument = undefined
      state = {
        phase: 'idle',
        entries: [],
        error: '保存点无法读取或当前文档/任务已变化，请刷新并检查记录。',
      }
    }
    publish()
  }
  return {
    snapshot: () => copy(state),
    subscribe(listener) {
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
    refresh,
    clear() {
      ++generation
      abort?.abort()
      abort = undefined
      saved = []
      bound = undefined
      boundDocument = undefined
      state = { phase: 'idle', entries: [] }
      publish()
    },
    async run(id, action) {
      if (state.phase !== 'idle') return
      const selected = saved.find((r) => r.entry.id === id),
        artifact = options.artifact(),
        scope = bound
      if (
        !selected ||
        !selected.entry.actions.includes(action) ||
        !scope ||
        scopeIdentity() !== scope
      ) {
        state = { ...state, error: '操作已过期或不可用，请刷新保存点。' }
        publish()
        return
      }
      const ticket = ++generation,
        cancellation = new AbortController()
      abort = cancellation
      state = { ...state, phase: 'acting', notice: undefined, error: undefined }
      publish()
      let notice: string | undefined, error: string | undefined
      try {
        const fresh = await read(copy(artifact), scope, ticket)
        if (ticket !== generation) return
        if (fresh.find((r) => r.entry.id === id)?.fingerprint !== selected.fingerprint)
          throw new Error('stale')
        const r = selected.record,
          kind = selected.entry.kind
        const suffix =
          kind === 'text'
            ? 'text_change'
            : kind === 'geometry'
              ? 'geometry_change'
              : kind === 'image'
                ? 'image_replacement'
                : 'page_replacement'
        const generated = r as RecordValue
        const input: Record<string, unknown> =
          selected.entry.source === 'existing' ||
          selected.entry.source === 'existing_batch' ||
          selected.entry.source === 'existing_image' ||
          selected.entry.source === 'existing_page' ||
          selected.entry.source === 'existing_chart'
            ? {
                change_id: (
                  r as
                    | PresentationExistingChange
                    | PresentationExistingBatch
                    | PresentationExistingImageChange
                    | PresentationExistingPageChange
                    | PresentationExistingChartChange
                ).changeId,
              }
            : {
                project_id: generated.projectId,
                ...(kind === 'page'
                  ? { change_id: (r as PresentationPageReplacement).changeId }
                  : { page_id: generated.pageId }),
                ...(kind === 'image' ? { shape_id: (r as ImageReplacementRecord).oldShapeId } : {}),
                ...(selected.historical && (kind === 'text' || kind === 'geometry')
                  ? {
                      change_id: (r as PresentationTextChange | PresentationGeometryChange)
                        .changeId,
                    }
                  : {}),
              }
        let result = await options.executeTool(
          {
            id: `change-${ticket}`,
            name:
              selected.entry.source === 'existing_page'
                ? `${action}_existing_presentation_page_change`
                : selected.entry.source === 'existing_chart'
                  ? `${action}_slide_chart_values_change`
                  : selected.entry.source === 'existing_image'
                    ? `${action}_existing_presentation_image_change`
                    : selected.entry.source === 'existing_batch'
                      ? `${action}_existing_presentation_batch`
                      : selected.entry.source === 'existing'
                        ? `${action}_existing_presentation_change`
                        : `${action}_presentation_${suffix}`,
            input,
          },
          cancellation.signal,
        )
        if ('kind' in result && result.kind === 'tool-execution-suspension')
          result = await result.result
        if (ticket !== generation) return
        if (!current(scope, await options.documentId())) throw new Error('stale')
        if (
          selected.entry.source === 'existing' ||
          selected.entry.source === 'existing_batch' ||
          selected.entry.source === 'existing_image' ||
          selected.entry.source === 'existing_page' ||
          selected.entry.source === 'existing_chart'
        ) {
          const latest = (await read(copy(artifact), scope, ticket)).find(
            (row) => row.entry.id === id,
          )
          if (!latest || ticket !== generation) throw new Error('stale')
          const core = (
            record:
              | RecordValue
              | PresentationExistingChange
              | PresentationExistingBatch
              | PresentationExistingImageChange
              | PresentationExistingPageChange
              | PresentationExistingChartChange,
          ) => {
            const {
              state: _state,
              review: _review,
              cursor: _cursor,
              newSlideId: _newSlideId,
              restoredSlideId: _restoredSlideId,
              backupReleasedAt: _backupReleasedAt,
              ...rest
            } = record as PresentationExistingChange & {
              cursor?: number
              newSlideId?: string
              restoredSlideId?: string
              backupReleasedAt?: string
            }
            return JSON.stringify(rest)
          }
          if (
            core(latest.record) !== core(selected.record) ||
            (action === 'inspect' && latest.fingerprint !== selected.fingerprint)
          )
            throw new Error('stale')
        }
        if (result.isError) throw new Error('tool')
        notice =
          action === 'inspect'
            ? inspectionNotice(result.output)
            : action === 'release'
              ? '备份释放提案已创建，确认后执行。'
              : '操作请求已处理；如有待确认提案，请确认后执行。变更后需重新采集页面 QA。'
      } catch {
        error = '操作未完成或保存点已变化，请刷新并检查；未自动重试。'
      } finally {
        if (ticket === generation) {
          abort = undefined
          state = { ...state, phase: 'idle' }
          await refresh()
          if (generation === ticket + 1) {
            state = { ...state, notice, error: error ?? state.error }
            publish()
          }
        }
      }
    },
  }
}
