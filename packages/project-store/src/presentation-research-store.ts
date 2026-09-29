import { createHash, randomUUID } from 'node:crypto'
import { constants } from 'node:fs'
import { lstat, mkdir, open, readdir, rename, rm } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { canonicalPresentationValue } from './presentation-canonical'
import {
  parsePresentationResearchDraft,
  parsePresentationResearchRecord,
  parsePresentationResearchHistory,
  parsePresentationResearchDeleteReceipt,
  type PresentationResearchDeleteReceipt,
  summarizePresentationResearchHistory,
  presentationResearchChecks,
  type PresentationResearchDraft,
  type PresentationResearchRecord,
  type PresentationResearchHistory,
  type PresentationResearchEvidence,
} from './presentation-research'
const LIMIT = 64 * 1024 * 1024
const hash = (v: string) => createHash('sha256').update(v).digest('hex')
const locks = new Map<string, Promise<void>>()
interface StateV1 {
  version: 1
  documentId: string
  projectId: string
  revision: number
  totalRecords: number
  records: PresentationResearchRecord[]
}
interface StateV2 extends Omit<StateV1, 'version'> {
  version: 2
  lastSequence: number
  tombstones: PresentationResearchDeleteReceipt[]
}
type State = StateV1 | StateV2
function check(signal?: AbortSignal) {
  if (signal?.aborted) throw new Error('aborted')
}
function validId(value: string) {
  if (typeof value !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(value))
    throw new Error('invalid_request')
}
export type PresentationResearchFinish =
  | { state: 'completed'; sources: PresentationResearchEvidence[] }
  | {
      state: 'failed'
      error: 'aborted' | 'source_unavailable' | 'invalid_state'
      sources?: PresentationResearchEvidence[]
    }
async function directory(path: string, create = false) {
  try {
    if (create)
      await mkdir(path, { mode: 0o700 }).catch((e) => {
        if ((e as NodeJS.ErrnoException).code !== 'EEXIST') throw e
      })
    const s = await lstat(path)
    if (!s.isDirectory() || s.isSymbolicLink()) throw new Error('invalid_state')
    return true
  } catch (e) {
    if (!create && (e as NodeJS.ErrnoException).code === 'ENOENT') return false
    throw e
  }
}
// ponytail: one atomic project state keeps history and archives consistent; the 128-record/64MiB caps bound copy and parse cost.
export class PresentationResearchStore {
  private readonly root: string
  constructor(userDataPath: string) {
    this.root = join(resolve(userDataPath), 'presentation-research')
  }
  private path(documentId: string, projectId: string) {
    parsePresentationResearchHistory({
      version: 1,
      documentId,
      projectId,
      revision: 0,
      totalRecords: 0,
      records: [],
    })
    return join(this.root, hash(documentId), hash(projectId))
  }
  private async load(documentId: string, projectId: string): Promise<State> {
    const project = this.path(documentId, projectId)
    const empty: State = {
      version: 1,
      documentId,
      projectId,
      revision: 0,
      totalRecords: 0,
      records: [],
    }
    try {
      if (
        !(await directory(this.root)) ||
        !(await directory(dirname(project))) ||
        !(await directory(project))
      )
        return empty
      let f
      try {
        f = await open(
          join(project, 'state.json'),
          constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK,
        )
      } catch (e) {
        if ((e as NodeJS.ErrnoException).code === 'ENOENT') {
          if ((await readdir(project)).some((n) => !/^state\.json\.[a-f0-9-]{36}\.tmp$/.test(n)))
            throw new Error('invalid_state', { cause: e })
          return empty
        }
        throw e
      }
      let raw: Buffer
      try {
        const s = await f.stat()
        if (!s.isFile() || s.size > LIMIT) throw new Error('invalid_state')
        raw = await f.readFile()
        if (raw.length > LIMIT) throw new Error('invalid_state')
      } finally {
        await f.close()
      }
      const v = JSON.parse(raw.toString())
      if (
        !v ||
        Object.keys(v).sort().join(',') !== 'checksum,state' ||
        v.checksum !== hash(JSON.stringify(v.state))
      )
        throw new Error('invalid_state')
      const state = v.state as State
      if (
        !state ||
        Object.keys(state).sort().join(',') !==
          (state.version === 2
            ? 'documentId,lastSequence,projectId,records,revision,tombstones,totalRecords,version'
            : 'documentId,projectId,records,revision,totalRecords,version') ||
        state.documentId !== documentId ||
        state.projectId !== projectId ||
        ![1, 2].includes(state.version) ||
        !Array.isArray(state.records) ||
        state.records.length !== state.totalRecords ||
        state.totalRecords > 128
      )
        throw new Error('invalid_state')
      const seen = new Set<string>(),
        sequences = new Set<number>()
      let completed = 0,
        last = ''
      for (const [index, r] of state.records.entries()) {
        parsePresentationResearchRecord(r)
        if (
          r.documentId !== documentId ||
          r.projectId !== projectId ||
          (state.version === 1
            ? r.sequence !== index + 1
            : r.sequence > state.lastSequence ||
              (index > 0 && r.sequence <= state.records[index - 1]!.sequence)) ||
          seen.has(r.id) ||
          r.startedAt < last ||
          r.draftDigest !== hash(canonicalPresentationValue(r.draft))
        )
          throw new Error('invalid_state')
        seen.add(r.id)
        sequences.add(r.sequence)
        last = r.startedAt
        if (r.state !== 'running') completed++
      }
      if (state.version === 1) {
        if (state.revision !== state.totalRecords + completed) throw new Error('invalid_state')
      } else {
        if (
          !Number.isSafeInteger(state.lastSequence) ||
          !Array.isArray(state.tombstones) ||
          state.tombstones.length < 1 ||
          state.tombstones.length > 4096 ||
          state.lastSequence !== state.totalRecords + state.tombstones.length ||
          state.revision !== state.lastSequence + completed + 2 * state.tombstones.length
        )
          throw new Error('invalid_state')
        const deleteIds = new Set<string>()
        let previousRevision = 0,
          previousDate = '',
          highestDeleted = 0
        for (const [index, t] of state.tombstones.entries()) {
          parsePresentationResearchDeleteReceipt(t)
          highestDeleted = Math.max(highestDeleted, t.sequence)
          if (
            t.documentId !== documentId ||
            t.projectId !== projectId ||
            seen.has(t.ledgerId) ||
            sequences.has(t.sequence) ||
            deleteIds.has(t.deleteId) ||
            t.sequence > state.lastSequence ||
            t.revision <= previousRevision ||
            t.revision < highestDeleted + 2 * (index + 1) ||
            t.revision > Math.min(state.revision, 2 * state.lastSequence + index + 1) ||
            t.deletedAt < previousDate
          )
            throw new Error('invalid_state')
          seen.add(t.ledgerId)
          sequences.add(t.sequence)
          deleteIds.add(t.deleteId)
          previousRevision = t.revision
          previousDate = t.deletedAt
        }
        if (sequences.size !== state.lastSequence) throw new Error('invalid_state')
      }
      this.toHistory(state)
      return state
    } catch {
      throw new Error('invalid_state')
    }
  }
  private toHistory(state: State): PresentationResearchHistory {
    const { tombstones: _tombstones, ...history } = state as StateV2
    return parsePresentationResearchHistory({ ...history, records: state.records.slice(-32) })
  }
  async history(documentId: string, projectId: string) {
    return this.toHistory(await this.load(documentId, projectId))
  }
  async summary(documentId: string, projectId: string) {
    return summarizePresentationResearchHistory(await this.history(documentId, projectId))
  }
  async latestCompleted(
    documentId: string,
    projectId: string,
  ): Promise<PresentationResearchRecord | null> {
    const record = (await this.load(documentId, projectId)).records
      .slice()
      .reverse()
      .find((r) => r.state === 'completed')
    return record ? structuredClone(record) : null
  }
  async read(documentId: string, projectId: string, id: string) {
    validId(id)
    const state = await this.load(documentId, projectId)
    if (state.version === 2 && state.tombstones.some((t) => t.ledgerId === id))
      throw new Error('record_deleted')
    const record = state.records.find((r) => r.id === id)
    if (!record) throw new Error('not_found')
    return structuredClone(record)
  }
  private async update<T>(
    documentId: string,
    projectId: string,
    change: (state: State) => T,
    signal?: AbortSignal,
  ): Promise<T> {
    check(signal)
    const project = this.path(documentId, projectId)
    const previous = locks.get(project) ?? Promise.resolve()
    let release!: () => void
    const tail = new Promise<void>((r) => {
      release = r
    })
    locks.set(project, tail)
    await previous
    try {
      check(signal)
      const state = await this.load(documentId, projectId)
      check(signal)
      const result = change(state)
      this.toHistory(state)
      await directory(this.root, true)
      await directory(dirname(project), true)
      await directory(project, true)
      const pattern =
        /^state\.json\.[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}\.tmp$/
      for (const name of await readdir(project)) {
        if (!pattern.test(name)) continue
        const path = join(project, name),
          s = await lstat(path)
        if (!s.isFile() || s.isSymbolicLink()) throw new Error('invalid_state')
        await rm(path)
      }
      check(signal)
      const raw = JSON.stringify({ state, checksum: hash(JSON.stringify(state)) })
      if (Buffer.byteLength(raw) > LIMIT) throw new Error('quota_exceeded')
      const tmp = join(project, 'state.json.' + randomUUID() + '.tmp')
      try {
        const f = await open(tmp, 'wx', 0o600)
        try {
          await f.writeFile(raw)
          await f.sync()
        } finally {
          await f.close()
        }
        check(signal)
        await rename(tmp, join(project, 'state.json'))
        if (process.platform !== 'win32') {
          const d = await open(project, constants.O_RDONLY | constants.O_NOFOLLOW)
          try {
            await d.sync()
          } finally {
            await d.close()
          }
        }
      } finally {
        await rm(tmp, { force: true })
      }
      return structuredClone(result)
    } finally {
      release()
      if (locks.get(project) === tail) locks.delete(project)
    }
  }
  private lastTime(state: State) {
    return (
      [
        ...state.records.flatMap((r) =>
          r.finishedAt ? [r.startedAt, r.finishedAt] : [r.startedAt],
        ),
        ...(state.version === 2 ? state.tombstones.map((t) => t.deletedAt) : []),
      ]
        .sort()
        .at(-1) ?? ''
    )
  }
  async deletedReceipt(
    documentId: string,
    projectId: string,
    deleteId: string,
  ): Promise<PresentationResearchDeleteReceipt> {
    validId(deleteId)
    const state = await this.load(documentId, projectId)
    const receipt =
      state.version === 2 ? state.tombstones.find((t) => t.deleteId === deleteId) : undefined
    if (!receipt) throw new Error('not_found')
    return structuredClone(receipt)
  }
  deleteRecord(
    documentId: string,
    projectId: string,
    expectedRevision: number,
    deleteId: string,
    ledgerId: string,
    expectedDraftDigest: string,
    signal?: AbortSignal,
  ): Promise<PresentationResearchDeleteReceipt> {
    validId(deleteId)
    validId(ledgerId)
    if (
      !Number.isSafeInteger(expectedRevision) ||
      expectedRevision < 0 ||
      typeof expectedDraftDigest !== 'string' ||
      !/^[a-f0-9]{64}$/.test(expectedDraftDigest)
    )
      throw new Error('invalid_request')
    return this.update(
      documentId,
      projectId,
      (state) => {
        if (state.version === 2) {
          const receipt = state.tombstones.find((t) => t.deleteId === deleteId)
          if (receipt) {
            if (receipt.ledgerId !== ledgerId || receipt.draftDigest !== expectedDraftDigest)
              throw new Error('request_conflict')
            return receipt
          }
          if (state.tombstones.some((t) => t.ledgerId === ledgerId))
            throw new Error('record_deleted')
        }
        if (state.revision !== expectedRevision) throw new Error('revision_conflict')
        const record = state.records.find((r) => r.id === ledgerId)
        if (!record) throw new Error('not_found')
        if (record.draftDigest !== expectedDraftDigest) throw new Error('request_conflict')
        if (record.state === 'running') throw new Error('record_running')
        if (state.version === 2 && state.tombstones.length >= 4096)
          throw new Error('cleanup_quota_exceeded')
        const receipt = parsePresentationResearchDeleteReceipt({
          version: 1,
          documentId,
          projectId,
          ledgerId,
          sequence: record.sequence,
          draftDigest: record.draftDigest,
          deleteId,
          deletedAt: [new Date().toISOString(), this.lastTime(state)].sort().at(-1)!,
          revision: state.revision + 1,
        })
        if (state.version === 1)
          Object.assign(state, { version: 2, lastSequence: state.totalRecords, tombstones: [] })
        const next = state as StateV2
        next.records = next.records.filter((r) => r.id !== ledgerId)
        next.totalRecords--
        next.revision++
        next.tombstones.push(receipt)
        return receipt
      },
      signal,
    )
  }
  /** Explicitly end one exact unfinished record; this does not prove which request ended it. */
  abandon(
    documentId: string,
    projectId: string,
    expectedRevision: number,
    ledgerId: string,
    expectedDraftDigest: string,
    signal?: AbortSignal,
  ): Promise<PresentationResearchRecord> {
    validId(ledgerId)
    if (
      !Number.isSafeInteger(expectedRevision) ||
      expectedRevision < 0 ||
      typeof expectedDraftDigest !== 'string' ||
      !/^[a-f0-9]{64}$/.test(expectedDraftDigest)
    )
      throw new Error('invalid_request')
    return this.update(
      documentId,
      projectId,
      (state) => {
        const index = state.records.findIndex((r) => r.id === ledgerId),
          record = state.records[index]
        if (!record)
          throw new Error(
            state.version === 2 && state.tombstones.some((t) => t.ledgerId === ledgerId)
              ? 'record_deleted'
              : 'not_found',
          )
        if (record.draftDigest !== expectedDraftDigest) throw new Error('request_conflict')
        if (record.state === 'failed' && record.error === 'aborted') return record
        if (record.state !== 'running') throw new Error('record_not_running')
        if (state.revision !== expectedRevision) throw new Error('revision_conflict')
        const ended = parsePresentationResearchRecord({
          ...record,
          state: 'failed',
          error: 'aborted',
          finishedAt: [new Date().toISOString(), this.lastTime(state)].sort().at(-1)!,
        })
        state.records[index] = ended
        state.revision++
        return ended
      },
      signal,
    )
  }
  begin(
    documentId: string,
    projectId: string,
    expectedRevision: number,
    id: string,
    draft: PresentationResearchDraft,
  ) {
    const parsed = parsePresentationResearchDraft(draft),
      draftDigest = hash(canonicalPresentationValue(parsed))
    return this.update(documentId, projectId, (state) => {
      if (!Number.isSafeInteger(expectedRevision) || expectedRevision < 0)
        throw new Error('invalid_request')
      if (state.version === 2 && state.tombstones.some((t) => t.ledgerId === id))
        throw new Error('record_deleted')
      const existing = state.records.find((r) => r.id === id)
      if (existing) {
        if (existing.draftDigest !== draftDigest) throw new Error('request_conflict')
        return { record: existing, created: false }
      }
      if (state.revision !== expectedRevision) throw new Error('revision_conflict')
      if (state.totalRecords >= 128) throw new Error('quota_exceeded')
      const last = this.lastTime(state)
      const record = parsePresentationResearchRecord({
        version: 1,
        documentId,
        projectId,
        id,
        sequence: (state.version === 2 ? state.lastSequence : state.totalRecords) + 1,
        draftDigest,
        draft: parsed,
        state: 'running',
        startedAt: [new Date().toISOString(), last].sort().at(-1)!,
        checks: presentationResearchChecks,
      })
      state.records.push(record)
      state.totalRecords++
      if (state.version === 2) state.lastSequence++
      state.revision++
      return { record, created: true }
    })
  }
  finish(documentId: string, projectId: string, id: string, result: PresentationResearchFinish) {
    return this.update(documentId, projectId, (state) => {
      if (
        !result ||
        Object.keys(result).some((k) => !['state', 'sources', 'error'].includes(k)) ||
        !['completed', 'failed'].includes(result.state)
      )
        throw new Error('invalid_state')
      const index = state.records.findIndex((r) => r.id === id),
        r = state.records[index]
      if (!r)
        throw new Error(
          state.version === 2 && state.tombstones.some((t) => t.ledgerId === id)
            ? 'record_deleted'
            : 'not_found',
        )
      if (r.state !== 'running') {
        const old = {
          state: r.state,
          ...(r.sources ? { sources: r.sources } : {}),
          ...(r.error ? { error: r.error } : {}),
        }
        if (canonicalPresentationValue(old) !== canonicalPresentationValue(result))
          throw new Error('request_conflict')
        return r
      }
      const next = parsePresentationResearchRecord({
        ...r,
        ...result,
        finishedAt: [new Date().toISOString(), this.lastTime(state)].sort().at(-1)!,
      })
      state.records[index] = next
      state.revision++
      return next
    })
  }
}
