import { createHash, randomUUID } from 'node:crypto'
import { constants } from 'node:fs'
import { lstat, mkdir, open, readdir, rename, rm } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { canonicalPresentationValue } from './presentation-canonical'
import {
  parsePresentationResearchDraft,
  parsePresentationResearchRecord,
  parsePresentationResearchHistory,
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
interface State {
  version: 1
  documentId: string
  projectId: string
  revision: number
  totalRecords: number
  records: PresentationResearchRecord[]
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
          'documentId,projectId,records,revision,totalRecords,version' ||
        state.documentId !== documentId ||
        state.projectId !== projectId ||
        state.version !== 1 ||
        !Array.isArray(state.records) ||
        state.records.length !== state.totalRecords ||
        state.totalRecords > 128
      )
        throw new Error('invalid_state')
      const seen = new Set<string>()
      let completed = 0,
        last = ''
      for (const [index, r] of state.records.entries()) {
        parsePresentationResearchRecord(r)
        if (
          r.documentId !== documentId ||
          r.projectId !== projectId ||
          r.sequence !== index + 1 ||
          seen.has(r.id) ||
          r.startedAt < last ||
          r.draftDigest !== hash(canonicalPresentationValue(r.draft))
        )
          throw new Error('invalid_state')
        seen.add(r.id)
        last = r.startedAt
        if (r.state !== 'running') completed++
      }
      if (state.revision !== state.totalRecords + completed) throw new Error('invalid_state')
      this.toHistory(state)
      return state
    } catch {
      throw new Error('invalid_state')
    }
  }
  private toHistory(state: State): PresentationResearchHistory {
    return parsePresentationResearchHistory({ ...state, records: state.records.slice(-32) })
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
    const record = (await this.load(documentId, projectId)).records.find((r) => r.id === id)
    if (!record) throw new Error('not_found')
    return structuredClone(record)
  }
  private async update<T>(
    documentId: string,
    projectId: string,
    change: (state: State) => T,
  ): Promise<T> {
    const project = this.path(documentId, projectId)
    const previous = locks.get(project) ?? Promise.resolve()
    let release!: () => void
    const tail = new Promise<void>((r) => {
      release = r
    })
    locks.set(project, tail)
    await previous
    try {
      const state = await this.load(documentId, projectId)
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
      const existing = state.records.find((r) => r.id === id)
      if (existing) {
        if (existing.draftDigest !== draftDigest) throw new Error('request_conflict')
        return { record: existing, created: false }
      }
      if (state.revision !== expectedRevision) throw new Error('revision_conflict')
      if (state.totalRecords >= 128) throw new Error('quota_exceeded')
      const last =
        state.records
          .flatMap((r) => (r.finishedAt ? [r.startedAt, r.finishedAt] : [r.startedAt]))
          .sort()
          .at(-1) ?? ''
      const record = parsePresentationResearchRecord({
        version: 1,
        documentId,
        projectId,
        id,
        sequence: state.totalRecords + 1,
        draftDigest,
        draft: parsed,
        state: 'running',
        startedAt: [new Date().toISOString(), last].sort().at(-1)!,
        checks: presentationResearchChecks,
      })
      state.records.push(record)
      state.totalRecords++
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
      if (!r) throw new Error('not_found')
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
        finishedAt: [new Date().toISOString(), r.startedAt].sort().at(-1)!,
      })
      state.records[index] = next
      state.revision++
      return next
    })
  }
}
