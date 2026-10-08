import { constants } from 'node:fs'
import { createHash, randomUUID } from 'node:crypto'
import { lstat, mkdir, rename, rm, open } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import {
  parsePresentationAcquisitionHistory,
  type PresentationAcquisitionHistory,
  type PresentationAcquisitionInput,
  type PresentationAcquisitionResult,
  type PresentationAcquisitionRecord,
} from './presentation-acquisition'
const hash = (v: string) => createHash('sha256').update(v).digest('hex')
const locks = new Map<string, Promise<void>>()
export class PresentationAcquisitionStore {
  private readonly root: string
  constructor(userDataPath: string) {
    this.root = join(resolve(userDataPath), 'presentation-acquisition-history')
  }
  private async safe(path: string, directory = false): Promise<boolean> {
    try {
      const s = await lstat(path)
      if (s.isSymbolicLink() || (directory ? !s.isDirectory() : !s.isFile()))
        throw new Error('invalid_state')
      return true
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code === 'ENOENT') return false
      // eslint-disable-next-line preserve-caught-error -- Disk paths and raw errors must not cross the service boundary.
      throw new Error('invalid_state')
    }
  }
  async read(documentId: string): Promise<PresentationAcquisitionHistory> {
    const empty: PresentationAcquisitionHistory = {
      version: 1,
      scope: 'remote_material_acquisition',
      documentId,
      revision: 0,
      totalAttempts: 0,
      records: [],
    }
    parsePresentationAcquisitionHistory(empty)
    try {
      if (!(await this.safe(this.root, true))) return empty
      const path = join(this.root, hash(documentId) + '.json')
      if (!(await this.safe(path))) return empty
      const file = await open(
        path,
        constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK,
      )
      let raw: Buffer
      try {
        const stat = await file.stat()
        if (!stat.isFile() || stat.size > 256 * 1024) throw new Error('invalid_state')
        raw = await file.readFile()
        if (raw.length > 256 * 1024) throw new Error('invalid_state')
      } finally {
        await file.close()
      }
      const v = JSON.parse(raw.toString())
      if (
        Object.keys(v).sort().join(',') !== 'checksum,history' ||
        v.checksum !== hash(JSON.stringify(v.history))
      )
        throw new Error('invalid_state')
      const history = parsePresentationAcquisitionHistory(v.history)
      if (history.documentId !== documentId) throw new Error('invalid_state')
      return history
    } catch {
      throw new Error('invalid_state')
    }
  }
  private async update<T>(
    documentId: string,
    change: (h: PresentationAcquisitionHistory) => T,
  ): Promise<T> {
    const key = join(this.root, hash(documentId))
    const previous = locks.get(key) ?? Promise.resolve()
    let release!: () => void
    const tail = new Promise<void>((r) => {
      release = r
    })
    locks.set(key, tail)
    await previous
    try {
      const h = await this.read(documentId)
      const result = change(h)
      parsePresentationAcquisitionHistory(h)
      await mkdir(this.root, { recursive: true })
      await this.safe(this.root, true)
      const raw = JSON.stringify({ history: h, checksum: hash(JSON.stringify(h)) })
      if (Buffer.byteLength(raw) > 256 * 1024) throw new Error('invalid_state')
      const path = key + '.json'
      await this.safe(path)
      const tmp = join(this.root, `.tmp-${randomUUID()}`)
      try {
        const f = await open(tmp, 'wx', 0o600)
        try {
          await f.writeFile(raw)
          await f.sync()
        } finally {
          await f.close()
        }
        await rename(tmp, path)
      } finally {
        await rm(tmp, { force: true })
      }
      return structuredClone(result)
    } finally {
      release()
      if (locks.get(key) === tail) locks.delete(key)
    }
  }
  begin(
    documentId: string,
    input: PresentationAcquisitionInput,
  ): Promise<PresentationAcquisitionRecord> {
    if (!input || Object.keys(input).sort().join(',') !== 'kind,source,sourceUrlHash')
      throw new Error('invalid_state')
    return this.update(documentId, (h) => {
      const last =
        h.records
          .flatMap((r) => (r.state === 'fetching' ? [r.startedAt] : [r.startedAt, r.finishedAt]))
          .sort()
          .at(-1) ?? ''
      const record: PresentationAcquisitionRecord = {
        ...input,
        id: randomUUID(),
        attempt: ++h.totalAttempts,
        state: 'fetching',
        startedAt: [new Date().toISOString(), last].sort().at(-1)!,
      }
      h.revision++
      h.records.push(record)
      h.records = h.records.slice(-64)
      return record
    })
  }
  finish(
    documentId: string,
    id: string,
    result: PresentationAcquisitionResult,
  ): Promise<PresentationAcquisitionRecord> {
    if (
      !result ||
      Object.keys(result).some(
        (key) =>
          !['state', 'attachmentId', 'sha256', 'sizeBytes', 'assetSha256', 'error'].includes(key),
      )
    )
      throw new Error('invalid_state')
    return this.update(documentId, (h) => {
      const index = h.records.findIndex((r) => r.id === id)
      const r = h.records[index]
      if (!r) throw new Error('invalid_state')
      if (r.state !== 'fetching') {
        const same =
          Object.keys(result).length === Object.keys(r).length - 7 &&
          Object.entries(result).every(
            ([key, value]) => (r as unknown as Record<string, unknown>)[key] === value,
          )
        if (!same) throw new Error('invalid_state')
        return r
      }
      const next = {
        ...r,
        ...result,
        finishedAt: [new Date().toISOString(), r.startedAt].sort().at(-1)!,
      } as PresentationAcquisitionRecord
      h.records[index] = next
      h.revision++
      return next
    })
  }
}
