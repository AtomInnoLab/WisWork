import { createHash, randomUUID } from 'node:crypto'
import {
  existsSync,
  lstatSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { dirname, join } from 'node:path'

const MAX_RECORD_BYTES = 17 * 1024 * 1024
export function assertPresentationId(value: unknown): asserts value is string {
  if (typeof value !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(value))
    throw new Error('invalid_request')
}
function digest(value: string): string {
  return createHash('sha256').update(value).digest('hex')
}
function canonical(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`
  if (value && typeof value === 'object')
    return `{${Object.keys(value)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${canonical((value as Record<string, unknown>)[key])}`)
      .join(',')}}`
  return JSON.stringify(value)
}
export interface PresentationReceipt {
  version: 1
  projectId: string
  documentId: string
  requestId: string
  sequence: number
  inputDigest: string
  deck: unknown
  status: 'pending' | 'compiled'
  result?: unknown
  resultDigest?: string
}

/** Each bounded receipt atomically contains its input and output; pending runs are safe to retry. */
export class PresentationStore {
  private readonly base: string
  constructor(userDataPath: string) {
    this.base = join(userDataPath, 'projects', 'presentations')
  }

  private directory(projectId: string): string {
    assertPresentationId(projectId)
    const directory = join(this.base, digest(projectId))
    for (const path of [dirname(this.base), this.base, directory]) {
      if (existsSync(path) && (!lstatSync(path).isDirectory() || lstatSync(path).isSymbolicLink()))
        throw new Error('invalid_state')
    }
    return directory
  }
  private read(path: string): unknown {
    const stat = lstatSync(path)
    if (!stat.isFile() || stat.isSymbolicLink() || stat.size > MAX_RECORD_BYTES)
      throw new Error('invalid_state')
    try {
      return JSON.parse(readFileSync(path, 'utf8'))
    } catch {
      throw new Error('invalid_state')
    }
  }
  private write(path: string, value: unknown): void {
    const json = JSON.stringify(value)
    if (Buffer.byteLength(json) > MAX_RECORD_BYTES) throw new Error('output_too_large')
    const temporary = `${path}.${randomUUID()}.tmp`
    try {
      writeFileSync(temporary, json, { encoding: 'utf8', flag: 'wx', mode: 0o600 })
      renameSync(temporary, path)
    } finally {
      rmSync(temporary, { force: true })
    }
  }
  private bind(projectId: string, documentId: string, create: boolean): string | undefined {
    if (typeof documentId !== 'string' || !documentId.trim() || documentId.length > 2048)
      throw new Error('invalid_request')
    const directory = this.directory(projectId)
    const path = join(directory, 'project.json')
    if (!existsSync(path)) {
      if (!create) return undefined
      mkdirSync(directory, { recursive: true })
      this.write(path, { version: 1, projectId, documentId })
    }
    const metadata = this.read(path) as Record<string, unknown> | null
    if (
      !metadata ||
      metadata.version !== 1 ||
      metadata.projectId !== projectId ||
      typeof metadata.documentId !== 'string'
    )
      throw new Error('invalid_state')
    if (metadata.documentId !== documentId) throw new Error('document_mismatch')
    return directory
  }
  private receipts(
    directory: string,
    projectId: string,
    documentId: string,
  ): PresentationReceipt[] {
    // ponytail: scan receipts per request; add an index if project histories make this measurable.
    return readdirSync(directory)
      .filter((name) => /^[a-f0-9]{64}\.json$/.test(name))
      .map((name) => {
        const record = this.read(join(directory, name)) as PresentationReceipt | null
        if (
          !record ||
          record.version !== 1 ||
          record.projectId !== projectId ||
          record.documentId !== documentId ||
          typeof record.requestId !== 'string' ||
          `${digest(record.requestId)}.json` !== name ||
          !Number.isSafeInteger(record.sequence) ||
          record.sequence < 1 ||
          (record.status !== 'pending' && record.status !== 'compiled') ||
          record.inputDigest !== digest(canonical(record.deck)) ||
          (record.status === 'compiled' &&
            (record.result === undefined ||
              record.resultDigest !== digest(canonical(record.result))))
        )
          throw new Error('invalid_state')
        return record
      })
  }
  begin(
    projectId: string,
    documentId: string,
    requestId: string,
    deck: unknown,
  ): PresentationReceipt {
    assertPresentationId(requestId)
    const directory = this.bind(projectId, documentId, true)!
    const records = this.receipts(directory, projectId, documentId)
    const inputDigest = digest(canonical(deck))
    const previous = records.find((record) => record.requestId === requestId)
    if (previous) {
      if (previous.inputDigest !== inputDigest) throw new Error('request_conflict')
      return previous
    }
    const record: PresentationReceipt = {
      version: 1,
      projectId,
      documentId,
      requestId,
      sequence: Math.max(0, ...records.map((record) => record.sequence)) + 1,
      inputDigest,
      deck,
      status: 'pending',
    }
    this.write(join(directory, `${digest(requestId)}.json`), record)
    return record
  }
  complete(record: PresentationReceipt, result: unknown): void {
    const existing = this.begin(record.projectId, record.documentId, record.requestId, record.deck)
    if (existing.status === 'compiled') return
    this.write(join(this.directory(record.projectId), `${digest(record.requestId)}.json`), {
      ...existing,
      status: 'compiled',
      result,
      resultDigest: digest(canonical(result)),
    })
  }
  request(
    projectId: string,
    documentId: string,
    requestId: string,
  ): PresentationReceipt | undefined {
    assertPresentationId(requestId)
    const directory = this.bind(projectId, documentId, false)
    if (!directory) return undefined
    return this.receipts(directory, projectId, documentId).find(
      (record) => record.requestId === requestId,
    )
  }
  history(projectId: string, documentId: string): PresentationReceipt[] {
    const directory = this.bind(projectId, documentId, false)
    if (!directory) return []
    return this.receipts(directory, projectId, documentId)
      .sort((a, b) => b.sequence - a.sequence)
      .slice(0, 20)
  }
  latest(projectId: string, documentId: string): PresentationReceipt | undefined {
    const directory = this.bind(projectId, documentId, false)
    if (!directory) return undefined
    return this.receipts(directory, projectId, documentId)
      .filter((record) => record.status === 'compiled')
      .sort((a, b) => b.sequence - a.sequence)[0]
  }
}
