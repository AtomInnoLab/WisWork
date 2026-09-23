import { createHash, randomUUID } from 'node:crypto'
import {
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
const MAX_PLAN_BYTES = 192 * 1024
function present(path: string): boolean {
  return lstatSync(path, { throwIfNoEntry: false }) !== undefined
}
function planDigest(plan: unknown, error = 'invalid_plan'): string {
  function validate(value: unknown, depth: number): void {
    if (depth > 64) throw new Error(error)
    if (value === null || typeof value === 'string' || typeof value === 'boolean') return
    if (typeof value === 'number' && Number.isFinite(value)) return
    if (
      !value ||
      typeof value !== 'object' ||
      (!Array.isArray(value) && Object.getPrototypeOf(value) !== Object.prototype)
    )
      throw new Error(error)
    const keys = Object.keys(value)
    if (
      Array.isArray(value) &&
      (keys.length !== value.length || keys.some((key, index) => key !== String(index)))
    )
      throw new Error(error)
    for (const key of keys) {
      const descriptor = Object.getOwnPropertyDescriptor(value, key)!
      if (!('value' in descriptor) || ['__proto__', 'constructor', 'prototype'].includes(key))
        throw new Error(error)
      validate(descriptor.value, depth + 1)
    }
  }
  if (!plan || typeof plan !== 'object' || Array.isArray(plan)) throw new Error(error)
  validate(plan, 0)
  const json = canonical(plan)
  if (Buffer.byteLength(json) > MAX_PLAN_BYTES) throw new Error(error)
  return digest(json)
}
export interface PresentationPlanBinding {
  revision: number
  plan: unknown
}
export interface PresentationPlanRecord extends PresentationPlanBinding {
  version: 1
  projectId: string
  documentId: string
  inputDigest: string
}
function bindingDigest(binding: PresentationPlanBinding, error = 'invalid_plan'): string {
  if (!binding || !Number.isSafeInteger(binding.revision) || binding.revision < 1)
    throw new Error(error)
  planDigest(binding.plan, error)
  return digest(canonical(binding))
}
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
  plan?: PresentationPlanBinding
  planDigest?: string
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
      if (present(path) && (!lstatSync(path).isDirectory() || lstatSync(path).isSymbolicLink()))
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
    if (!present(path)) {
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
          (record.plan === undefined
            ? record.planDigest !== undefined
            : record.planDigest !== bindingDigest(record.plan, 'invalid_state')) ||
          (record.status === 'compiled' &&
            (record.result === undefined ||
              record.resultDigest !== digest(canonical(record.result))))
        )
          throw new Error('invalid_state')
        return record
      })
  }
  plan(projectId: string, documentId: string): PresentationPlanRecord | undefined {
    const directory = this.bind(projectId, documentId, false)
    if (!directory) return undefined
    const path = join(directory, 'plan.json')
    if (!present(path)) return undefined
    const record = this.read(path) as PresentationPlanRecord | null
    if (
      !record ||
      record.version !== 1 ||
      record.projectId !== projectId ||
      record.documentId !== documentId ||
      !Number.isSafeInteger(record.revision) ||
      record.revision < 1 ||
      record.inputDigest !== planDigest(record.plan, 'invalid_state')
    )
      throw new Error('invalid_state')
    return record
  }
  savePlan(
    projectId: string,
    documentId: string,
    expectedRevision: number,
    plan: unknown,
  ): PresentationPlanRecord {
    if (!Number.isSafeInteger(expectedRevision) || expectedRevision < 0)
      throw new Error('invalid_request')
    const inputDigest = planDigest(plan)
    const previous = this.plan(projectId, documentId)
    const revision = previous?.revision ?? 0
    if (
      previous &&
      previous.inputDigest === inputDigest &&
      (expectedRevision === revision || expectedRevision === revision - 1)
    )
      return previous
    if (expectedRevision !== revision || revision === Number.MAX_SAFE_INTEGER)
      throw new Error('revision_conflict')
    const directory = this.bind(projectId, documentId, true)!
    const record: PresentationPlanRecord = {
      version: 1,
      projectId,
      documentId,
      revision: revision + 1,
      plan,
      inputDigest,
    }
    this.write(join(directory, 'plan.json'), record)
    return record
  }
  begin(
    projectId: string,
    documentId: string,
    requestId: string,
    deck: unknown,
    planBinding?: PresentationPlanBinding,
  ): PresentationReceipt {
    assertPresentationId(requestId)
    const planHash = planBinding === undefined ? undefined : bindingDigest(planBinding)
    const directory = this.bind(projectId, documentId, true)!
    const records = this.receipts(directory, projectId, documentId)
    const inputDigest = digest(canonical(deck))
    const previous = records.find((record) => record.requestId === requestId)
    if (previous) {
      if (previous.inputDigest !== inputDigest || previous.planDigest !== planHash)
        throw new Error('request_conflict')
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
      ...(planBinding === undefined ? {} : { plan: planBinding, planDigest: planHash! }),
    }
    this.write(join(directory, `${digest(requestId)}.json`), record)
    return record
  }
  complete(record: PresentationReceipt, result: unknown): void {
    const existing = this.begin(
      record.projectId,
      record.documentId,
      record.requestId,
      record.deck,
      record.plan,
    )
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
