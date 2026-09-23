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
function jsonDigest(plan: unknown, limit: number, error: string): string {
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
  validate(plan, 0)
  const json = canonical(plan)
  if (Buffer.byteLength(json) > limit) throw new Error(error)
  return digest(json)
}
function planDigest(plan: unknown, error = 'invalid_plan'): string {
  if (!plan || typeof plan !== 'object' || Array.isArray(plan)) throw new Error(error)
  return jsonDigest(plan, MAX_PLAN_BYTES, error)
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
export interface PresentationProductionPage {
  pageId: string
  state: 'pending' | 'building' | 'compiled' | 'failed'
  attempt: number
  error?: string
  result?: { pptxBase64: string; sourceSlideId: string; report: unknown }
  resultDigest?: string
}
export interface PresentationProductionRecord {
  version: 1
  projectId: string
  documentId: string
  requestId: string
  sequence: number
  inputDigest: string
  deck: unknown
  plan: PresentationPlanBinding
  planDigest: string
  pages: PresentationProductionPage[]
  revision?: { parentRequestId: string; pageId: string; parentInputDigest: string }
}
const PRODUCTION_ERRORS = new Set([
  'compile_failed',
  'invalid_deck',
  'aborted',
  'output_too_large',
  'asset_unavailable',
])
function productionIds(deck: unknown, error: string): string[] {
  jsonDigest(deck, MAX_RECORD_BYTES, error)
  if (!deck || typeof deck !== 'object' || Array.isArray(deck)) throw new Error(error)
  const slides = (deck as { slides?: unknown }).slides
  if (!Array.isArray(slides) || slides.length < 1 || slides.length > 32) throw new Error(error)
  const ids = slides.map((slide) => {
    const id = (slide as { id?: unknown } | null)?.id
    if (typeof id !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(id)) throw new Error(error)
    return id
  })
  if (new Set(ids).size !== ids.length) throw new Error(error)
  return ids
}
function productionPage(page: PresentationProductionPage): number {
  const invalid = () => {
    throw new Error('invalid_state')
  }
  if (
    !page ||
    typeof page !== 'object' ||
    Object.keys(page).some(
      (key) => !['pageId', 'state', 'attempt', 'error', 'result', 'resultDigest'].includes(key),
    ) ||
    !Number.isSafeInteger(page.attempt) ||
    page.attempt < 0 ||
    !['pending', 'building', 'compiled', 'failed'].includes(page.state)
  )
    invalid()
  if (page.state === 'pending' ? page.attempt !== 0 : page.attempt < 1) invalid()
  if (
    page.state === 'failed'
      ? !PRODUCTION_ERRORS.has(page.error ?? '')
      : Object.hasOwn(page, 'error')
  )
    invalid()
  if (page.state !== 'compiled') {
    if (Object.hasOwn(page, 'result') || Object.hasOwn(page, 'resultDigest')) invalid()
    return 0
  }
  const result = page.result
  if (
    !result ||
    typeof result !== 'object' ||
    Object.keys(result).some((key) => !['pptxBase64', 'sourceSlideId', 'report'].includes(key)) ||
    typeof result.pptxBase64 !== 'string' ||
    !result.pptxBase64.length ||
    result.pptxBase64.length > Math.ceil((10 * 1024 * 1024) / 3) * 4 ||
    typeof result.sourceSlideId !== 'string' ||
    !/^[1-9]\d*#$/.test(result.sourceSlideId) ||
    !Number.isSafeInteger(Number(result.sourceSlideId.slice(0, -1))) ||
    Number(result.sourceSlideId.slice(0, -1)) < 256 ||
    Number(result.sourceSlideId.slice(0, -1)) > 0xffffffff
  )
    invalid()
  const bytes = Buffer.from(result!.pptxBase64, 'base64')
  if (
    !bytes.length ||
    bytes.length > 10 * 1024 * 1024 ||
    bytes.toString('base64') !== result!.pptxBase64
  )
    invalid()
  jsonDigest(result!.report, MAX_PLAN_BYTES, 'invalid_state')
  if (page.resultDigest !== jsonDigest(result, MAX_RECORD_BYTES, 'invalid_state')) invalid()
  return bytes.length
}
function productionRecord(
  record: PresentationProductionRecord,
  budgetError = 'invalid_state',
): void {
  if (
    !record ||
    typeof record !== 'object' ||
    Object.keys(record).some(
      (key) =>
        ![
          'version',
          'projectId',
          'documentId',
          'requestId',
          'sequence',
          'inputDigest',
          'deck',
          'plan',
          'planDigest',
          'pages',
          'revision',
        ].includes(key),
    ) ||
    record.version !== 1 ||
    typeof record.projectId !== 'string' ||
    !/^[A-Za-z0-9_-]{1,128}$/.test(record.projectId) ||
    typeof record.requestId !== 'string' ||
    !/^[A-Za-z0-9_-]{1,128}$/.test(record.requestId) ||
    typeof record.documentId !== 'string' ||
    !record.documentId.trim() ||
    record.documentId.length > 2048 ||
    !Number.isSafeInteger(record.sequence) ||
    record.sequence < 1
  )
    throw new Error('invalid_state')
  const revision = record.revision
  if (
    revision !== undefined &&
    (!revision ||
      typeof revision !== 'object' ||
      Array.isArray(revision) ||
      Object.keys(revision).length !== 3 ||
      Object.keys(revision).some(
        (key) => !['parentRequestId', 'pageId', 'parentInputDigest'].includes(key),
      ) ||
      typeof revision.parentRequestId !== 'string' ||
      !/^[A-Za-z0-9_-]{1,128}$/.test(revision.parentRequestId) ||
      revision.parentRequestId === record.requestId ||
      typeof revision.pageId !== 'string' ||
      !/^[A-Za-z0-9_-]{1,128}$/.test(revision.pageId) ||
      typeof revision.parentInputDigest !== 'string' ||
      !/^[a-f0-9]{64}$/.test(revision.parentInputDigest))
  )
    throw new Error('invalid_state')
  const ids = productionIds(record.deck, 'invalid_state')
  if (revision && !ids.includes(revision.pageId)) throw new Error('invalid_state')
  if (
    record.inputDigest !== jsonDigest(record.deck, MAX_RECORD_BYTES, 'invalid_state') ||
    !record.plan ||
    Object.keys(record.plan).some((key) => !['revision', 'plan'].includes(key)) ||
    record.planDigest !== bindingDigest(record.plan, 'invalid_state') ||
    !Array.isArray(record.pages) ||
    record.pages.length !== ids.length
  )
    throw new Error('invalid_state')
  let total = 0
  record.pages.forEach((page, index) => {
    if (!page || page.pageId !== ids[index]) throw new Error('invalid_state')
    total += productionPage(page)
  })
  if (total > 10 * 1024 * 1024) throw new Error(budgetError)
  jsonDigest(record, MAX_RECORD_BYTES, budgetError)
}

function sameUnchangedPages(parent: unknown, child: unknown, pageId: string): boolean {
  const { slides: original, ...parentShared } = parent as { slides: Array<{ id: string }> }
  const { slides: revised, ...childShared } = child as { slides: Array<{ id: string }> }
  return (
    canonical(parentShared) === canonical(childShared) &&
    original.length === revised.length &&
    original.every(
      (slide, index) =>
        slide.id === revised[index]?.id &&
        (slide.id === pageId || canonical(slide) === canonical(revised[index])),
    )
  )
}

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
  private productions(
    directory: string,
    projectId: string,
    documentId: string,
  ): PresentationProductionRecord[] {
    const files = readdirSync(directory).filter((name) =>
      /^production-[a-f0-9]{64}\.json$/.test(name),
    )
    if (files.length > 32) throw new Error('invalid_state')
    const records = files.map((name) => {
      const record = this.read(join(directory, name)) as PresentationProductionRecord
      productionRecord(record)
      if (
        record.projectId !== projectId ||
        record.documentId !== documentId ||
        name !== `production-${digest(record.requestId)}.json`
      )
        throw new Error('invalid_state')
      return record
    })
    for (const record of records) {
      if (!record.revision) continue
      const parent = records.find((value) => value.requestId === record.revision!.parentRequestId)
      if (
        !parent ||
        parent.sequence >= record.sequence ||
        parent.inputDigest !== record.revision.parentInputDigest ||
        parent.planDigest !== record.planDigest ||
        parent.pages.some((page) => page.state !== 'compiled') ||
        !sameUnchangedPages(parent.deck, record.deck, record.revision.pageId) ||
        record.pages.some(
          (page, index) =>
            page.pageId !== record.revision!.pageId &&
            canonical(page) !== canonical(parent.pages[index]),
        )
      )
        throw new Error('invalid_state')
    }
    if (new Set(records.map((record) => record.sequence)).size !== records.length)
      throw new Error('invalid_state')
    return records
  }
  beginProduction(
    projectId: string,
    documentId: string,
    requestId: string,
    deck: unknown,
    plan: PresentationPlanBinding,
  ): PresentationProductionRecord {
    assertPresentationId(requestId)
    const ids = productionIds(deck, 'invalid_request')
    if (!plan || Object.keys(plan).some((key) => !['revision', 'plan'].includes(key)))
      throw new Error('invalid_plan')
    const planHash = bindingDigest(plan)
    const inputDigest = jsonDigest(deck, MAX_RECORD_BYTES, 'invalid_request')
    const directory = this.bind(projectId, documentId, true)!
    const records = this.productions(directory, projectId, documentId)
    const previous = records.find((record) => record.requestId === requestId)
    if (previous) {
      if (
        previous.revision ||
        previous.inputDigest !== inputDigest ||
        previous.planDigest !== planHash
      )
        throw new Error('request_conflict')
      return previous
    }
    if (records.length >= 32) throw new Error('output_too_large')
    const record: PresentationProductionRecord = {
      version: 1,
      projectId,
      documentId,
      requestId,
      sequence: Math.max(0, ...records.map((record) => record.sequence)) + 1,
      inputDigest,
      deck,
      plan,
      planDigest: planHash,
      pages: ids.map((pageId) => ({ pageId, state: 'pending', attempt: 0 })),
    }
    productionRecord(record, 'output_too_large')
    const frozen = JSON.parse(canonical(record)) as PresentationProductionRecord
    this.write(join(directory, `production-${digest(requestId)}.json`), frozen)
    return frozen
  }
  deriveProduction(
    projectId: string,
    documentId: string,
    parentRequestId: string,
    requestId: string,
    pageId: string,
    deck: unknown,
  ): PresentationProductionRecord {
    assertPresentationId(parentRequestId)
    assertPresentationId(requestId)
    assertPresentationId(pageId)
    if (parentRequestId === requestId) throw new Error('invalid_request')
    const ids = productionIds(deck, 'invalid_request')
    if (!ids.includes(pageId)) throw new Error('invalid_request')
    const directory = this.bind(projectId, documentId, false)
    if (!directory) throw new Error('page_not_ready')
    const records = this.productions(directory, projectId, documentId)
    const parent = records.find((record) => record.requestId === parentRequestId)
    if (!parent || parent.pages.some((page) => page.state !== 'compiled'))
      throw new Error('page_not_ready')
    const inputDigest = jsonDigest(deck, MAX_RECORD_BYTES, 'invalid_request')
    const revision = { parentRequestId, pageId, parentInputDigest: parent.inputDigest }
    const previous = records.find((record) => record.requestId === requestId)
    if (previous) {
      if (
        previous.inputDigest !== inputDigest ||
        canonical(previous.revision) !== canonical(revision)
      )
        throw new Error('request_conflict')
      return previous
    }
    if (!sameUnchangedPages(parent.deck, deck, pageId)) throw new Error('invalid_request')
    if (records.length >= 32) throw new Error('output_too_large')
    const child: PresentationProductionRecord = {
      version: 1,
      projectId,
      documentId,
      requestId,
      sequence: Math.max(...records.map((record) => record.sequence)) + 1,
      inputDigest,
      deck,
      plan: parent.plan,
      planDigest: parent.planDigest,
      revision,
      pages: parent.pages.map((page) =>
        page.pageId === pageId ? { pageId, state: 'pending', attempt: 0 } : page,
      ),
    }
    productionRecord(child, 'output_too_large')
    const frozen = JSON.parse(canonical(child)) as PresentationProductionRecord
    this.write(join(directory, `production-${digest(requestId)}.json`), frozen)
    return frozen
  }
  production(
    projectId: string,
    documentId: string,
    requestId?: string,
  ): PresentationProductionRecord | undefined {
    if (requestId !== undefined) assertPresentationId(requestId)
    const directory = this.bind(projectId, documentId, false)
    if (!directory) return undefined
    const records = this.productions(directory, projectId, documentId)
    return requestId === undefined
      ? records.sort((a, b) => b.sequence - a.sequence)[0]
      : records.find((record) => record.requestId === requestId)
  }
  updateProductionPage(
    record: PresentationProductionRecord,
    pageId: string,
    update: Pick<PresentationProductionPage, 'state' | 'attempt' | 'error' | 'result'>,
  ): PresentationProductionRecord {
    productionRecord(record)
    const current = this.production(record.projectId, record.documentId, record.requestId)
    if (
      !current ||
      current.sequence !== record.sequence ||
      current.inputDigest !== record.inputDigest ||
      current.planDigest !== record.planDigest ||
      canonical(current.revision) !== canonical(record.revision)
    )
      throw new Error('invalid_state')
    const index = current.pages.findIndex((page) => page.pageId === pageId)
    if (index < 0 || !record.pages[index] || record.pages[index]!.pageId !== pageId)
      throw new Error('invalid_state')
    if (canonical(current.pages[index]) !== canonical(record.pages[index]))
      throw new Error('revision_conflict')
    if (
      !update ||
      typeof update !== 'object' ||
      Object.keys(update).some((key) => !['state', 'attempt', 'error', 'result'].includes(key))
    )
      throw new Error('invalid_state')
    const before = current.pages[index]!
    const building =
      update.state === 'building' &&
      ['pending', 'building', 'failed'].includes(before.state) &&
      update.attempt === before.attempt + 1
    const finishing =
      before.state === 'building' &&
      ['compiled', 'failed'].includes(update.state) &&
      update.attempt === before.attempt
    if (!building && !finishing) throw new Error('invalid_state')
    const nextPage: PresentationProductionPage = {
      pageId,
      ...update,
      ...(update.state === 'compiled'
        ? { resultDigest: jsonDigest(update.result, MAX_RECORD_BYTES, 'invalid_state') }
        : {}),
    }
    const next = {
      ...current,
      pages: current.pages.map((page, i) => (i === index ? nextPage : page)),
    }
    productionRecord(next, 'output_too_large')
    const frozen = JSON.parse(canonical(next)) as PresentationProductionRecord
    this.write(
      join(this.directory(record.projectId), `production-${digest(record.requestId)}.json`),
      frozen,
    )
    return frozen
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
