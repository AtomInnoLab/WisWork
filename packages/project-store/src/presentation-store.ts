import { parsePresentationSourceAssessment } from './presentation-source-assessment.js'
import {
  canonicalPresentationValue as canonical,
  presentationPlanSnapshotInputs,
} from './presentation-canonical.js'
import {
  parsePresentationIssueLedger,
  parsePresentationIssueActionInput,
  type PresentationIssueLedger,
  type PresentationIssueActionInput,
} from './presentation-issue.js'
import {
  parsePresentationProductionJob,
  presentationProductionJobStateAfter,
  type PresentationProductionJob,
  type PresentationProductionJobEventInput,
} from './presentation-job.js'
import { createHash, randomUUID } from 'node:crypto'
import { parsePresentationDeck } from '@wiswork/pptx-engine/presentation'
import {
  parsePresentationAssetLedger,
  type PresentationAssetLedger,
  type PresentationAssetEventInput,
} from './presentation-asset-events.js'
import {
  parsePresentationPlanAcceptances,
  type PresentationPlanAcceptanceLedger,
  type PresentationPlanAcceptance,
} from './presentation-plan-acceptance.js'
import {
  parsePresentationPlan,
  presentationSourceAttachmentId,
} from '@wiswork/pptx-engine/presentation-plan'
import {
  parsePresentationSourceAuditLedger,
  parsePresentationSourceAuditRun,
  type PresentationSourceAuditLedger,
  type PresentationSourceAuditRun,
  type PresentationSourceAuditResult,
} from './presentation-source-audit.js'
import { presentationPageInput, presentationPlanPageInput } from './presentation-page-input.js'
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
const MAX_CLAIM_REVIEW_BYTES = 8 * 1024
const MAX_CLAIM_REVIEWS_BYTES = 256 * 1024
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
function planTasks(plan: unknown): { id: string; locked?: boolean }[] {
  const slides = (plan as { slides?: unknown } | undefined)?.slides
  return Array.isArray(slides)
    ? slides.filter((page) => page && typeof page === 'object' && typeof page.id === 'string')
    : []
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
  pageLocks?: { pageId: string; productionInputDigest?: string }[]
  pageLocksDigest?: string
  revisions?: {
    revision: number
    inputDigest: string
    createdAt: string
    snapshot?: PresentationPlanRevisionSnapshot
  }[]
}
export interface PresentationPlanRevisionSnapshot {
  sourceCount: number
  claimCount: number
  slideCount: number
  sourcesDigest: string
  claimsDigest: string
  slidesDigest: string
  styleDigest: string
}
function planRevisionSnapshot(plan: unknown): PresentationPlanRevisionSnapshot | undefined {
  const value = plan as Record<string, unknown>
  if (
    !Array.isArray(value.sources) ||
    value.sources.length > 256 ||
    !Array.isArray(value.claims) ||
    value.claims.length > 256 ||
    !Array.isArray(value.slides) ||
    value.slides.length > 32 ||
    !value.style ||
    typeof value.style !== 'object' ||
    Array.isArray(value.style)
  )
    return undefined
  const inputs = presentationPlanSnapshotInputs(value)
  return {
    sourceCount: value.sources.length,
    claimCount: value.claims.length,
    slideCount: value.slides.length,
    sourcesDigest: digest(inputs.sourcesDigest),
    claimsDigest: digest(inputs.claimsDigest),
    slidesDigest: digest(inputs.slidesDigest),
    styleDigest: digest(inputs.styleDigest),
  }
}
function validPlanRevisionSnapshot(value: unknown): value is PresentationPlanRevisionSnapshot {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const snapshot = value as PresentationPlanRevisionSnapshot
  return (
    Object.keys(snapshot).sort().join(',') ===
      'claimCount,claimsDigest,slideCount,slidesDigest,sourceCount,sourcesDigest,styleDigest' &&
    [snapshot.sourceCount, snapshot.claimCount, snapshot.slideCount].every(
      (count) => Number.isSafeInteger(count) && count >= 0,
    ) &&
    snapshot.sourceCount <= 256 &&
    snapshot.claimCount <= 256 &&
    snapshot.slideCount <= 32 &&
    [
      snapshot.sourcesDigest,
      snapshot.claimsDigest,
      snapshot.slidesDigest,
      snapshot.styleDigest,
    ].every((hash) => typeof hash === 'string' && /^[a-f0-9]{64}$/.test(hash))
  )
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
  reusedFrom?: { requestId: string; inputDigest: string; planDigest: string }
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
export interface PresentationClaimReviewRecord {
  version: 1
  projectId: string
  documentId: string
  requestId: string
  reviewId: string
  inputDigest: string
  planDigest: string
  planRevision: number
  createdAt: string
  review: unknown
  reviewDigest: string
}
function claimReviewDigest(review: unknown, error: string): string {
  const assessed = Boolean(
    review && typeof review === 'object' && Object.hasOwn(review, 'sourceAssessment'),
  )
  const hash = jsonDigest(review, assessed ? 24 * 1024 : MAX_CLAIM_REVIEW_BYTES, error)
  if (!review || typeof review !== 'object' || Array.isArray(review)) throw new Error(error)
  const value = review as Record<string, unknown>
  const fields = [
    'pageId',
    'claimId',
    'sourceId',
    'attachmentId',
    'offset',
    'maxChars',
    'evidenceDigest',
    'outcome',
    'notes',
    'reviewer',
    ...(assessed ? ['sourceAssessment'] : []),
  ]
  if (
    Object.keys(value).length !== fields.length ||
    Object.keys(value).some((key) => !fields.includes(key)) ||
    ['pageId', 'claimId', 'sourceId', 'attachmentId'].some(
      (key) =>
        typeof value[key] !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(value[key] as string),
    ) ||
    !Number.isSafeInteger(value.offset) ||
    Number(value.offset) < 0 ||
    Number(value.offset) > 1000000 ||
    !Number.isSafeInteger(value.maxChars) ||
    Number(value.maxChars) < 1 ||
    Number(value.maxChars) > 8000 ||
    typeof value.evidenceDigest !== 'string' ||
    !/^[a-f0-9]{64}$/.test(value.evidenceDigest) ||
    !['supported', 'contradicted', 'insufficient_evidence'].includes(value.outcome as string) ||
    typeof value.notes !== 'string' ||
    !value.notes.trim() ||
    value.notes.length > 2000 ||
    // XML text permits tab, newline and carriage return; other controls are rejected.
    // eslint-disable-next-line no-control-regex
    /[^\u0009\u000a\u000d\u0020-\ud7ff\ue000-\ufffd\u{10000}-\u{10ffff}]/u.test(value.notes) ||
    value.reviewer !== 'agent'
  )
    throw new Error(error)
  if (assessed) {
    try {
      parsePresentationSourceAssessment(value.sourceAssessment)
    } catch {
      throw new Error(error)
    }
  }
  return hash
}
const PRODUCTION_ERRORS = new Set([
  'compile_failed',
  'invalid_deck',
  'aborted',
  'output_too_large',
  'asset_unavailable',
  'source_unavailable',
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
      (key) =>
        !['pageId', 'state', 'attempt', 'error', 'result', 'resultDigest', 'reusedFrom'].includes(
          key,
        ),
    ) ||
    !Number.isSafeInteger(page.attempt) ||
    page.attempt < 0 ||
    !['pending', 'building', 'compiled', 'failed'].includes(page.state)
  )
    invalid()
  if (
    page.reusedFrom !== undefined &&
    (page.state !== 'compiled' ||
      !page.reusedFrom ||
      typeof page.reusedFrom !== 'object' ||
      Array.isArray(page.reusedFrom) ||
      Object.keys(page.reusedFrom).sort().join(',') !== 'inputDigest,planDigest,requestId' ||
      typeof page.reusedFrom.requestId !== 'string' ||
      !/^[A-Za-z0-9_-]{1,128}$/.test(page.reusedFrom.requestId) ||
      !/^[a-f0-9]{64}$/.test(page.reusedFrom.inputDigest) ||
      !/^[a-f0-9]{64}$/.test(page.reusedFrom.planDigest))
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
      for (const page of record.pages) {
        if (!page.reusedFrom) continue
        const parent = records.find((value) => value.requestId === page.reusedFrom!.requestId)
        const original = parent?.pages.find((value) => value.pageId === page.pageId)
        const input = parent && presentationPageInput(parent.deck, parent.plan, page.pageId)
        const { reusedFrom: _reuse, ...payload } = page
        const { reusedFrom: _originalReuse, ...originalPayload } = original ?? {}
        if (
          !parent ||
          parent.sequence >= record.sequence ||
          parent.inputDigest !== page.reusedFrom.inputDigest ||
          parent.planDigest !== page.reusedFrom.planDigest ||
          !input ||
          input !== presentationPageInput(record.deck, record.plan, page.pageId) ||
          canonical(payload) !== canonical(originalPayload)
        )
          throw new Error('invalid_state')
      }
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
            presentationPageInput(parent.deck, parent.plan, page.pageId) ===
              presentationPageInput(record.deck, record.plan, page.pageId) &&
            canonical(page) !== canonical(parent.pages[index]),
        )
      )
        throw new Error('invalid_state')
    }
    if (new Set(records.map((record) => record.sequence)).size !== records.length)
      throw new Error('invalid_state')
    return records
  }
  issueActions(projectId: string, documentId: string, requestId: string): PresentationIssueLedger {
    assertPresentationId(requestId)
    const production = this.production(projectId, documentId, requestId)
    if (!production) throw new Error('not_found')
    const path = join(this.directory(projectId), `issue-actions-${digest(requestId)}.json`)
    if (!present(path))
      return {
        version: 1,
        projectId,
        documentId,
        requestId,
        inputDigest: production.inputDigest,
        planDigest: production.planDigest,
        revision: 0,
        actions: [],
      }
    try {
      if (lstatSync(path).size > 2 * 1024 * 1024) throw new Error('invalid_state')
      const record = this.read(path) as { ledger: unknown; checksum: unknown }
      if (
        !record ||
        Object.keys(record).length !== 2 ||
        !Object.hasOwn(record, 'ledger') ||
        record.checksum !== jsonDigest(record.ledger, 2 * 1024 * 1024, 'invalid_state')
      )
        throw new Error('invalid_state')
      const ledger = parsePresentationIssueLedger(record.ledger)
      if (
        ledger.projectId !== projectId ||
        ledger.documentId !== documentId ||
        ledger.requestId !== requestId ||
        ledger.inputDigest !== production.inputDigest ||
        ledger.planDigest !== production.planDigest
      )
        throw new Error('invalid_state')
      return ledger
    } catch {
      throw new Error('invalid_state')
    }
  }
  appendIssueAction(
    projectId: string,
    documentId: string,
    requestId: string,
    expectedRevision: number,
    input: PresentationIssueActionInput,
  ): PresentationIssueLedger {
    const action = parsePresentationIssueActionInput(input)
    if (!Number.isSafeInteger(expectedRevision) || expectedRevision < 0)
      throw new Error('invalid_request')
    const ledger = this.issueActions(projectId, documentId, requestId)
    const previous = ledger.actions.find((item) => item.actionId === action.actionId)
    if (previous) {
      const { sequence: _sequence, createdAt: _createdAt, ...content } = previous
      if (canonical(content) !== canonical(action)) throw new Error('request_conflict')
      return ledger
    }
    if (ledger.revision !== expectedRevision) throw new Error('revision_conflict')
    if (ledger.actions.length >= 128) throw new Error('quota_exceeded')
    const next = {
      ...ledger,
      revision: ledger.revision + 1,
      actions: [
        ...ledger.actions,
        { ...action, sequence: ledger.revision + 1, createdAt: new Date().toISOString() },
      ],
    }
    const path = join(this.directory(projectId), `issue-actions-${digest(requestId)}.json`)
    this.write(path, {
      ledger: next,
      checksum: jsonDigest(next, 2 * 1024 * 1024, 'invalid_request'),
    })
    return next
  }
  private claimReviews(
    path: string,
    production: PresentationProductionRecord,
  ): PresentationClaimReviewRecord[] {
    if (!present(path)) return []
    if (lstatSync(path).size > MAX_CLAIM_REVIEWS_BYTES) throw new Error('invalid_state')
    const records = this.read(path)
    jsonDigest(records, MAX_CLAIM_REVIEWS_BYTES, 'invalid_state')
    if (!Array.isArray(records) || records.length > 32) throw new Error('invalid_state')
    const ids = new Set<string>()
    for (const record of records as PresentationClaimReviewRecord[]) {
      if (!record || typeof record !== 'object' || Array.isArray(record))
        throw new Error('invalid_state')
      const { reviewDigest, ...content } = record
      if (
        Object.keys(record).length !== 11 ||
        Object.keys(record).some(
          (key) =>
            ![
              'version',
              'projectId',
              'documentId',
              'requestId',
              'reviewId',
              'inputDigest',
              'planDigest',
              'planRevision',
              'createdAt',
              'review',
              'reviewDigest',
            ].includes(key),
        ) ||
        record.version !== 1 ||
        record.projectId !== production.projectId ||
        record.documentId !== production.documentId ||
        record.requestId !== production.requestId ||
        record.inputDigest !== production.inputDigest ||
        record.planDigest !== production.planDigest ||
        record.planRevision !== production.plan.revision ||
        typeof record.reviewId !== 'string' ||
        !/^[A-Za-z0-9_-]{1,128}$/.test(record.reviewId) ||
        ids.has(record.reviewId) ||
        typeof record.createdAt !== 'string' ||
        !Number.isFinite(Date.parse(record.createdAt)) ||
        new Date(record.createdAt).toISOString() !== record.createdAt ||
        reviewDigest !== jsonDigest(content, MAX_CLAIM_REVIEWS_BYTES, 'invalid_state')
      )
        throw new Error('invalid_state')
      claimReviewDigest(record.review, 'invalid_state')
      ids.add(record.reviewId)
    }
    return records as PresentationClaimReviewRecord[]
  }
  saveClaimReview(
    projectId: string,
    documentId: string,
    requestId: string,
    reviewId: string,
    review: unknown,
  ): PresentationClaimReviewRecord {
    assertPresentationId(requestId)
    assertPresentationId(reviewId)
    const hash = claimReviewDigest(review, 'invalid_request')
    const production = this.production(projectId, documentId, requestId)
    if (!production) throw new Error('page_not_ready')
    const path = join(this.directory(projectId), `claim-reviews-${digest(requestId)}.json`)
    const records = this.claimReviews(path, production)
    const previous = records.find((record) => record.reviewId === reviewId)
    if (previous) {
      if (claimReviewDigest(previous.review, 'invalid_state') !== hash)
        throw new Error('request_conflict')
      return previous
    }
    if (records.length >= 32) throw new Error('quota_exceeded')
    const content = {
      version: 1 as const,
      projectId,
      documentId,
      requestId,
      reviewId,
      inputDigest: production.inputDigest,
      planDigest: production.planDigest,
      planRevision: production.plan.revision,
      createdAt: new Date().toISOString(),
      review,
    }
    const record: PresentationClaimReviewRecord = {
      ...content,
      reviewDigest: jsonDigest(content, MAX_CLAIM_REVIEWS_BYTES, 'invalid_request'),
    }
    const next = [...records, record]
    jsonDigest(next, MAX_CLAIM_REVIEWS_BYTES, 'quota_exceeded')
    const frozen = JSON.parse(canonical(next)) as PresentationClaimReviewRecord[]
    this.write(path, frozen)
    return frozen[frozen.length - 1]!
  }
  listClaimReviews(
    projectId: string,
    documentId: string,
    requestId: string,
  ): PresentationClaimReviewRecord[] {
    assertPresentationId(requestId)
    const production = this.production(projectId, documentId, requestId)
    if (!production) return []
    const path = join(this.directory(projectId), `claim-reviews-${digest(requestId)}.json`)
    return structuredClone(this.claimReviews(path, production))
  }
  claimReview(
    projectId: string,
    documentId: string,
    requestId: string,
    reviewId: string,
  ): PresentationClaimReviewRecord | undefined {
    assertPresentationId(requestId)
    assertPresentationId(reviewId)
    const production = this.production(projectId, documentId, requestId)
    if (!production) return undefined
    const path = join(this.directory(projectId), `claim-reviews-${digest(requestId)}.json`)
    return this.claimReviews(path, production).find((record) => record.reviewId === reviewId)
  }
  productionJob(
    projectId: string,
    documentId: string,
    requestId: string,
  ): PresentationProductionJob | undefined {
    const production = this.production(projectId, documentId, requestId)
    if (!production) return undefined
    const path = join(this.directory(projectId), `production-job-${digest(requestId)}.json`)
    if (!present(path)) return undefined
    if (lstatSync(path).size > 256 * 1024) throw new Error('invalid_state')
    const record = this.read(path) as { job?: unknown; checksum?: unknown } | null
    if (
      !record ||
      Object.keys(record).length !== 2 ||
      !Object.hasOwn(record, 'job') ||
      record.checksum !== jsonDigest(record.job, 256 * 1024, 'invalid_state')
    )
      throw new Error('invalid_state')
    const job = parsePresentationProductionJob(record.job)
    if (
      job.projectId !== projectId ||
      job.documentId !== documentId ||
      job.requestId !== requestId ||
      job.inputDigest !== production.inputDigest ||
      job.planDigest !== production.planDigest ||
      job.planRevision !== production.plan.revision
    )
      throw new Error('invalid_state')
    for (const event of job.events) {
      if (!('pageId' in event)) continue
      const page = production.pages.find((page) => page.pageId === event.pageId)
      if (
        !page ||
        event.attempt > page.attempt ||
        (event.type === 'page.compiled' &&
          (page.state !== 'compiled' || page.attempt !== event.attempt)) ||
        (event.type === 'page.failed' &&
          event.attempt === page.attempt &&
          (page.state !== 'failed' || page.error !== event.error))
      )
        throw new Error('invalid_state')
    }
    if (job.state === 'completed' && production.pages.some((page) => page.state !== 'compiled'))
      throw new Error('invalid_state')
    return job
  }
  appendProductionJobEvent(
    projectId: string,
    documentId: string,
    requestId: string,
    expectedRevision: number,
    event: PresentationProductionJobEventInput,
  ): PresentationProductionJob {
    if (!Number.isSafeInteger(expectedRevision) || expectedRevision < 0)
      throw new Error('invalid_request')
    const production = this.production(projectId, documentId, requestId)
    if (!production) throw new Error('page_not_ready')
    const previous = this.productionJob(projectId, documentId, requestId)
    if (
      expectedRevision !== (previous?.revision ?? 0) ||
      expectedRevision === Number.MAX_SAFE_INTEGER
    )
      throw new Error('revision_conflict')
    jsonDigest(event, 2048, 'invalid_state')
    if (
      !event ||
      typeof event !== 'object' ||
      Object.hasOwn(event, 'sequence') ||
      Object.hasOwn(event, 'createdAt')
    )
      throw new Error('invalid_state')
    const state = presentationProductionJobStateAfter(previous?.state, event)
    const now = new Date().toISOString()
    const lastTime = previous?.events.at(-1)?.createdAt ?? now
    const job = parsePresentationProductionJob({
      version: 1,
      projectId,
      documentId,
      requestId,
      inputDigest: production.inputDigest,
      planDigest: production.planDigest,
      planRevision: production.plan.revision,
      revision: expectedRevision + 1,
      state,
      events: [
        ...(previous?.events ?? []),
        { ...event, sequence: expectedRevision + 1, createdAt: now < lastTime ? lastTime : now },
      ].slice(-128),
    })
    if ('pageId' in event) {
      const page = production.pages.find((page) => page.pageId === event.pageId)
      const wanted =
        event.type === 'page.started'
          ? 'building'
          : event.type === 'page.compiled'
            ? 'compiled'
            : 'failed'
      if (
        !page ||
        page.state !== wanted ||
        page.attempt !== event.attempt ||
        (event.type === 'page.failed' && page.error !== event.error)
      )
        throw new Error('invalid_state')
      const latest = previous?.events
        .slice()
        .reverse()
        .find((value) => 'pageId' in value && value.pageId === event.pageId)
      if (
        latest &&
        'attempt' in latest &&
        (event.type === 'page.started'
          ? event.attempt <= latest.attempt
          : latest.type !== 'page.started' || latest.attempt !== event.attempt)
      )
        throw new Error('invalid_state')
      if (event.type !== 'page.started' && !latest) throw new Error('invalid_state')
    }
    if (
      event.type === 'run.completed' &&
      production.pages.some((page) => page.state !== 'compiled')
    )
      throw new Error('invalid_state')
    this.write(join(this.directory(projectId), `production-job-${digest(requestId)}.json`), {
      job,
      checksum: jsonDigest(job, 256 * 1024, 'invalid_state'),
    })
    return job
  }
  beginProduction(
    projectId: string,
    documentId: string,
    requestId: string,
    deck: unknown,
    plan: PresentationPlanBinding,
    reuseUnchanged = false,
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
    this.assertLockedPageInputs(projectId, documentId, deck, plan)
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
    // Only opt in from the schema-validated page production path. Never infer
    // factual/visual approval or host import state from a compiled page.
    const candidates = reuseUnchanged
      ? records
          .filter((value) => value.plan.revision !== plan.revision)
          .sort((a, b) => b.sequence - a.sequence)
      : []
    if (candidates.length) {
      record.pages = record.pages.map((page) => {
        const input = presentationPageInput(deck, plan, page.pageId)
        const parent =
          input &&
          candidates.find(
            (value) =>
              value.pages.some(
                (previousPage) =>
                  previousPage.pageId === page.pageId && previousPage.state === 'compiled',
              ) && input === presentationPageInput(value.deck, value.plan, page.pageId),
          )
        const previousPage = parent && parent.pages.find((value) => value.pageId === page.pageId)!
        return parent && previousPage
          ? {
              ...previousPage,
              reusedFrom: {
                requestId: parent.requestId,
                inputDigest: parent.inputDigest,
                planDigest: parent.planDigest,
              },
            }
          : page
      })
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
    this.assertLockedPageInputs(projectId, documentId, deck, parent.plan, pageId)
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
        page.pageId === pageId ||
        presentationPageInput(parent.deck, parent.plan, page.pageId) !==
          presentationPageInput(deck, parent.plan, page.pageId)
          ? { pageId: page.pageId, state: 'pending', attempt: 0 }
          : page,
      ),
    }
    productionRecord(child, 'output_too_large')
    const frozen = JSON.parse(canonical(child)) as PresentationProductionRecord
    this.write(join(directory, `production-${digest(requestId)}.json`), frozen)
    return frozen
  }
  productionHistory(projectId: string, documentId: string): PresentationProductionRecord[] {
    const directory = this.bind(projectId, documentId, false)
    if (!directory) return []
    return structuredClone(
      this.productions(directory, projectId, documentId).sort((a, b) => b.sequence - a.sequence),
    )
  }
  production(
    projectId: string,
    documentId: string,
    requestId?: string,
  ): PresentationProductionRecord | undefined {
    if (requestId !== undefined) assertPresentationId(requestId)
    const records = this.productionHistory(projectId, documentId)
    return requestId === undefined
      ? records[0]
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
  productionAssets(
    projectId: string,
    documentId: string,
    requestId: string,
  ): PresentationAssetLedger {
    const production = this.production(projectId, documentId, requestId)
    if (!production) throw new Error('not_found')
    const path = join(this.directory(projectId), `asset-events-${digest(requestId)}.json`)
    if (!present(path))
      return parsePresentationAssetLedger({
        version: 1,
        scope: 'production_asset_resolution',
        projectId,
        documentId,
        requestId,
        inputDigest: production.inputDigest,
        planDigest: production.planDigest,
        revision: 0,
        events: [],
      })
    const value = this.read(path) as PresentationAssetLedger & { ledgerDigest: string }
    if (!value || typeof value !== 'object' || Array.isArray(value))
      throw new Error('invalid_state')
    const { ledgerDigest, ...content } = value
    const ledger = parsePresentationAssetLedger(content)
    const deck = parsePresentationDeck(production.deck)
    if (
      ledger.projectId !== projectId ||
      ledger.documentId !== documentId ||
      ledger.requestId !== requestId ||
      ledger.inputDigest !== production.inputDigest ||
      ledger.planDigest !== production.planDigest ||
      ledgerDigest !== jsonDigest(ledger, 128 * 1024, 'invalid_state') ||
      ledger.events.some((event) => {
        const page = deck.slides.find((page) => page.id === event.pageId)
        return (
          !page ||
          !page.elements.some(
            (element) => element.kind === 'image' && element.assetId === event.assetId,
          ) ||
          !deck.assets.some((asset) => asset.id === event.assetId) ||
          event.attempt >
            (production.pages.find((page) => page.pageId === event.pageId)?.attempt ?? 0)
        )
      })
    )
      throw new Error('invalid_state')
    return ledger
  }
  appendProductionAsset(
    projectId: string,
    documentId: string,
    requestId: string,
    event: PresentationAssetEventInput,
  ): PresentationAssetLedger {
    const ledger = this.productionAssets(projectId, documentId, requestId)
    const production = this.production(projectId, documentId, requestId)!
    const deck = parsePresentationDeck(production.deck)
    const page = production.pages.find((page) => page.pageId === event.pageId),
      slide = deck.slides.find((page) => page.id === event.pageId)
    if (
      !page ||
      page.state !== 'building' ||
      page.attempt !== event.attempt ||
      !slide?.elements.some(
        (element) => element.kind === 'image' && element.assetId === event.assetId,
      ) ||
      !deck.assets.some((asset) => asset.id === event.assetId)
    )
      throw new Error('invalid_state')
    const now = new Date().toISOString(),
      last = ledger.events.at(-1)?.createdAt ?? ''
    const next = parsePresentationAssetLedger({
      ...ledger,
      revision: ledger.revision + 1,
      events: [
        ...ledger.events,
        { ...event, sequence: ledger.revision + 1, createdAt: now < last ? last : now },
      ].slice(-128),
    })
    this.write(join(this.directory(projectId), `asset-events-${digest(requestId)}.json`), {
      ...next,
      ledgerDigest: jsonDigest(next, 128 * 1024, 'output_too_large'),
    })
    return next
  }
  planAcceptances(projectId: string, documentId: string): PresentationPlanAcceptanceLedger {
    const directory = this.bind(projectId, documentId, false)
    const path = directory && join(directory, 'plan-acceptances.json')
    if (!path || !present(path))
      return parsePresentationPlanAcceptances({ version: 1, projectId, documentId, records: [] })
    const value = this.read(path) as PresentationPlanAcceptanceLedger & { ledgerDigest: string }
    if (!value || typeof value !== 'object' || Array.isArray(value))
      throw new Error('invalid_state')
    const { ledgerDigest, ...content } = value
    const ledger = parsePresentationPlanAcceptances(content)
    if (
      ledger.projectId !== projectId ||
      ledger.documentId !== documentId ||
      ledgerDigest !== jsonDigest(ledger, 128 * 1024, 'invalid_state')
    )
      throw new Error('invalid_state')
    return ledger
  }
  acceptPlan(
    projectId: string,
    documentId: string,
    decisionId: string,
    expectedRevision: number,
    expectedDigest: string,
  ): PresentationPlanAcceptance {
    assertPresentationId(decisionId)
    if (
      !Number.isSafeInteger(expectedRevision) ||
      expectedRevision < 1 ||
      !/^[a-f0-9]{64}$/.test(expectedDigest)
    )
      throw new Error('invalid_request')
    const ledger = this.planAcceptances(projectId, documentId)
    const previous = ledger.records.find((record) => record.decisionId === decisionId)
    if (previous) {
      if (previous.planRevision !== expectedRevision || previous.planDigest !== expectedDigest)
        throw new Error('request_conflict')
      return previous
    }
    const saved = this.plan(projectId, documentId)
    if (!saved) throw new Error('not_found')
    if (saved.revision !== expectedRevision || saved.inputDigest !== expectedDigest)
      throw new Error('revision_conflict')
    if (ledger.records.length >= 64) throw new Error('acceptance_capacity')
    const now = new Date().toISOString(),
      last = ledger.records.at(-1)?.acceptedAt ?? ''
    const acceptance: PresentationPlanAcceptance = {
      decisionId,
      planRevision: saved.revision,
      planDigest: saved.inputDigest,
      styleDigest: digest(presentationPlanSnapshotInputs(saved.plan).styleDigest),
      acceptedAt: now < last ? last : now,
    }
    const next = parsePresentationPlanAcceptances({
      ...ledger,
      records: [...ledger.records, acceptance],
    })
    this.write(join(this.directory(projectId), 'plan-acceptances.json'), {
      ...next,
      ledgerDigest: jsonDigest(next, 128 * 1024, 'output_too_large'),
    })
    return structuredClone(acceptance)
  }
  sourceAudits(projectId: string, documentId: string): PresentationSourceAuditLedger {
    const directory = this.bind(projectId, documentId, false)
    const path = directory ? join(directory, 'source-audits.json') : undefined
    if (!path || !present(path))
      return parsePresentationSourceAuditLedger({
        version: 1,
        projectId,
        documentId,
        revision: 0,
        runs: [],
      })
    const value = this.read(path) as PresentationSourceAuditLedger & { ledgerDigest: string }
    if (!value || typeof value !== 'object' || Array.isArray(value))
      throw new Error('invalid_state')
    const { ledgerDigest, ...content } = value
    const ledger = parsePresentationSourceAuditLedger(content)
    if (
      ledger.projectId !== projectId ||
      ledger.documentId !== documentId ||
      ledgerDigest !== jsonDigest(ledger, 4 * 1024 * 1024, 'invalid_state')
    )
      throw new Error('invalid_state')
    return ledger
  }

  sourceAudit(
    projectId: string,
    documentId: string,
    auditId: string,
  ): PresentationSourceAuditRun | undefined {
    assertPresentationId(auditId)
    const current = this.sourceAudits(projectId, documentId).runs.find((run) => run.id === auditId)
    if (current) return current
    const directory = this.bind(projectId, documentId, false)
    if (!directory) return undefined
    const path = join(directory, `source-audit-${digest(auditId)}.json`)
    if (!present(path)) return undefined
    const value = this.read(path) as {
      version: number
      projectId: string
      documentId: string
      run: PresentationSourceAuditRun
      runDigest: string
    }
    if (
      !value ||
      Object.keys(value).sort().join(',') !== 'documentId,projectId,run,runDigest,version' ||
      value.version !== 1 ||
      value.projectId !== projectId ||
      value.documentId !== documentId
    )
      throw new Error('invalid_state')
    const run = parsePresentationSourceAuditRun(value.run)
    if (
      run.id !== auditId ||
      run.state === 'running' ||
      value.runDigest !== jsonDigest(run, 4 * 1024 * 1024, 'invalid_state')
    )
      throw new Error('invalid_state')
    return run
  }

  beginSourceAudit(
    projectId: string,
    documentId: string,
    auditId: string,
  ): PresentationSourceAuditRun {
    assertPresentationId(auditId)
    const ledger = this.sourceAudits(projectId, documentId)
    const previous = this.sourceAudit(projectId, documentId, auditId)
    if (previous) return previous
    const saved = this.plan(projectId, documentId)
    if (!saved) throw new Error('not_found')
    const plan = parsePresentationPlan(saved.plan)
    if (plan.projectId !== projectId) throw new Error('invalid_state')
    const previousTime = ledger.runs.reduce(
      (latest, run) => [latest, run.startedAt, run.finishedAt ?? run.startedAt].sort().at(-1)!,
      '',
    )
    const now = new Date().toISOString()
    const run = parsePresentationSourceAuditRun({
      id: auditId,
      sequence: ledger.revision + 1,
      scope: 'source_excerpt_audit',
      planRevision: saved.revision,
      planDigest: saved.inputDigest,
      state: 'running',
      startedAt: now < previousTime ? previousTime : now,
      sourceRefs: plan.sources.flatMap((source) => {
        const attachmentId = presentationSourceAttachmentId(source)
        return attachmentId ? [{ sourceId: source.id, attachmentId }] : []
      }),
    })
    const runs = [...ledger.runs]
    if (runs.length === 32) {
      const removable = runs.findIndex((entry) => entry.state !== 'running')
      if (removable < 0) throw new Error('busy')
      const archived = runs[removable]!
      const path = join(this.directory(projectId), `source-audit-${digest(archived.id)}.json`)
      const value = {
        version: 1,
        projectId,
        documentId,
        run: archived,
        runDigest: jsonDigest(archived, 4 * 1024 * 1024, 'output_too_large'),
      }
      if (present(path)) {
        if (canonical(this.read(path)) !== canonical(value)) throw new Error('invalid_state')
      } else this.write(path, value)
      runs.splice(removable, 1)
    }
    runs.push(run)
    const next = parsePresentationSourceAuditLedger({
      ...ledger,
      revision: ledger.revision + 1,
      runs,
    })
    this.write(join(this.directory(projectId), 'source-audits.json'), {
      ...next,
      ledgerDigest: jsonDigest(next, 4 * 1024 * 1024, 'output_too_large'),
    })
    return structuredClone(run)
  }

  finishSourceAudit(
    projectId: string,
    documentId: string,
    auditId: string,
    result:
      | { sources: PresentationSourceAuditResult[] }
      | { error: NonNullable<PresentationSourceAuditRun['error']> },
  ): PresentationSourceAuditRun {
    assertPresentationId(auditId)
    const ledger = this.sourceAudits(projectId, documentId)
    const previous = this.sourceAudit(projectId, documentId, auditId)
    if (!previous) throw new Error('not_found')
    if (
      previous.state !== 'running' &&
      previous.state !== ('sources' in result ? 'completed' : 'failed')
    )
      throw new Error('request_conflict')
    const now = new Date().toISOString()
    const finished = parsePresentationSourceAuditRun({
      ...previous,
      state: 'sources' in result ? 'completed' : 'failed',
      finishedAt: previous.finishedAt ?? (now < previous.startedAt ? previous.startedAt : now),
      ...result,
    })
    if (previous.state !== 'running') {
      if (canonical(previous) !== canonical(finished)) throw new Error('request_conflict')
      return previous
    }
    const next = parsePresentationSourceAuditLedger({
      ...ledger,
      revision: ledger.revision + 1,
      runs: ledger.runs.map((run) => (run.id === auditId ? finished : run)),
    })
    this.write(join(this.directory(projectId), 'source-audits.json'), {
      ...next,
      ledgerDigest: jsonDigest(next, 4 * 1024 * 1024, 'output_too_large'),
    })
    return structuredClone(finished)
  }

  plan(
    projectId: string,
    documentId: string,
    requireCompleteHistory = false,
  ): PresentationPlanRecord | undefined {
    if (requireCompleteHistory) {
      const directory = this.directory(projectId)
      if (present(directory)) {
        const names = readdirSync(directory)
        if (!present(join(directory, 'project.json')) && names.length)
          throw new Error('invalid_state')
        if (
          !present(join(directory, 'plan.json')) &&
          names.some((name) => name.startsWith('plan-revision-'))
        )
          throw new Error('invalid_state')
      }
    }
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
    if (
      record.pageLocks !== undefined ||
      record.pageLocksDigest !== undefined ||
      planTasks(record.plan).some((page) => page.locked === true)
    ) {
      const lockedIds = planTasks(record.plan)
        .filter((page) => page.locked === true)
        .map((page) => page.id)
      if (
        !Array.isArray(record.pageLocks) ||
        record.pageLocks.length > 32 ||
        canonical(record.pageLocks.map((entry) => entry?.pageId)) !== canonical(lockedIds) ||
        record.pageLocks.some(
          (entry) =>
            !entry ||
            Object.keys(entry).sort().join(',') !==
              (entry.productionInputDigest === undefined
                ? 'pageId'
                : 'pageId,productionInputDigest') ||
            (entry.productionInputDigest !== undefined &&
              !/^[a-f0-9]{64}$/.test(entry.productionInputDigest)),
        ) ||
        record.pageLocksDigest !== jsonDigest(record.pageLocks, MAX_PLAN_BYTES, 'invalid_state')
      )
        throw new Error('invalid_state')
    }
    if (record.revisions !== undefined) {
      if (
        !Array.isArray(record.revisions) ||
        record.revisions.length < 1 ||
        record.revisions.length > 32
      )
        throw new Error('invalid_state')
      let previous = 0
      let previousTime = ''
      for (const event of record.revisions) {
        if (
          !event ||
          typeof event !== 'object' ||
          Array.isArray(event) ||
          Object.keys(event).sort().join(',') !==
            (event.snapshot === undefined
              ? 'createdAt,inputDigest,revision'
              : 'createdAt,inputDigest,revision,snapshot') ||
          !Number.isSafeInteger(event.revision) ||
          event.revision < 1 ||
          (previous > 0 && event.revision !== previous + 1) ||
          typeof event.inputDigest !== 'string' ||
          !/^[a-f0-9]{64}$/.test(event.inputDigest) ||
          (event.snapshot !== undefined && !validPlanRevisionSnapshot(event.snapshot)) ||
          typeof event.createdAt !== 'string' ||
          !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(event.createdAt) ||
          !Number.isFinite(Date.parse(event.createdAt)) ||
          new Date(event.createdAt).toISOString() !== event.createdAt ||
          event.createdAt < previousTime
        )
          throw new Error('invalid_state')
        previous = event.revision
        previousTime = event.createdAt
      }
      if (
        previous !== record.revision ||
        record.revisions.at(-1)!.inputDigest !== record.inputDigest
      )
        throw new Error('invalid_state')
      const latestSnapshot = record.revisions.at(-1)!.snapshot
      if (
        latestSnapshot &&
        canonical(latestSnapshot) !== canonical(planRevisionSnapshot(record.plan))
      )
        throw new Error('invalid_state')
    }
    return record
  }
  savePlan(
    projectId: string,
    documentId: string,
    expectedRevision: number,
    plan: unknown,
  ): PresentationPlanRecord {
    return this.commitPlan(projectId, documentId, expectedRevision, plan)
  }
  private commitPlan(
    projectId: string,
    documentId: string,
    expectedRevision: number,
    plan: unknown,
    unlockedPageId?: string,
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
    if (previous) {
      const before = planTasks(previous.plan)
      const after = planTasks(plan)
      for (const [index, page] of before.entries()) {
        if (page.locked !== true || page.id === unlockedPageId) continue
        const input = presentationPlanPageInput(previous.plan, page.id)
        if (
          !input ||
          after?.[index]?.id !== page.id ||
          after[index]?.locked !== true ||
          input !== presentationPlanPageInput(plan, page.id)
        )
          throw new Error('page_locked')
      }
    }
    const directory = this.bind(projectId, documentId, true)!
    const legacySnapshot =
      previous && !previous.revisions ? planRevisionSnapshot(previous.plan) : undefined
    // Register the observed legacy revision at migration time; do not invent earlier versions.
    const history =
      previous?.revisions ??
      (previous
        ? [
            {
              revision: previous.revision,
              inputDigest: previous.inputDigest,
              createdAt: new Date().toISOString(),
              ...(legacySnapshot ? { snapshot: legacySnapshot } : {}),
            },
          ]
        : [])
    const previousTime = history.at(-1)?.createdAt
    const snapshot = planRevisionSnapshot(plan)
    const pageLocks = this.planLockInputs(projectId, documentId, plan, previous)
    const record: PresentationPlanRecord = {
      version: 1,
      projectId,
      documentId,
      revision: revision + 1,
      plan,
      inputDigest,
      ...(pageLocks.length
        ? { pageLocks, pageLocksDigest: jsonDigest(pageLocks, MAX_PLAN_BYTES, 'invalid_plan') }
        : {}),
      revisions: [
        ...history,
        {
          revision: revision + 1,
          inputDigest,
          createdAt: new Date(
            Math.max(Date.now(), previousTime ? Date.parse(previousTime) : 0),
          ).toISOString(),
          ...(snapshot ? { snapshot } : {}),
        },
      ].slice(-32),
    }
    if (previous) {
      const archive = {
        version: 1,
        projectId,
        documentId,
        revision: previous.revision,
        inputDigest: previous.inputDigest,
        plan: previous.plan,
      }
      const path = join(directory, `plan-revision-${previous.revision}.json`)
      if (present(path)) {
        if (canonical(this.read(path)) !== canonical(archive)) throw new Error('invalid_state')
      } else this.write(path, archive)
    }
    this.write(join(directory, 'plan.json'), record)
    const earliest = record.revisions![0]!.revision
    for (const file of readdirSync(directory)) {
      const match = /^plan-revision-([1-9]\d*)\.json$/.exec(file)
      if (match && Number(match[1]) < earliest) rmSync(join(directory, file))
    }
    return record
  }
  setPlanPageLock(
    projectId: string,
    documentId: string,
    expectedRevision: number,
    pageId: string,
    locked: boolean,
  ): PresentationPlanRecord {
    assertPresentationId(pageId)
    if (
      !Number.isSafeInteger(expectedRevision) ||
      expectedRevision < 1 ||
      typeof locked !== 'boolean'
    )
      throw new Error('invalid_request')
    const current = this.plan(projectId, documentId)
    if (!current) throw new Error('not_found')
    if (current.revision !== expectedRevision) throw new Error('revision_conflict')
    const plan = structuredClone(current.plan) as { slides?: { id: string; locked?: boolean }[] }
    const page = plan.slides?.find((value) => value.id === pageId)
    if (!page || !presentationPlanPageInput(plan, pageId)) throw new Error('invalid_request')
    if ((page.locked === true) === locked) return current
    if (locked) page.locked = true
    else delete page.locked
    return this.commitPlan(
      projectId,
      documentId,
      expectedRevision,
      plan,
      locked ? undefined : pageId,
    )
  }
  planRevision(
    projectId: string,
    documentId: string,
    revision: number,
  ): PresentationPlanRecord | undefined {
    if (!Number.isSafeInteger(revision) || revision < 1) throw new Error('invalid_request')
    const current = this.plan(projectId, documentId)
    if (!current) return undefined
    if (current.revision === revision) return current
    const event = current.revisions?.find((value) => value.revision === revision)
    if (!event) return undefined
    const path = join(this.directory(projectId), `plan-revision-${revision}.json`)
    if (present(path)) {
      const record = this.read(path) as PresentationPlanRecord | null
      if (
        !record ||
        Array.isArray(record) ||
        Object.keys(record).sort().join(',') !==
          'documentId,inputDigest,plan,projectId,revision,version' ||
        record.version !== 1 ||
        record.projectId !== projectId ||
        record.documentId !== documentId ||
        record.revision !== revision ||
        record.inputDigest !== event.inputDigest ||
        planDigest(record.plan, 'invalid_state') !== event.inputDigest
      )
        throw new Error('invalid_state')
      return record
    }
    // Pre-archive installations may still have an exact full plan frozen in a
    // production/compile request. Never substitute the current plan or guesses.
    const bindings = [
      ...this.productionHistory(projectId, documentId).map((record) => record.plan),
      ...this.history(projectId, documentId).flatMap((record) =>
        record.plan ? [record.plan] : [],
      ),
    ]
    const binding = bindings.find(
      (value) =>
        value.revision === revision &&
        planDigest(value.plan, 'invalid_state') === event.inputDigest,
    )
    return binding
      ? {
          version: 1,
          projectId,
          documentId,
          revision,
          inputDigest: event.inputDigest,
          plan: binding.plan,
        }
      : undefined
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
    this.assertLockedPageInputs(projectId, documentId, deck, planBinding)
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
  private planLockInputs(
    projectId: string,
    documentId: string,
    plan: unknown,
    previous?: PresentationPlanRecord,
  ): NonNullable<PresentationPlanRecord['pageLocks']> {
    const tasks = planTasks(plan)
    const locked = tasks.filter((page) => page.locked === true)
    if (!locked.length) return []
    const candidates = [
      ...this.productionHistory(projectId, documentId),
      ...this.history(projectId, documentId).filter((record) => record.plan),
    ]
    return locked.map((page) => {
      const retained = previous?.pageLocks?.find((entry) => entry.pageId === page.id)
      if (retained) return { ...retained }
      const expected = presentationPlanPageInput(plan, page.id)
      if (!expected) throw new Error('page_locked')
      const baseline = candidates.find(
        (record) =>
          record.plan && presentationPlanPageInput(record.plan.plan, page.id) === expected,
      )
      const input = baseline && presentationPageInput(baseline.deck, baseline.plan!, page.id)
      if (baseline && !input) throw new Error('page_locked')
      return { pageId: page.id, ...(input ? { productionInputDigest: digest(input) } : {}) }
    })
  }
  private assertLockedPageInputs(
    projectId: string,
    documentId: string,
    deck: unknown,
    binding?: PresentationPlanBinding,
    rebuildingPageId?: string,
  ): void {
    const current = this.plan(projectId, documentId)
    const tasks = planTasks(current?.plan)
    const locked = tasks.filter((page) => page.locked === true)
    if (!locked.length) return
    if (!binding) throw new Error('page_locked')
    const pageLocks = this.planLockInputs(projectId, documentId, current!.plan, current)
    for (const page of locked) {
      if (page.id === rebuildingPageId) throw new Error('page_locked')
      const expected = presentationPlanPageInput(current!.plan, page.id)
      const boundTasks = (binding.plan as { slides?: { id: string }[] }).slides
      if (
        !expected ||
        boundTasks?.[tasks!.findIndex((task) => task.id === page.id)]?.id !== page.id ||
        presentationPlanPageInput(binding.plan, page.id) !== expected
      )
        throw new Error('page_locked')
      const input = presentationPageInput(deck, binding, page.id)
      if (!input) throw new Error('page_locked')
      const lock = pageLocks.find((entry) => entry.pageId === page.id)!
      if (lock.productionInputDigest !== undefined && lock.productionInputDigest !== digest(input))
        throw new Error('page_locked')
      lock.productionInputDigest = digest(input)
    }
    if (canonical(current!.pageLocks ?? []) !== canonical(pageLocks))
      this.write(join(this.directory(projectId), 'plan.json'), {
        ...current,
        pageLocks,
        pageLocksDigest: jsonDigest(pageLocks, MAX_PLAN_BYTES, 'invalid_state'),
      })
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
  history(projectId: string, documentId: string, includeAll = false): PresentationReceipt[] {
    const directory = this.bind(projectId, documentId, false)
    if (!directory) return []
    const records = this.receipts(directory, projectId, documentId).sort(
      (a, b) => b.sequence - a.sequence,
    )
    return includeAll ? records : records.slice(0, 20)
  }
  latest(projectId: string, documentId: string): PresentationReceipt | undefined {
    const directory = this.bind(projectId, documentId, false)
    if (!directory) return undefined
    return this.receipts(directory, projectId, documentId)
      .filter((record) => record.status === 'compiled')
      .sort((a, b) => b.sequence - a.sequence)[0]
  }
}
