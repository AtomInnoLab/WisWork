import {
  PROFESSIONAL_CONTEXT_SCHEMA,
  parsePresentationProfessionalContext,
  type PresentationProfessionalContext,
} from './presentation-professional-context.js'
export interface PresentationResearchSource {
  id: string
  title: string
  uri: string
  snapshotAttachmentId?: string
  excerpt: string
  locator?: string
  asOf?: string
}
export interface PresentationResearchFact {
  claimId: string
  statement: string
  type: 'fact' | 'quote' | 'calculation' | 'judgment' | 'assumption'
  sourceRefs: string[]
  sourceTier: 'primary' | 'authoritative_secondary' | 'secondary' | 'unverified'
  slideRefs: string[]
  confidence: 'high' | 'medium' | 'low'
  reviewStatus: 'needs_review'
  conflictsWith: string[]
  asOf?: string
  jurisdiction?: string
  professionalContext?: PresentationProfessionalContext
  calculation?: { formula: string; inputs: string[]; unit?: string; currency?: string }
}
export interface PresentationResearchDraft {
  scope: string
  sources: PresentationResearchSource[]
  facts: PresentationResearchFact[]
}
export interface PresentationResearchEvidence {
  sourceId: string
  attachmentId?: string
  status:
    | 'found'
    | 'not_found'
    | 'empty_excerpt'
    | 'not_ready'
    | 'unsupported'
    | 'missing'
    | 'source_mismatch'
  offset?: number
  locator?: string
  provenance: 'user_supplied' | 'fetched_url_matched' | 'unavailable'
  retrievedAt?: string
  sha256?: string
  parsedTextSha256?: string
}
export interface PresentationResearchRecord {
  version: 1
  documentId: string
  projectId: string
  id: string
  sequence: number
  draftDigest: string
  draft: PresentationResearchDraft
  state: 'running' | 'completed' | 'failed'
  startedAt: string
  finishedAt?: string
  sources?: PresentationResearchEvidence[]
  error?: 'aborted' | 'source_unavailable' | 'invalid_state'
  checks: {
    scope: 'research_draft'
    support: 'not_verified'
    sourceAuthority: 'not_verified'
    timeliness: 'not_verified'
  }
}
export interface PresentationResearchHistory {
  version: 1
  documentId: string
  projectId: string
  revision: number
  totalRecords: number
  records: PresentationResearchRecord[]
}
export type PresentationResearchSummaryRecord = Pick<
  PresentationResearchRecord,
  'id' | 'sequence' | 'draftDigest' | 'state' | 'startedAt' | 'finishedAt' | 'error'
> & { sourceCount: number; factCount: number; conflictCount: number }
export interface PresentationResearchSummary {
  version: 1
  documentId: string
  projectId: string
  revision: number
  totalRecords: number
  records: PresentationResearchSummaryRecord[]
}
export type PresentationResearchHistorySummary = PresentationResearchSummary
const textSchema = (maxLength: number, minLength = 0) => ({ type: 'string', minLength, maxLength })
const idSchema = { ...textSchema(128, 1), pattern: '^[A-Za-z0-9_-]+$' }
const arraySchema = (items: unknown, maxItems = 64, minItems = 0) => ({
  type: 'array',
  items,
  minItems,
  maxItems,
  uniqueItems: true,
})
const objectSchema = (properties: Record<string, unknown>, required = Object.keys(properties)) => ({
  type: 'object',
  properties,
  required,
  additionalProperties: false,
})
const choiceSchema = (...values: string[]) => ({ type: 'string', enum: values })
export const PRESENTATION_RESEARCH_DRAFT_SCHEMA = objectSchema({
  scope: textSchema(4000, 1),
  sources: arraySchema(
    objectSchema(
      {
        id: idSchema,
        title: textSchema(300, 1),
        uri: textSchema(500, 1),
        snapshotAttachmentId: { ...textSchema(64, 64), pattern: '^[a-f0-9]{64}$' },
        excerpt: textSchema(12000),
        locator: textSchema(200),
        asOf: textSchema(100, 1),
      },
      ['id', 'title', 'uri', 'excerpt'],
    ),
  ),
  facts: arraySchema(
    objectSchema(
      {
        claimId: idSchema,
        statement: textSchema(12000, 1),
        type: choiceSchema('fact', 'quote', 'calculation', 'judgment', 'assumption'),
        sourceRefs: arraySchema(idSchema),
        sourceTier: choiceSchema('primary', 'authoritative_secondary', 'secondary', 'unverified'),
        slideRefs: arraySchema(idSchema),
        confidence: choiceSchema('high', 'medium', 'low'),
        reviewStatus: choiceSchema('needs_review'),
        conflictsWith: arraySchema(idSchema),
        asOf: textSchema(100, 1),
        jurisdiction: textSchema(300, 1),
        professionalContext: PROFESSIONAL_CONTEXT_SCHEMA,
        calculation: objectSchema(
          {
            formula: textSchema(2000, 1),
            inputs: { ...arraySchema(textSchema(1000, 1), 32, 1), uniqueItems: false },
            unit: textSchema(100, 1),
            currency: textSchema(100, 1),
          },
          ['formula', 'inputs'],
        ),
      },
      [
        'claimId',
        'statement',
        'type',
        'sourceRefs',
        'sourceTier',
        'slideRefs',
        'confidence',
        'reviewStatus',
        'conflictsWith',
      ],
    ),
  ),
})
function fail(): never {
  throw new Error('invalid_state')
}
const object = (v: unknown): v is Record<string, unknown> =>
  !!v && typeof v === 'object' && !Array.isArray(v)
const exact = (v: unknown, required: string[], optional: string[] = []) =>
  object(v) &&
  required.every((k) => Object.hasOwn(v, k)) &&
  Object.keys(v).every((k) => (required.includes(k) || optional.includes(k)) && v[k] !== undefined)
const text = (v: unknown, max: number, min = 0): v is string =>
  typeof v === 'string' &&
  v.length >= min &&
  v.length <= max &&
  (min === 0 || !!v.trim()) &&
  !v.includes('\0')
const id = (v: unknown): v is string => typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
const digest = (v: unknown): v is string => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v)
const integer = (v: unknown, min: number, max = Number.MAX_SAFE_INTEGER): v is number =>
  typeof v === 'number' && Number.isSafeInteger(v) && v >= min && v <= max
const time = (v: unknown): v is string =>
  typeof v === 'string' &&
  /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(v) &&
  Number.isFinite(Date.parse(v)) &&
  new Date(v).toISOString() === v
function size(value: unknown, max: number) {
  try {
    if (new TextEncoder().encode(JSON.stringify(value)).length > max) fail()
  } catch {
    fail()
  }
}
function ids(v: unknown, max = 64): v is string[] {
  return Array.isArray(v) && v.length <= max && v.every(id) && new Set(v).size === v.length
}
function asOf(v: unknown) {
  if (v === undefined) return true
  if (!text(v, 100, 1)) return false
  if (/^\d{4}-\d\d-\d\d$/.test(v))
    return Number.isFinite(Date.parse(v)) && new Date(v).toISOString().slice(0, 10) === v
  return true
}
function uri(v: unknown) {
  if (!text(v, 500, 1)) return false
  if (/^attachment:[a-f0-9]{64}$/.test(v)) return true
  try {
    const u = new URL(v)
    return (
      ['https:', 'http:'].includes(u.protocol) &&
      !u.username &&
      !u.password &&
      !!u.hostname &&
      !['localhost', '127.0.0.1', '::1', '[::1]'].includes(u.hostname)
    )
  } catch {
    return false
  }
}
export function parsePresentationResearchDraft(value: unknown): PresentationResearchDraft {
  size(value, 256 * 1024)
  if (!exact(value, ['scope', 'sources', 'facts'])) fail()
  const d = value as PresentationResearchDraft
  if (
    !text(d.scope, 4000, 1) ||
    !Array.isArray(d.sources) ||
    d.sources.length > 64 ||
    !Array.isArray(d.facts) ||
    d.facts.length > 64
  )
    fail()
  const sourceIds = new Set<string>(),
    factIds = new Set<string>()
  for (const s of d.sources) {
    if (
      !exact(s, ['id', 'title', 'uri', 'excerpt'], ['snapshotAttachmentId', 'locator', 'asOf']) ||
      !id(s.id) ||
      sourceIds.has(s.id) ||
      !text(s.title, 300, 1) ||
      !uri(s.uri) ||
      !text(s.excerpt, 12000) ||
      (s.snapshotAttachmentId !== undefined && !digest(s.snapshotAttachmentId)) ||
      (s.locator !== undefined && !text(s.locator, 200)) ||
      !asOf(s.asOf)
    )
      fail()
    sourceIds.add(s.id)
    const attached = /^attachment:([a-f0-9]{64})$/.exec(s.uri)?.[1]
    if (attached && s.snapshotAttachmentId !== undefined && s.snapshotAttachmentId !== attached)
      fail()
  }
  for (const f of d.facts) {
    if (
      !exact(
        f,
        [
          'claimId',
          'statement',
          'type',
          'sourceRefs',
          'sourceTier',
          'slideRefs',
          'confidence',
          'reviewStatus',
          'conflictsWith',
        ],
        ['asOf', 'jurisdiction', 'calculation', 'professionalContext'],
      ) ||
      !id(f.claimId) ||
      factIds.has(f.claimId) ||
      !text(f.statement, 12000, 1) ||
      !['fact', 'quote', 'calculation', 'judgment', 'assumption'].includes(f.type) ||
      !ids(f.sourceRefs) ||
      f.sourceRefs.some((s) => !sourceIds.has(s)) ||
      !['primary', 'authoritative_secondary', 'secondary', 'unverified'].includes(f.sourceTier) ||
      !ids(f.slideRefs) ||
      !['high', 'medium', 'low'].includes(f.confidence) ||
      f.reviewStatus !== 'needs_review' ||
      !ids(f.conflictsWith) ||
      f.conflictsWith.includes(f.claimId) ||
      !asOf(f.asOf) ||
      (f.jurisdiction !== undefined && !text(f.jurisdiction, 300, 1))
    )
      fail()
    if (Object.hasOwn(f, 'professionalContext')) {
      try {
        parsePresentationProfessionalContext(f.professionalContext)
      } catch {
        fail()
      }
    }
    factIds.add(f.claimId)
    if (f.type === 'calculation' && !f.calculation) fail()
    if (f.calculation !== undefined) {
      const c = f.calculation
      if (
        f.type !== 'calculation' ||
        !exact(c, ['formula', 'inputs'], ['unit', 'currency']) ||
        !text(c.formula, 2000, 1) ||
        !Array.isArray(c.inputs) ||
        !c.inputs.length ||
        c.inputs.length > 32 ||
        !c.inputs.every((s) => text(s, 1000, 1)) ||
        (c.unit !== undefined && !text(c.unit, 100, 1)) ||
        (c.currency !== undefined && !text(c.currency, 100, 1))
      )
        fail()
    }
  }
  if (d.facts.some((f) => f.conflictsWith.some((other) => !factIds.has(other)))) fail()
  return structuredClone(d)
}
const checks = {
  scope: 'research_draft',
  support: 'not_verified',
  sourceAuthority: 'not_verified',
  timeliness: 'not_verified',
} as const
export const presentationResearchChecks = checks
export function parsePresentationResearchRecord(value: unknown): PresentationResearchRecord {
  size(value, 512 * 1024)
  if (
    !exact(
      value,
      [
        'version',
        'documentId',
        'projectId',
        'id',
        'sequence',
        'draftDigest',
        'draft',
        'state',
        'startedAt',
        'checks',
      ],
      ['finishedAt', 'sources', 'error'],
    )
  )
    fail()
  const r = value as PresentationResearchRecord
  const draft = parsePresentationResearchDraft(r.draft)
  if (
    r.version !== 1 ||
    !text(r.documentId, 4096, 1) ||
    !id(r.projectId) ||
    !id(r.id) ||
    !integer(r.sequence, 1, 128) ||
    !digest(r.draftDigest) ||
    !time(r.startedAt) ||
    !['running', 'completed', 'failed'].includes(r.state) ||
    !exact(r.checks, Object.keys(checks)) ||
    Object.entries(checks).some(
      ([k, v]) => (r.checks as unknown as Record<string, unknown>)[k] !== v,
    )
  )
    fail()
  if (r.state === 'running') {
    if (r.finishedAt !== undefined || r.sources !== undefined || r.error !== undefined) fail()
  } else {
    if (!time(r.finishedAt) || r.finishedAt < r.startedAt) fail()
    if (r.state === 'failed') {
      if (!['aborted', 'source_unavailable', 'invalid_state'].includes(r.error!)) fail()
    } else if (
      r.error !== undefined ||
      !Array.isArray(r.sources) ||
      r.sources.length !== draft.sources.length
    )
      fail()
  }
  if (r.sources !== undefined) {
    if (!Array.isArray(r.sources) || r.sources.length > draft.sources.length) fail()
    for (const [i, e] of r.sources.entries()) {
      const source = draft.sources[i]
      if (
        !source ||
        !exact(
          e,
          ['sourceId', 'status', 'provenance'],
          ['attachmentId', 'offset', 'locator', 'retrievedAt', 'sha256', 'parsedTextSha256'],
        ) ||
        e.sourceId !== source.id ||
        ![
          'found',
          'not_found',
          'empty_excerpt',
          'not_ready',
          'unsupported',
          'missing',
          'source_mismatch',
        ].includes(e.status) ||
        !['user_supplied', 'fetched_url_matched', 'unavailable'].includes(e.provenance)
      )
        fail()
      const attachment =
        /^attachment:([a-f0-9]{64})$/.exec(source.uri)?.[1] ?? source.snapshotAttachmentId
      if (
        e.attachmentId !== undefined &&
        (!digest(e.attachmentId) || e.attachmentId !== attachment)
      )
        fail()
      if (e.status === 'found') {
        if (!e.attachmentId || !integer(e.offset, 0, 8_000_000) || e.provenance === 'unavailable')
          fail()
      } else if (e.offset !== undefined || e.locator !== undefined) fail()
      if (
        e.locator !== undefined &&
        (!text(e.locator, 200) || !/^第 [1-9]\d{0,5} (页|段)$/.test(e.locator))
      )
        fail()
      if (e.sha256 !== undefined && (!digest(e.sha256) || e.sha256 !== e.attachmentId)) fail()
      if (e.parsedTextSha256 !== undefined && !digest(e.parsedTextSha256)) fail()
      if (
        e.provenance === 'unavailable' &&
        (e.retrievedAt !== undefined || e.sha256 !== undefined || e.parsedTextSha256 !== undefined)
      )
        fail()
      if (
        e.provenance === 'fetched_url_matched' &&
        (!source.snapshotAttachmentId ||
          !/^https?:/.test(source.uri) ||
          !e.attachmentId ||
          !e.sha256 ||
          !time(e.retrievedAt) ||
          ['missing', 'not_ready', 'unsupported', 'source_mismatch'].includes(e.status))
      )
        fail()
      if (
        e.provenance === 'user_supplied' &&
        (!e.attachmentId ||
          !e.sha256 ||
          e.retrievedAt !== undefined ||
          ['missing', 'not_ready', 'unsupported', 'source_mismatch'].includes(e.status))
      )
        fail()
    }
  }
  return structuredClone(r)
}
function envelope(value: unknown) {
  if (!exact(value, ['version', 'documentId', 'projectId', 'revision', 'totalRecords', 'records']))
    fail()
  const h = value as PresentationResearchHistory
  if (
    h.version !== 1 ||
    !text(h.documentId, 4096, 1) ||
    !id(h.projectId) ||
    !integer(h.revision, 0, 256) ||
    !integer(h.totalRecords, 0, 128) ||
    h.revision < h.totalRecords ||
    h.revision > 2 * h.totalRecords ||
    !Array.isArray(h.records) ||
    h.records.length !== Math.min(32, h.totalRecords)
  )
    fail()
  return h
}
function window(value: PresentationResearchHistory | PresentationResearchSummary) {
  let completed = 0,
    last = ''
  const idsSeen = new Set<string>()
  for (const [i, r] of value.records.entries()) {
    if (
      'documentId' in r &&
      ((r as PresentationResearchRecord).documentId !== value.documentId ||
        (r as PresentationResearchRecord).projectId !== value.projectId)
    )
      fail()
    if (
      r.sequence !== value.totalRecords - value.records.length + i + 1 ||
      idsSeen.has(r.id) ||
      r.startedAt < last
    )
      fail()
    idsSeen.add(r.id)
    last = r.startedAt
    if (r.state !== 'running') completed++
  }
  if (
    value.revision < value.totalRecords + completed ||
    value.revision > value.totalRecords + completed + value.totalRecords - value.records.length
  )
    fail()
}
export function parsePresentationResearchHistory(value: unknown): PresentationResearchHistory {
  size(value, 16 * 1024 * 1024)
  const h = envelope(value)
  h.records.forEach(parsePresentationResearchRecord)
  window(h)
  return structuredClone(h)
}
export function parsePresentationResearchSummary(value: unknown): PresentationResearchSummary {
  size(value, 64 * 1024)
  const h = envelope(value) as unknown as PresentationResearchSummary
  for (const r of h.records) {
    if (
      !exact(
        r,
        [
          'id',
          'sequence',
          'draftDigest',
          'state',
          'startedAt',
          'sourceCount',
          'factCount',
          'conflictCount',
        ],
        ['finishedAt', 'error'],
      ) ||
      !id(r.id) ||
      !integer(r.sequence, 1, 128) ||
      !digest(r.draftDigest) ||
      !['running', 'completed', 'failed'].includes(r.state) ||
      !time(r.startedAt) ||
      !integer(r.sourceCount, 0, 64) ||
      !integer(r.factCount, 0, 64) ||
      !integer(r.conflictCount, 0, (r.factCount * (r.factCount - 1)) / 2)
    )
      fail()
    if (
      r.state === 'running'
        ? r.finishedAt !== undefined || r.error !== undefined
        : !time(r.finishedAt) ||
          r.finishedAt < r.startedAt ||
          (r.state === 'failed'
            ? !['aborted', 'source_unavailable', 'invalid_state'].includes(r.error!)
            : r.error !== undefined)
    )
      fail()
  }
  window(h)
  return structuredClone(h)
}
export function summarizePresentationResearchHistory(
  value: PresentationResearchHistory,
): PresentationResearchSummary {
  const h = parsePresentationResearchHistory(value)
  return parsePresentationResearchSummary({
    ...h,
    records: h.records.map((r) => ({
      id: r.id,
      sequence: r.sequence,
      draftDigest: r.draftDigest,
      state: r.state,
      startedAt: r.startedAt,
      ...(r.finishedAt ? { finishedAt: r.finishedAt } : {}),
      ...(r.error ? { error: r.error } : {}),
      sourceCount: r.draft.sources.length,
      factCount: r.draft.facts.length,
      conflictCount: new Set(
        r.draft.facts.flatMap((f) =>
          f.conflictsWith.map((other) => [f.claimId, other].sort().join(':')),
        ),
      ).size,
    })),
  })
}
