import { MAX_PRESENTATION_SOURCE_TEXT_CHARS } from '@wiswork/pptx-engine/presentation-source-limits'

export interface PresentationSourceAuditResult {
  sourceId: string
  attachmentId: string
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
}
export interface PresentationSourceAuditRun {
  id: string
  sequence: number
  scope: 'source_excerpt_audit'
  planRevision: number
  planDigest: string
  sourceRefs: { sourceId: string; attachmentId: string }[]
  state: 'running' | 'completed' | 'failed'
  startedAt: string
  finishedAt?: string
  sources?: PresentationSourceAuditResult[]
  error?: 'aborted' | 'source_unavailable' | 'invalid_state'
}
export interface PresentationSourceAuditLedger {
  version: 1
  projectId: string
  documentId: string
  revision: number
  runs: PresentationSourceAuditRun[]
}
export interface PresentationSourceAuditHistory extends Omit<
  PresentationSourceAuditLedger,
  'runs'
> {
  runs: (Omit<PresentationSourceAuditRun, 'sourceRefs' | 'sources'> & {
    sourceCount: number
    foundCount?: number
  })[]
}
const id = (value: unknown) => typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const hash = (value: unknown) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const date = (value: unknown) =>
  typeof value === 'string' &&
  Number.isFinite(Date.parse(value)) &&
  new Date(value).toISOString() === value
const fail = () => new Error('invalid_state')
function identity(run: Omit<PresentationSourceAuditRun, 'sourceRefs' | 'sources'>): boolean {
  return Boolean(
    run &&
    id(run.id) &&
    Number.isSafeInteger(run.sequence) &&
    run.sequence > 0 &&
    run.scope === 'source_excerpt_audit' &&
    Number.isSafeInteger(run.planRevision) &&
    run.planRevision > 0 &&
    hash(run.planDigest) &&
    date(run.startedAt) &&
    ['running', 'completed', 'failed'].includes(run.state) &&
    (run.state === 'running'
      ? run.finishedAt === undefined
      : date(run.finishedAt) && run.finishedAt! >= run.startedAt) &&
    (run.state === 'failed'
      ? ['aborted', 'source_unavailable', 'invalid_state'].includes(run.error!)
      : run.error === undefined),
  )
}
function root(value: PresentationSourceAuditLedger): boolean {
  return Boolean(
    value &&
    typeof value === 'object' &&
    !Array.isArray(value) &&
    Object.keys(value).sort().join(',') === 'documentId,projectId,revision,runs,version' &&
    value.version === 1 &&
    id(value.projectId) &&
    typeof value.documentId === 'string' &&
    value.documentId.trim() &&
    value.documentId.length <= 2048 &&
    Number.isSafeInteger(value.revision) &&
    value.revision >= 0 &&
    Array.isArray(value.runs) &&
    value.runs.length <= 32 &&
    new Set(value.runs.map((run) => run?.id)).size === value.runs.length &&
    value.runs.every(
      (run, index) =>
        run &&
        run.sequence <= value.revision &&
        (!index || run.sequence > value.runs[index - 1]!.sequence),
    ),
  )
}
export function parsePresentationSourceAuditRun(value: unknown): PresentationSourceAuditRun {
  const run = value as PresentationSourceAuditRun
  const keys = [
    'id',
    'sequence',
    'scope',
    'planRevision',
    'planDigest',
    'sourceRefs',
    'state',
    'startedAt',
    ...(run?.state === 'running' ? [] : ['finishedAt']),
    ...(run?.state === 'completed' ? ['sources'] : []),
    ...(run?.state === 'failed' ? ['error'] : []),
  ]
  if (
    !identity(run) ||
    Object.keys(run).sort().join(',') !== keys.sort().join(',') ||
    !Array.isArray(run.sourceRefs) ||
    run.sourceRefs.length > 256 ||
    run.sourceRefs.some(
      (ref) =>
        !ref ||
        Object.keys(ref).sort().join(',') !== 'attachmentId,sourceId' ||
        !id(ref.sourceId) ||
        !hash(ref.attachmentId),
    ) ||
    new Set(run.sourceRefs.map((ref) => ref.sourceId)).size !== run.sourceRefs.length ||
    (run.state === 'completed' &&
      (!Array.isArray(run.sources) ||
        run.sources.length !== run.sourceRefs.length ||
        run.sources.some(
          (source, index) =>
            !source ||
            source.sourceId !== run.sourceRefs[index]!.sourceId ||
            source.attachmentId !== run.sourceRefs[index]!.attachmentId ||
            ![
              'found',
              'not_found',
              'empty_excerpt',
              'not_ready',
              'unsupported',
              'missing',
              'source_mismatch',
            ].includes(source.status) ||
            (source.status === 'found'
              ? !Number.isSafeInteger(source.offset) ||
                source.offset! < 0 ||
                source.offset! > MAX_PRESENTATION_SOURCE_TEXT_CHARS
              : source.offset !== undefined || source.locator !== undefined) ||
            (source.locator !== undefined &&
              (typeof source.locator !== 'string' ||
                !/^第 [1-9]\d{0,5} (页|段)$/.test(source.locator))) ||
            Object.keys(source).sort().join(',') !==
              (source.status === 'found'
                ? source.locator === undefined
                  ? 'attachmentId,offset,sourceId,status'
                  : 'attachmentId,locator,offset,sourceId,status'
                : 'attachmentId,sourceId,status'),
        )))
  )
    throw fail()
  return structuredClone(run)
}
export function parsePresentationSourceAuditLedger(value: unknown): PresentationSourceAuditLedger {
  const ledger = value as PresentationSourceAuditLedger
  if (!root(ledger)) throw fail()
  return { ...structuredClone(ledger), runs: ledger.runs.map(parsePresentationSourceAuditRun) }
}
export function presentationSourceAuditHistory(
  ledger: PresentationSourceAuditLedger,
): PresentationSourceAuditHistory {
  const value = parsePresentationSourceAuditLedger(ledger)
  return {
    ...value,
    runs: value.runs.map(({ sourceRefs, sources, ...run }) => ({
      ...run,
      sourceCount: sourceRefs.length,
      ...(sources
        ? { foundCount: sources.filter((source) => source.status === 'found').length }
        : {}),
    })),
  }
}
export function parsePresentationSourceAuditHistory(
  value: unknown,
): PresentationSourceAuditHistory {
  const history = value as PresentationSourceAuditHistory
  if (
    !root(history as unknown as PresentationSourceAuditLedger) ||
    history.runs.some(
      (run) =>
        !identity(run) ||
        Object.keys(run).sort().join(',') !==
          [
            'id',
            'sequence',
            'scope',
            'planRevision',
            'planDigest',
            'state',
            'startedAt',
            'sourceCount',
            ...(run.state === 'running' ? [] : ['finishedAt']),
            ...(run.state === 'completed' ? ['foundCount'] : []),
            ...(run.state === 'failed' ? ['error'] : []),
          ]
            .sort()
            .join(',') ||
        !Number.isSafeInteger(run.sourceCount) ||
        run.sourceCount < 0 ||
        run.sourceCount > 256 ||
        (run.state === 'completed' &&
          (!Number.isSafeInteger(run.foundCount) ||
            run.foundCount! < 0 ||
            run.foundCount! > run.sourceCount)),
    )
  )
    throw fail()
  return structuredClone(history)
}
