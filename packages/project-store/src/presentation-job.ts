/** Browser-safe bounded production history; retained events are a recovery view, not an audit log. */
export type PresentationProductionJobState =
  | 'running'
  | 'pausing'
  | 'paused'
  | 'cancelling'
  | 'cancelled'
  | 'interrupted'
  | 'completed'
  | 'failed'
export const PRESENTATION_PRODUCTION_ERRORS = Object.freeze([
  'compile_failed',
  'invalid_deck',
  'aborted',
  'output_too_large',
  'asset_unavailable',
  'source_unavailable',
  'font_unavailable',
] as const)
export type PresentationProductionError = (typeof PRESENTATION_PRODUCTION_ERRORS)[number]
export type PresentationProductionJobEventInput =
  | {
      type:
        | 'run.started'
        | 'run.pause_requested'
        | 'run.paused'
        | 'run.cancel_requested'
        | 'run.cancelled'
        | 'run.interrupted'
        | 'run.completed'
    }
  | { type: 'run.failed'; error?: PresentationProductionError | 'invalid_state' }
  | { type: 'page.started' | 'page.compiled'; pageId: string; attempt: number }
  | { type: 'page.failed'; pageId: string; attempt: number; error: PresentationProductionError }
export type PresentationProductionJobEvent = PresentationProductionJobEventInput & {
  sequence: number
  createdAt: string
}
export interface PresentationProductionJob {
  version: 1
  projectId: string
  documentId: string
  requestId: string
  inputDigest: string
  planDigest: string
  planRevision: number
  revision: number
  state: PresentationProductionJobState
  events: PresentationProductionJobEvent[]
}
const states: PresentationProductionJobState[] = [
  'running',
  'pausing',
  'paused',
  'cancelling',
  'cancelled',
  'interrupted',
  'completed',
  'failed',
]
function invalid(): never {
  throw new Error('invalid_state')
}
function object(value: unknown): Record<string, unknown> {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.getPrototypeOf(value) !== Object.prototype
  )
    invalid()
  const record = value as Record<string, unknown>
  if (Object.values(Object.getOwnPropertyDescriptors(record)).some((d) => !('value' in d)))
    invalid()
  return record
}
function exact(value: Record<string, unknown>, fields: string[]): void {
  if (Object.keys(value).some((key) => !fields.includes(key))) invalid()
}
function positive(value: unknown): boolean {
  return Number.isSafeInteger(value) && Number(value) > 0
}
function id(value: unknown): boolean {
  return typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
}
export function presentationProductionJobStateAfter(
  state: PresentationProductionJobState | undefined,
  event: PresentationProductionJobEventInput,
): PresentationProductionJobState {
  const allowed = (
    from: Array<PresentationProductionJobState | undefined>,
    next: PresentationProductionJobState,
  ) => {
    if (!from.includes(state)) invalid()
    return next
  }
  switch (event.type) {
    case 'run.started':
      return allowed([undefined, 'paused', 'interrupted', 'failed'], 'running')
    case 'run.pause_requested':
      return allowed(['running'], 'pausing')
    case 'run.paused':
      return allowed(['pausing'], 'paused')
    case 'run.cancel_requested':
      return allowed(['running', 'pausing', 'paused', 'interrupted', 'failed'], 'cancelling')
    case 'run.cancelled':
      return allowed(['cancelling'], 'cancelled')
    case 'run.interrupted':
      return allowed(['running'], 'interrupted')
    case 'run.completed':
      return allowed(['running', 'pausing'], 'completed')
    case 'run.failed':
      return allowed(['running', 'pausing', 'cancelling'], 'failed')
    case 'page.started':
      return allowed(['running'], 'running')
    case 'page.compiled':
    case 'page.failed':
      return allowed(['running', 'pausing', 'cancelling'], state!)
  }
}
export function parsePresentationProductionJob(value: unknown): PresentationProductionJob {
  const job = object(value)
  exact(job, [
    'version',
    'projectId',
    'documentId',
    'requestId',
    'inputDigest',
    'planDigest',
    'planRevision',
    'revision',
    'state',
    'events',
  ])
  if (
    job.version !== 1 ||
    !id(job.projectId) ||
    !id(job.requestId) ||
    typeof job.documentId !== 'string' ||
    !job.documentId.trim() ||
    job.documentId.length > 2048 ||
    !positive(job.planRevision) ||
    !positive(job.revision) ||
    !states.includes(job.state as PresentationProductionJobState) ||
    [job.inputDigest, job.planDigest].some(
      (v) => typeof v !== 'string' || !/^[a-f0-9]{64}$/.test(v),
    ) ||
    !Array.isArray(job.events) ||
    job.events.length !== Math.min(Number(job.revision), 128)
  )
    invalid()
  const events = job.events as unknown[]
  let possibilities: Array<PresentationProductionJobState | undefined> =
    Number(job.revision) <= 128 ? [undefined] : [...states]
  let previousTime = ''
  const active = new Map<string, number>()
  let unknownActive = Number(job.revision) > 128 ? 2 : 0
  const attempts = new Map<string, number>()
  for (const [index, value] of events.entries()) {
    const event = object(value)
    const page = typeof event.type === 'string' && event.type.startsWith('page.')
    const failed = event.type === 'page.failed' || event.type === 'run.failed'
    exact(event, [
      'sequence',
      'createdAt',
      'type',
      ...(page ? ['pageId', 'attempt'] : []),
      ...(failed ? ['error'] : []),
    ])
    if (
      event.sequence !== Number(job.revision) - events.length + index + 1 ||
      typeof event.createdAt !== 'string' ||
      !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(event.createdAt) ||
      !Number.isFinite(Date.parse(event.createdAt)) ||
      new Date(event.createdAt).toISOString() !== event.createdAt ||
      event.createdAt < previousTime ||
      (page && (!id(event.pageId) || !positive(event.attempt))) ||
      (event.type === 'page.failed' &&
        !(PRESENTATION_PRODUCTION_ERRORS as readonly string[]).includes(event.error as string)) ||
      (event.type === 'run.failed' &&
        Object.hasOwn(event, 'error') &&
        ![...PRESENTATION_PRODUCTION_ERRORS, 'invalid_state'].includes(
          event.error as PresentationProductionError | 'invalid_state',
        ))
    )
      invalid()
    if (event.type === 'run.started') {
      active.clear()
      unknownActive = 0
    }
    if (page) {
      const pageId = event.pageId as string
      const attempt = event.attempt as number
      if (event.type === 'page.started') {
        if (active.size >= 2 || active.has(pageId) || attempt <= (attempts.get(pageId) ?? 0))
          invalid()
        active.set(pageId, attempt)
      } else {
        if (active.has(pageId)) {
          if (active.get(pageId) !== attempt) invalid()
        } else if (unknownActive) unknownActive--
        else invalid()
        active.delete(pageId)
      }
      attempts.set(pageId, attempt)
    }
    if (
      typeof event.type === 'string' &&
      ['run.paused', 'run.cancelled', 'run.completed'].includes(event.type) &&
      active.size
    )
      invalid()
    previousTime = event.createdAt
    possibilities = possibilities.flatMap((state) => {
      try {
        const next = presentationProductionJobStateAfter(
          state,
          event as PresentationProductionJobEvent,
        )
        return next ? [next] : []
      } catch {
        return []
      }
    })
    if (!possibilities.length) invalid()
  }
  if (!possibilities.includes(job.state as PresentationProductionJobState)) invalid()
  return structuredClone(job) as unknown as PresentationProductionJob
}
