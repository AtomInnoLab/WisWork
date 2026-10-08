/** Browser-safe immutable issue disposition contract. Explained never means resolved. */
export interface PresentationIssueActionInput {
  actionId: string
  issueId: string
  issueDigest: string
  state: 'open' | 'deferred' | 'explained'
  note: string
}
export interface PresentationIssueAction extends PresentationIssueActionInput {
  sequence: number
  createdAt: string
}
export interface PresentationIssueLedger {
  version: 1
  projectId: string
  documentId: string
  requestId: string
  inputDigest: string
  planDigest: string
  revision: number
  actions: PresentationIssueAction[]
}
function object(value: unknown, keys: string[]): Record<string, unknown> {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).length !== keys.length ||
    keys.some((key) => !Object.hasOwn(value, key))
  )
    throw new Error('invalid_request')
  return value as Record<string, unknown>
}
const id = (value: unknown) => typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const hash = (value: unknown) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const inputKeys = ['actionId', 'issueId', 'issueDigest', 'state', 'note']
export function parsePresentationIssueActionInput(value: unknown): PresentationIssueActionInput {
  const item = object(value, inputKeys)
  if (
    !id(item.actionId) ||
    !id(item.issueId) ||
    !hash(item.issueDigest) ||
    !['open', 'deferred', 'explained'].includes(item.state as string) ||
    typeof item.note !== 'string' ||
    !item.note.trim() ||
    item.note.length > 2000 ||
    // XML 1.0 deliberately permits tab, LF, and CR.
    // eslint-disable-next-line no-control-regex
    !/^[\u0009\u000a\u000d\u0020-\ud7ff\ue000-\ufffd\u{10000}-\u{10ffff}]*$/u.test(item.note)
  )
    throw new Error('invalid_request')
  return {
    actionId: item.actionId as string,
    issueId: item.issueId as string,
    issueDigest: item.issueDigest as string,
    state: item.state as PresentationIssueActionInput['state'],
    note: item.note.trim(),
  }
}
export function parsePresentationIssueLedger(value: unknown): PresentationIssueLedger {
  const item = object(value, [
    'version',
    'projectId',
    'documentId',
    'requestId',
    'inputDigest',
    'planDigest',
    'revision',
    'actions',
  ])
  if (
    item.version !== 1 ||
    !id(item.projectId) ||
    !id(item.requestId) ||
    typeof item.documentId !== 'string' ||
    !item.documentId.trim() ||
    item.documentId.length > 2048 ||
    !hash(item.inputDigest) ||
    !hash(item.planDigest) ||
    !Array.isArray(item.actions) ||
    item.actions.length > 128 ||
    item.revision !== item.actions.length
  )
    throw new Error('invalid_request')
  const ids = new Set<string>()
  const actions = item.actions.map((value: unknown, index: number) => {
    const action = object(value, [...inputKeys, 'sequence', 'createdAt'])
    const { sequence, createdAt, ...input } = action
    const parsed = parsePresentationIssueActionInput(input)
    if (
      parsed.note !== input.note ||
      sequence !== index + 1 ||
      ids.has(parsed.actionId) ||
      typeof createdAt !== 'string' ||
      !Number.isFinite(Date.parse(createdAt)) ||
      new Date(createdAt).toISOString() !== createdAt
    )
      throw new Error('invalid_request')
    ids.add(parsed.actionId)
    return { ...parsed, sequence: index + 1, createdAt }
  })
  return {
    version: 1,
    projectId: item.projectId as string,
    documentId: item.documentId,
    requestId: item.requestId as string,
    inputDigest: item.inputDigest as string,
    planDigest: item.planDigest as string,
    revision: actions.length,
    actions,
  }
}
