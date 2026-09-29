import { parsePresentationPlan, type PresentationPlan } from './presentation-plan.js'
export interface PresentationTeamContext {
  version: 1
  actorSubject: string
  pcSubject: string
}
export type PresentationTeamRole = 'reviewer' | 'viewer'
export interface PresentationTeamComment {
  id: string
  targetKind: 'slide' | 'claim' | 'source'
  targetId: string
  authorSubject: string
  text: string
  planRevision: number
  state: 'open' | 'resolved'
  createdAt: string
  updatedAt: string
}
export interface PresentationTeamLedger {
  version: 1
  teamId: string
  documentId: string
  projectId: string
  ownerSubject: string
  revision: number
  createdAt: string
  updatedAt: string
  publishedPlan: { revision: number; plan: PresentationPlan }
  members: { subject: string; role: PresentationTeamRole }[]
  comments: PresentationTeamComment[]
}
const fail = (): never => {
  throw Error('invalid_request')
}
const object = (value: unknown): Record<string, unknown> => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return fail()
  return value as Record<string, unknown>
}
const exact = (value: Record<string, unknown>, keys: string[]) =>
  Object.keys(value).sort().join(',') === keys.sort().join(',')
const subject = (value: unknown) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const id = (value: unknown, max = 128) =>
  typeof value === 'string' && new RegExp(`^[A-Za-z0-9_-]{1,${max}}$`).test(value)
const integer = (value: unknown) =>
  typeof value === 'number' && Number.isSafeInteger(value) && value >= 1
const text = (value: unknown, max: number) =>
  typeof value === 'string' &&
  !!value.trim() &&
  value.length <= max &&
  !Array.from(value).some(
    (c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159),
  )
const time = (value: unknown) =>
  typeof value === 'string' &&
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(value) &&
  Number.isFinite(Date.parse(value)) &&
  new Date(value).toISOString() === value
export function parsePresentationTeamContext(value: unknown): PresentationTeamContext {
  const context = object(value)
  if (
    !exact(context, ['version', 'actorSubject', 'pcSubject']) ||
    context.version !== 1 ||
    !subject(context.actorSubject) ||
    !subject(context.pcSubject)
  )
    return fail()
  return structuredClone(context) as unknown as PresentationTeamContext
}
export function parsePresentationTeamLedger(value: unknown): PresentationTeamLedger {
  const ledger = object(value)
  if (
    new TextEncoder().encode(JSON.stringify(value)).length > 768 * 1024 ||
    !exact(ledger, [
      'version',
      'teamId',
      'documentId',
      'projectId',
      'ownerSubject',
      'revision',
      'createdAt',
      'updatedAt',
      'publishedPlan',
      'members',
      'comments',
    ]) ||
    ledger.version !== 1 ||
    typeof ledger.teamId !== 'string' ||
    !/^team_[a-f0-9]{64}$/.test(ledger.teamId) ||
    typeof ledger.documentId !== 'string' ||
    !ledger.documentId ||
    ledger.documentId.length > 4096 ||
    !id(ledger.projectId, 80) ||
    !subject(ledger.ownerSubject) ||
    !integer(ledger.revision) ||
    !time(ledger.createdAt) ||
    !time(ledger.updatedAt) ||
    (ledger.updatedAt as string) < (ledger.createdAt as string)
  )
    return fail()
  const published = object(ledger.publishedPlan)
  if (!exact(published, ['revision', 'plan']) || !integer(published.revision)) return fail()
  const plan = parsePresentationPlan(published.plan)
  if (plan.projectId !== ledger.projectId) return fail()
  if (
    !Array.isArray(ledger.members) ||
    ledger.members.length > 32 ||
    !Array.isArray(ledger.comments) ||
    ledger.comments.length > 128
  )
    return fail()
  const members = ledger.members.map((value) => {
    const member = object(value)
    if (
      !exact(member, ['subject', 'role']) ||
      !subject(member.subject) ||
      member.subject === ledger.ownerSubject ||
      !['reviewer', 'viewer'].includes(String(member.role))
    )
      return fail()
    return member
  })
  if (new Set(members.map((m) => m.subject)).size !== members.length) return fail()
  const comments = ledger.comments.map((value) => {
    const comment = object(value)
    if (
      !exact(comment, [
        'id',
        'targetKind',
        'targetId',
        'authorSubject',
        'text',
        'planRevision',
        'state',
        'createdAt',
        'updatedAt',
      ]) ||
      !id(comment.id) ||
      !id(comment.targetId) ||
      !subject(comment.authorSubject) ||
      !text(comment.text, 2000) ||
      !integer(comment.planRevision) ||
      (comment.planRevision as number) > (published.revision as number) ||
      !['slide', 'claim', 'source'].includes(String(comment.targetKind)) ||
      !['open', 'resolved'].includes(String(comment.state)) ||
      !time(comment.createdAt) ||
      !time(comment.updatedAt) ||
      (comment.createdAt as string) < (ledger.createdAt as string) ||
      (comment.updatedAt as string) < (comment.createdAt as string) ||
      (comment.updatedAt as string) > (ledger.updatedAt as string)
    )
      return fail()
    if (
      comment.planRevision === published.revision &&
      !(
        comment.targetKind === 'slide'
          ? plan.slides
          : comment.targetKind === 'claim'
            ? plan.claims
            : plan.sources
      ).some((item) => item.id === comment.targetId)
    )
      return fail()
    return comment
  })
  if (new Set(comments.map((c) => c.id)).size !== comments.length) return fail()
  return structuredClone({
    ...ledger,
    publishedPlan: { revision: published.revision, plan },
  }) as unknown as PresentationTeamLedger
}
export async function presentationTeamId(
  pcSubject: string,
  documentId: string,
  projectId: string,
): Promise<string> {
  if (!subject(pcSubject) || !documentId || documentId.length > 4096 || !id(projectId, 80))
    return fail()
  const digest = await crypto.subtle.digest(
    'SHA-256',
    new TextEncoder().encode(JSON.stringify([pcSubject, documentId, projectId])),
  )
  return (
    'team_' + Array.from(new Uint8Array(digest), (n) => n.toString(16).padStart(2, '0')).join('')
  )
}
