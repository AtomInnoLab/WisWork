import {
  registerPresentationProjectWork,
  type PresentationProjectWork,
} from './presentation-project-work'
import { createHash, randomUUID } from 'node:crypto'
import {
  constants,
  closeSync,
  fstatSync,
  lstatSync,
  mkdirSync,
  openSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { join } from 'node:path'
import {
  parsePresentationTeamContext,
  parsePresentationTeamLedger,
  type PresentationTeamContext,
  type PresentationTeamLedger,
} from '@wiswork/pptx-engine/presentation-team'
import {
  parsePresentationPlan,
  type PresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'
const LIMIT = 768 * 1024
const present = (path: string) => lstatSync(path, { throwIfNoEntry: false })
const fail = (code: string): never => {
  throw Error(code)
}
const subject = (value: unknown) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const integer = (value: unknown) =>
  typeof value === 'number' && Number.isSafeInteger(value) && value >= 1
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const canonical = (value: unknown): string =>
  JSON.stringify(value, (_, v) =>
    v && typeof v === 'object' && !Array.isArray(v)
      ? Object.fromEntries(Object.entries(v).sort(([a], [b]) => a.localeCompare(b)))
      : v,
  )
const teamId = (owner: string, doc: string, project: string) =>
  'team_' +
  createHash('sha256')
    .update(JSON.stringify([owner, doc, project]))
    .digest('hex')
const exact = (body: Record<string, unknown>, keys: string[]) => {
  if (Object.keys(body).sort().join(',') !== keys.sort().join(',')) fail('invalid_request')
}
const check = (signal: AbortSignal) => {
  if (signal.aborted) fail('aborted')
}
/** Local PC ACL and published content share one atomic file. Private plans remain outside this ledger. */
export class PresentationTeamStore {
  private root: string
  constructor(userDataPath: string) {
    this.root = join(userDataPath, 'presentation-teams')
  }
  private path(id: string, create = false) {
    if (!/^team_[a-f0-9]{64}$/.test(id)) return fail('invalid_request')
    const stat = present(this.root)
    if (stat && (!stat.isDirectory() || stat.isSymbolicLink())) return fail('invalid_state')
    if (!stat && create) mkdirSync(this.root, { recursive: true, mode: 0o700 })
    return join(this.root, id + '.json')
  }
  read(id: string): PresentationTeamLedger | undefined {
    const path = this.path(id),
      stat = present(path)
    if (!stat) return undefined
    if (!stat.isFile() || stat.isSymbolicLink() || stat.size > LIMIT) return fail('invalid_state')
    let fd: number | undefined
    try {
      fd = openSync(path, constants.O_RDONLY | constants.O_NOFOLLOW)
      const opened = fstatSync(fd)
      if (!opened.isFile() || opened.size > LIMIT) return fail('invalid_state')
      const raw = readFileSync(fd)
      if (raw.length > LIMIT) return fail('invalid_state')
      const ledger = parsePresentationTeamLedger(JSON.parse(raw.toString('utf8')))
      if (
        ledger.teamId !== id ||
        teamId(ledger.ownerSubject, ledger.documentId, ledger.projectId) !== id
      )
        return fail('invalid_state')
      return ledger
    } catch {
      return fail('invalid_state')
    } finally {
      if (fd !== undefined) closeSync(fd)
    }
  }
  write(
    next: PresentationTeamLedger,
    expected: PresentationTeamLedger | undefined,
    assertWritable?: () => void,
  ): PresentationTeamLedger {
    next = structuredClone(next)
    expected = expected === undefined ? undefined : structuredClone(expected)
    if (!same(this.read(next.teamId), expected)) return fail('revision_conflict')
    const serialized = JSON.stringify(next)
    if (Buffer.byteLength(serialized) > LIMIT) return fail('quota_exceeded')
    const parsed = parsePresentationTeamLedger(next)
    if (teamId(parsed.ownerSubject, parsed.documentId, parsed.projectId) !== parsed.teamId)
      return fail('invalid_state')
    this.path(next.teamId)
    assertWritable?.()
    const path = this.path(next.teamId, true),
      temporary = path + '.' + randomUUID() + '.tmp'
    let created = false
    try {
      assertWritable?.()
      writeFileSync(temporary, serialized, { encoding: 'utf8', mode: 0o600, flag: 'wx' })
      created = true
      assertWritable?.()
      renameSync(temporary, path)
      created = false
    } finally {
      if (created) rmSync(temporary, { force: true })
    }
    return structuredClone(parsed)
  }
}
interface Options {
  userDataPath: string
  readPlan(
    documentId: string,
    projectId: string,
  ): { revision: number; plan: PresentationPlan } | undefined
  acquireProjectLock(projectId: string, signal?: AbortSignal): Promise<() => void>
  captureProjectLease?(
    scope: Readonly<{ documentId: string; projectId: string }>,
    mode: 'read' | 'write',
    signal: AbortSignal,
  ): { assertCurrent(): void }
}
export function createPresentationTeamService(options: Options) {
  options = { ...options }
  const store = new PresentationTeamStore(options.userDataPath)
  const capture = (
    documentId: string,
    projectId: string,
    mode: 'read' | 'write',
    signal: AbortSignal,
    register?: (scope: Readonly<{ documentId: string; projectId: string }>) => AbortSignal,
  ) => {
    signal = register?.({ documentId, projectId }) ?? signal
    const lease = options.captureProjectLease?.(
      Object.freeze({ documentId, projectId }),
      mode,
      signal,
    )
    return () => {
      check(signal)
      lease?.assertCurrent()
      check(signal)
    }
  }
  const authorized = (id: unknown, context: PresentationTeamContext) => {
    if (typeof id !== 'string') return fail('invalid_request')
    const ledger = store.read(id)
    if (!ledger) return fail('not_found')
    if (
      ledger.ownerSubject !== context.pcSubject ||
      (context.actorSubject !== ledger.ownerSubject &&
        !ledger.members.some((m) => m.subject === context.actorSubject))
    )
      return fail('access_denied')
    return ledger
  }
  const owner = (ledger: PresentationTeamLedger, context: PresentationTeamContext) => {
    if (context.actorSubject !== ledger.ownerSubject) return fail('access_denied')
  }
  const snapshot = (doc: string, project: string, revision: unknown) => {
    if (!integer(revision)) return fail('invalid_request')
    const saved = options.readPlan(doc, project)
    if (!saved) return fail('not_found')
    if (saved.revision !== revision) return fail('revision_conflict')
    const plan = parsePresentationPlan(saved.plan)
    if (plan.projectId !== project) return fail('invalid_state')
    return { revision: saved.revision, plan }
  }
  return async (
    body: Record<string, unknown>,
    contextValue: unknown,
    signal: AbortSignal,
  ): Promise<
    | { identity: PresentationTeamContext }
    | { team: PresentationTeamLedger }
    | { projectId: string; revision: number; plan: PresentationPlan }
  > => {
    check(signal)
    body = structuredClone(body)
    let context: PresentationTeamContext
    try {
      context = structuredClone(parsePresentationTeamContext(contextValue))
    } catch {
      return fail('access_denied')
    }
    let work: PresentationProjectWork | undefined
    const register = (scope: Readonly<{ documentId: string; projectId: string }>) => {
      if (options.captureProjectLease) {
        work = registerPresentationProjectWork({
          scope: { root: options.userDataPath, ...scope },
          signal,
        })
        signal = work.signal
      }
      return signal
    }
    try {
      const operation = body.operation
      if (!['team_identity', 'team_project_read', 'team_plan_read'].includes(String(operation))) {
        let expected: PresentationTeamContext
        try {
          expected = parsePresentationTeamContext(body.expectedIdentity)
        } catch {
          return fail('access_denied')
        }
        if (
          expected.version !== context.version ||
          expected.actorSubject !== context.actorSubject ||
          expected.pcSubject !== context.pcSubject
        )
          return fail('access_denied')
      }
      if (operation === 'team_identity') {
        exact(body, ['operation'])
        return { identity: context }
      }
      if (operation === 'team_plan_read') {
        exact(body, ['operation', 'documentId', 'projectId', 'planRevision'])
        if (context.actorSubject !== context.pcSubject) return fail('access_denied')
        if (
          typeof body.documentId !== 'string' ||
          !body.documentId ||
          body.documentId.length > 4096 ||
          typeof body.projectId !== 'string' ||
          !/^[A-Za-z0-9_-]{1,80}$/.test(body.projectId)
        )
          return fail('invalid_request')
        snapshot(body.documentId, body.projectId, body.planRevision)
        const guard = capture(body.documentId, body.projectId, 'read', signal, register)
        guard()
        const release = await options.acquireProjectLock(body.projectId, signal)
        try {
          check(signal)
          guard()
          const saved = snapshot(body.documentId, body.projectId, body.planRevision)
          guard()
          return { projectId: body.projectId, ...saved }
        } finally {
          release()
        }
      }
      if (operation === 'team_project_create') {
        exact(body, ['operation', 'documentId', 'projectId', 'planRevision', 'expectedIdentity'])
        if (context.actorSubject !== context.pcSubject) return fail('access_denied')
        if (
          typeof body.documentId !== 'string' ||
          !body.documentId ||
          body.documentId.length > 4096 ||
          typeof body.projectId !== 'string' ||
          !/^[A-Za-z0-9_-]{1,80}$/.test(body.projectId)
        )
          return fail('invalid_request')
        const doc = body.documentId,
          project = body.projectId
        snapshot(doc, project, body.planRevision)
        const guard = capture(doc, project, 'write', signal, register)
        guard()
        const release = await options.acquireProjectLock(project, signal)
        try {
          check(signal)
          guard()
          const publishedPlan = snapshot(doc, project, body.planRevision),
            id = teamId(context.pcSubject, doc, project),
            existing = store.read(id)
          if (existing) {
            if (
              existing.ownerSubject !== context.pcSubject ||
              existing.documentId !== doc ||
              existing.projectId !== project ||
              canonical(existing.publishedPlan) !== canonical(publishedPlan)
            )
              return fail('revision_conflict')
            return { team: existing }
          }
          const now = new Date().toISOString()
          return {
            team: store.write(
              {
                version: 1,
                teamId: id,
                documentId: doc,
                projectId: project,
                ownerSubject: context.pcSubject,
                revision: 1,
                createdAt: now,
                updatedAt: now,
                publishedPlan,
                members: [],
                comments: [],
              },
              undefined,
              guard,
            ),
          }
        } finally {
          release()
        }
      }
      const expectedKeys: Record<string, string[]> = {
        team_project_read: ['operation', 'teamId'],
        team_member_set: ['operation', 'teamId', 'expectedRevision', 'memberSubject', 'role'],
        team_member_revoke: ['operation', 'teamId', 'expectedRevision', 'memberSubject'],
        team_plan_publish: ['operation', 'teamId', 'expectedRevision', 'planRevision'],
        team_comment_add: ['operation', 'teamId', 'expectedRevision', 'planRevision', 'comment'],
        team_comment_resolve: ['operation', 'teamId', 'expectedRevision', 'commentId'],
      }
      if (typeof operation !== 'string' || !expectedKeys[operation]) return fail('invalid_request')
      exact(body, [
        ...expectedKeys[operation]!,
        ...(operation === 'team_project_read' ? [] : ['expectedIdentity']),
      ])
      let ledger = authorized(body.teamId, context)
      if (
        operation !== 'team_project_read' &&
        ['team_plan_publish', 'team_member_set', 'team_member_revoke'].includes(operation)
      )
        owner(ledger, context)
      else if (
        operation !== 'team_project_read' &&
        context.actorSubject !== ledger.ownerSubject &&
        !ledger.members.some((m) => m.subject === context.actorSubject && m.role === 'reviewer')
      )
        return fail('access_denied')
      const guard = capture(
        ledger.documentId,
        ledger.projectId,
        operation === 'team_project_read' ? 'read' : 'write',
        signal,
        register,
      )
      guard()
      if (operation === 'team_project_read') return { team: ledger }
      const release = await options.acquireProjectLock(ledger.projectId, signal)
      try {
        guard()
        ledger = authorized(body.teamId, context)
        guard()
        const commit = (next: PresentationTeamLedger) => {
          check(signal)
          if (ledger.revision === Number.MAX_SAFE_INTEGER) return fail('quota_exceeded')
          return {
            team: store.write(
              {
                ...next,
                revision: ledger.revision + 1,
                updatedAt: new Date(
                  Math.max(
                    Date.now(),
                    Date.parse(ledger.updatedAt),
                    ...next.comments.map((comment) => Date.parse(comment.updatedAt)),
                  ),
                ).toISOString(),
              },
              ledger,
              guard,
            ),
          }
        }
        const cas = () => {
          if (body.expectedRevision !== ledger.revision) return fail('revision_conflict')
        }
        if (operation === 'team_plan_publish') {
          owner(ledger, context)
          cas()
          const publishedPlan = snapshot(ledger.documentId, ledger.projectId, body.planRevision)
          if (publishedPlan.revision < ledger.publishedPlan.revision)
            return fail('revision_conflict')
          return commit({ ...ledger, publishedPlan })
        }
        if (operation === 'team_member_set' || operation === 'team_member_revoke') {
          owner(ledger, context)
          cas()
          if (!subject(body.memberSubject) || body.memberSubject === ledger.ownerSubject)
            return fail('invalid_request')
          const members = ledger.members.filter((m) => m.subject !== body.memberSubject)
          if (operation === 'team_member_set') {
            if (body.role !== 'reviewer' && body.role !== 'viewer') return fail('invalid_request')
            members.push({ subject: body.memberSubject as string, role: body.role })
            if (members.length > 32) return fail('quota_exceeded')
          } else if (members.length === ledger.members.length) return fail('not_found')
          return commit({ ...ledger, members })
        }
        const role =
          context.actorSubject === ledger.ownerSubject
            ? 'owner'
            : ledger.members.find((m) => m.subject === context.actorSubject)?.role
        if (role !== 'owner' && role !== 'reviewer') return fail('access_denied')
        if (operation === 'team_comment_add') {
          if (!body.comment || typeof body.comment !== 'object' || Array.isArray(body.comment))
            return fail('invalid_request')
          const comment = body.comment as Record<string, unknown>
          exact(comment, ['id', 'targetKind', 'targetId', 'text'])
          if (
            typeof comment.id !== 'string' ||
            !/^[A-Za-z0-9_-]{1,128}$/.test(comment.id) ||
            typeof comment.targetId !== 'string' ||
            !/^[A-Za-z0-9_-]{1,128}$/.test(comment.targetId) ||
            !['slide', 'claim', 'source'].includes(String(comment.targetKind)) ||
            typeof comment.text !== 'string' ||
            !comment.text.trim() ||
            comment.text.length > 2000 ||
            Array.from(comment.text).some(
              (c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159),
            ) ||
            !integer(body.planRevision)
          )
            return fail('invalid_request')
          const existing = ledger.comments.find((c) => c.id === comment.id)
          if (existing) {
            if (
              existing.authorSubject !== context.actorSubject ||
              existing.targetKind !== comment.targetKind ||
              existing.targetId !== comment.targetId ||
              existing.text !== comment.text ||
              existing.planRevision !== body.planRevision
            )
              return fail('revision_conflict')
            return { team: ledger }
          }
          cas()
          const plan = ledger.publishedPlan.plan
          if (
            body.planRevision !== ledger.publishedPlan.revision ||
            !(
              comment.targetKind === 'slide'
                ? plan.slides
                : comment.targetKind === 'claim'
                  ? plan.claims
                  : plan.sources
            ).some((item) => item.id === comment.targetId)
          )
            return fail('invalid_request')
          if (ledger.comments.length >= 128) return fail('quota_exceeded')
          const now = new Date(Math.max(Date.now(), Date.parse(ledger.updatedAt))).toISOString()
          return commit({
            ...ledger,
            comments: [
              ...ledger.comments,
              {
                id: comment.id,
                targetKind: comment.targetKind as 'slide' | 'claim' | 'source',
                targetId: comment.targetId,
                authorSubject: context.actorSubject,
                text: comment.text,
                planRevision: body.planRevision as number,
                state: 'open',
                createdAt: now,
                updatedAt: now,
              },
            ],
          })
        }
        cas()
        if (typeof body.commentId !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(body.commentId))
          return fail('invalid_request')
        const comment = ledger.comments.find((c) => c.id === body.commentId)
        if (!comment) return fail('not_found')
        if (role !== 'owner' && comment.authorSubject !== context.actorSubject)
          return fail('access_denied')
        if (comment.state !== 'open') return fail('invalid_state')
        const now = new Date(Math.max(Date.now(), Date.parse(ledger.updatedAt))).toISOString()
        return commit({
          ...ledger,
          comments: ledger.comments.map((c) =>
            c.id === comment.id ? { ...c, state: 'resolved', updatedAt: now } : c,
          ),
        })
      } finally {
        release()
      }
    } finally {
      work?.finish()
    }
  }
}
