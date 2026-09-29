import type { AgentSkill } from '@wiswork/agent-core'
import {
  parsePresentationTeamContext,
  parsePresentationTeamLedger,
  presentationTeamId,
  type PresentationTeamContext,
  type PresentationTeamLedger,
  type PresentationTeamRole,
} from '@wiswork/pptx-engine/presentation-team'
import type { StructuredProposalController } from './proposal-controller.js'
export interface PresentationTeamSnapshot {
  available: boolean
  phase: 'idle' | 'loading' | 'proposing' | 'awaiting_confirmation'
  identity?: PresentationTeamContext
  team?: PresentationTeamLedger
  role?: 'owner' | PresentationTeamRole
  proposalId?: string
  error?: string
  notice?: string
}
export interface PresentationTeamController {
  snapshot(): PresentationTeamSnapshot
  subscribe(listener: () => void): () => void
  refresh(teamId?: string): Promise<void>
  clear(): void
  create(projectId: string, planRevision: number): Promise<string>
  publish(planRevision: number): Promise<string>
  setMember(subject: string, role: PresentationTeamRole): Promise<string>
  revokeMember(subject: string): Promise<string>
  addComment(
    targetKind: 'slide' | 'claim' | 'source',
    targetId: string,
    text: string,
  ): Promise<string>
  resolveComment(id: string): Promise<string>
}
interface Options {
  skill: AgentSkill & { clear?(): void }
  documentId(): Promise<string>
  proposals: StructuredProposalController
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const safe = new Set([
  'access_denied',
  'revision_conflict',
  'not_found',
  'invalid_tool_input',
  'cancelled',
  'presentation_team_unavailable',
  'presentation_document_changed',
  'presentation_response_invalid',
  'presentation_team_account_changed',
  'presentation_team_busy',
  'proposal_stale',
  'presentation_team_source_changed',
  'invalid_state',
  'quota_exceeded',
  'invalid_request',
])
function freeze<T>(value: T): T {
  if (value && typeof value === 'object') {
    Object.values(value).forEach(freeze)
    Object.freeze(value)
  }
  return value
}
export function createPresentationTeamController(options: Options): PresentationTeamController {
  const available = () =>
    options.skill.tools.some((t) => t.name === 'read_presentation_team_identity')
  let state: PresentationTeamSnapshot = freeze({ available: available(), phase: 'idle' }),
    epoch = 0,
    active: AbortController | undefined,
    document: string | undefined
  let ownedProposal: string | undefined
  const listeners = new Set<() => void>()
  const publish = (next: PresentationTeamSnapshot) => {
    state = freeze(structuredClone(next))
    listeners.forEach((l) => l())
  }
  const errorCode = (e: unknown) =>
    e instanceof Error && safe.has(e.message) ? e.message : 'presentation_response_invalid'
  const fail = (code: string): never => {
    throw Error(code)
  }
  const currentGuard = (captured: number, signal: AbortSignal) => {
    if (captured !== epoch || signal.aborted) fail('cancelled')
    if (!available()) fail('presentation_team_unavailable')
  }
  const check = async (captured: number, doc: string, signal: AbortSignal) => {
    if (captured !== epoch || signal.aborted) fail('cancelled')
    if (!available()) fail('presentation_team_unavailable')
    if ((await options.documentId()) !== doc) fail('presentation_document_changed')
    if (captured !== epoch || signal.aborted) fail('cancelled')
    if (!available()) fail('presentation_team_unavailable')
  }
  const call = async (name: string, input: Record<string, unknown>, signal: AbortSignal) => {
    if (!available()) fail('presentation_team_unavailable')
    const result = await options.skill.executeTool({ id: crypto.randomUUID(), name, input }, signal)
    // Mutation callers first retain the proposal ID so a failed guard can cancel it.
    if (name.startsWith('read_') && !available()) fail('presentation_team_unavailable')
    if (result.isError)
      fail(safe.has(result.output) ? result.output : 'presentation_response_invalid')
    try {
      return JSON.parse(result.output) as Record<string, unknown>
    } catch {
      return fail('presentation_response_invalid')
    }
  }
  const read = async (
    teamId: string | undefined,
    captured: number,
    doc: string,
    signal: AbortSignal,
  ) => {
    if (!available()) fail('presentation_team_unavailable')
    const initial = await call('read_presentation_team_identity', {}, signal)
    await check(captured, doc, signal)
    if (Object.keys(initial).join(',') !== 'identity') fail('presentation_response_invalid')
    const identity = parsePresentationTeamContext(initial.identity)
    let team: PresentationTeamLedger | undefined
    if (teamId) {
      const value = await call('read_presentation_team', { team_id: teamId }, signal)
      await check(captured, doc, signal)
      if (Object.keys(value).join(',') !== 'team') fail('presentation_response_invalid')
      team = parsePresentationTeamLedger(value.team)
      if (team.teamId !== teamId || team.ownerSubject !== identity.pcSubject) fail('access_denied')
    }
    const final = await call('read_presentation_team_identity', {}, signal)
    await check(captured, doc, signal)
    if (
      Object.keys(final).join(',') !== 'identity' ||
      !same(parsePresentationTeamContext(final.identity), identity)
    )
      fail('presentation_team_account_changed')
    const role: PresentationTeamSnapshot['role'] = team
      ? team.ownerSubject === identity.actorSubject
        ? 'owner'
        : team.members.find((m) => m.subject === identity.actorSubject)?.role
      : undefined
    if (team && !role) fail('access_denied')
    return { identity, team, role }
  }
  const refresh = async (
    teamId = state.team?.teamId,
    expected?: { doc: string; identity: PresentationTeamContext },
  ) => {
    if (active || state.phase === 'awaiting_confirmation') fail('presentation_team_busy')
    const captured = epoch,
      controller = new AbortController()
    active = controller
    publish({ available: available(), phase: 'loading' })
    try {
      const doc = await options.documentId()
      if (expected && doc !== expected.doc) fail('presentation_document_changed')
      const result = await read(teamId, captured, doc, controller.signal)
      if (expected && !same(expected.identity, result.identity))
        fail('presentation_team_account_changed')
      currentGuard(captured, controller.signal)
      document = doc
      publish({ available: true, phase: 'idle', ...result })
    } catch (e) {
      if (captured === epoch)
        publish({ available: available(), phase: 'idle', error: errorCode(e) })
    } finally {
      if (active === controller) active = undefined
    }
  }
  const propose = async (
    name: string,
    input: Record<string, unknown>,
    ownerOnly: boolean,
    createScope?: { projectId: string },
  ) => {
    if (active || options.proposals.pending() || state.phase === 'awaiting_confirmation')
      fail('presentation_team_busy')
    const prior = state,
      captured = epoch,
      controller = new AbortController()
    active = controller
    publish({ ...prior, phase: 'proposing', error: undefined, notice: undefined })
    let proposedId: string | undefined
    try {
      const doc = await options.documentId()
      if (document !== undefined && document !== doc) fail('presentation_document_changed')
      const current = await read(prior.team?.teamId, captured, doc, controller.signal)
      if (prior.identity && !same(prior.identity, current.identity))
        fail('presentation_team_account_changed')
      if (prior.team && !same(prior.team, current.team)) fail('revision_conflict')
      if (
        ownerOnly
          ? current.identity.actorSubject !== current.identity.pcSubject
          : current.role !== 'owner' && current.role !== 'reviewer'
      )
        fail('access_denied')
      const body = createScope
        ? input
        : { team_id: current.team?.teamId, expected_revision: current.team?.revision, ...input }
      const value = await call(name, body, controller.signal)
      const id = value.proposalId
      if (typeof id === 'string' && options.proposals.pending()?.id === id) {
        proposedId = id
        if (captured === epoch) ownedProposal = id
      }
      await check(captured, doc, controller.signal)
      if (
        typeof value.proposalId !== 'string' ||
        options.proposals.pending()?.id !== value.proposalId
      )
        fail('presentation_response_invalid')
      if (typeof id !== 'string') return fail('presentation_response_invalid')
      const teamId = createScope
        ? await presentationTeamId(current.identity.pcSubject, doc, createScope.projectId)
        : current.team!.teamId
      await check(captured, doc, controller.signal)
      currentGuard(captured, controller.signal)
      const decision = options.proposals.waitForDecision(id)
      document = doc
      publish({
        available: available(),
        phase: 'awaiting_confirmation',
        ...current,
        proposalId: id,
      })
      void decision
        .then(async (result) => {
          if (ownedProposal === id) ownedProposal = undefined
          if (captured !== epoch) return
          if (result.status === 'confirmed') {
            publish({ available: available(), phase: 'idle' })
            // Read the persisted receipt; a successful proposal is not factual or QA acceptance.
            await refresh(teamId, { doc, identity: current.identity })
          } else
            publish({
              available: available(),
              phase: 'idle',
              ...(result.status === 'failed' ? { error: errorCode(Error(result.error)) } : {}),
            })
        })
        .catch(() => {
          if (captured === epoch)
            publish({
              available: available(),
              phase: 'idle',
              error: 'presentation_response_invalid',
            })
        })
      return id
    } catch (e) {
      if (proposedId && options.proposals.pending()?.id === proposedId) options.proposals.reject()
      if (ownedProposal === proposedId) ownedProposal = undefined
      if (captured === epoch)
        publish({ available: available(), phase: 'idle', error: errorCode(e) })
      // The UI receives only a finite safe code, never callback payloads.
      // eslint-disable-next-line preserve-caught-error
      throw Error(errorCode(e))
    } finally {
      if (active === controller) active = undefined
    }
  }
  return {
    snapshot: () => state,
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    refresh,
    clear() {
      epoch++
      active?.abort()
      active = undefined
      document = undefined
      if (ownedProposal && ownedProposal === options.proposals.pending()?.id)
        options.proposals.reject()
      ownedProposal = undefined
      options.skill.clear?.()
      publish({ available: available(), phase: 'idle' })
    },
    create: (projectId, planRevision) =>
      propose(
        'create_presentation_team',
        { project_id: projectId, plan_revision: planRevision },
        true,
        { projectId },
      ),
    publish: (planRevision) =>
      propose('publish_presentation_team_plan', { plan_revision: planRevision }, true),
    setMember: (subject, role) =>
      propose('set_presentation_team_member', { member_subject: subject, role }, true),
    revokeMember: (subject) =>
      propose('revoke_presentation_team_member', { member_subject: subject }, true),
    addComment: (targetKind, targetId, text) =>
      propose(
        'add_presentation_team_comment',
        { comment: { id: crypto.randomUUID(), targetKind, targetId, text } },
        false,
      ),
    resolveComment: (id) => propose('resolve_presentation_team_comment', { comment_id: id }, false),
  }
}
