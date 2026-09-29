import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import {
  selectionFingerprint,
  type StructuredProposalController,
} from '../../agent/proposal-controller.js'
import { parsePresentationPlan } from '@wiswork/pptx-engine/presentation-plan'
import {
  parsePresentationTeamContext,
  parsePresentationTeamLedger,
  presentationTeamId,
  type PresentationTeamLedger,
} from '@wiswork/pptx-engine/presentation-team'
interface Options {
  teamAvailable?(): boolean
  teamRequest?(body: unknown, signal?: AbortSignal): Promise<Response>
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  documentId(): Promise<string>
  proposals: StructuredProposalController
}
const operations = {
  read_presentation_team_identity: 'team_identity',
  create_presentation_team: 'team_project_create',
  read_presentation_team: 'team_project_read',
  publish_presentation_team_plan: 'team_plan_publish',
  set_presentation_team_member: 'team_member_set',
  revoke_presentation_team_member: 'team_member_revoke',
  add_presentation_team_comment: 'team_comment_add',
  resolve_presentation_team_comment: 'team_comment_resolve',
} as const
const fields: Record<string, string[]> = {
  team_identity: [],
  team_project_create: ['project_id', 'plan_revision'],
  team_project_read: ['team_id'],
  team_plan_publish: ['team_id', 'expected_revision', 'plan_revision'],
  team_member_set: ['team_id', 'expected_revision', 'member_subject', 'role'],
  team_member_revoke: ['team_id', 'expected_revision', 'member_subject'],
  team_comment_add: ['team_id', 'expected_revision', 'comment'],
  team_comment_resolve: ['team_id', 'expected_revision', 'comment_id'],
}
const properties: Record<string, unknown> = {
  project_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,80}$' },
  plan_revision: { type: 'integer', minimum: 1 },
  team_id: { type: 'string', pattern: '^team_[a-f0-9]{64}$' },
  expected_revision: { type: 'integer', minimum: 1 },
  member_subject: { type: 'string', pattern: '^[a-f0-9]{64}$' },
  role: { type: 'string', enum: ['reviewer', 'viewer'] },
  comment_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
  comment: {
    type: 'object',
    properties: {
      id: { type: 'string' },
      targetKind: { type: 'string', enum: ['slide', 'claim', 'source'] },
      targetId: { type: 'string' },
      text: { type: 'string', maxLength: 2000 },
    },
    required: ['id', 'targetKind', 'targetId', 'text'],
    additionalProperties: false,
  },
}
const tools: AgentToolDef[] = Object.entries(operations).map(([name, op]) => ({
  name,
  description:
    op === 'team_identity' || op === 'team_project_read'
      ? 'Read authenticated identity or exact published team snapshot. Private sources are not shared automatically.'
      : 'Propose a team ledger change requiring visible confirmation. This only changes the shared local-PC ledger, never the current host deck or factual/QA status.',
  inputSchema: {
    type: 'object',
    properties: Object.fromEntries(fields[op]!.map((f) => [f, properties[f]])),
    required: fields[op]!,
    additionalProperties: false,
  },
}))
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const exact = (v: unknown, keys: string[]): v is Record<string, unknown> =>
  !!v &&
  typeof v === 'object' &&
  !Array.isArray(v) &&
  Object.keys(v).sort().join(',') === keys.slice().sort().join(',')
function fail(code: string): never {
  throw Error(code)
}
const id = (v: unknown, max = 128) =>
  typeof v === 'string' && new RegExp(`^[A-Za-z0-9_-]{1,${max}}$`).test(v)
const positive = (v: unknown) => typeof v === 'number' && Number.isSafeInteger(v) && v >= 1
export function createPresentationTeamSkill(options: Options): AgentSkill & { clear(): void } {
  let epoch = 0
  const available = () => Boolean(options.teamAvailable?.() && options.teamRequest)
  return {
    id: 'presentation-team',
    get tools() {
      return available() ? tools : []
    },
    systemPrompt:
      'Team collaboration uses authenticated team_identity and explicit ACLs on this local PC. Only a visibly confirmed create/publish shares the exact private saved plan snapshot. Reviewer/viewer reads use only the published team ledger. Never invent an author, scan private projects, or treat team comments/resolve as factual, professional or QA acceptance. Team writes do not edit the current plan or PowerPoint. OAuth login must be real; absent authentication hides these tools.',
    clear() {
      epoch++
    },
    async executeTool(call, signal) {
      const captured = epoch
      const check = () => {
        if (signal?.aborted || captured !== epoch) fail('cancelled')
        if (!available()) fail('presentation_team_unavailable')
      }
      try {
        check()
        const op = operations[call.name as keyof typeof operations]
        if (!op || call.inputError || call.truncated || !exact(call.input, fields[op]!))
          fail('invalid_tool_input')
        const input = call.input
        for (const [field, value] of Object.entries(input)) {
          if (
            ['project_id', 'comment_id'].includes(field) &&
            !id(value, field === 'project_id' ? 80 : 128)
          )
            fail('invalid_tool_input')
          if (
            field === 'team_id' &&
            (typeof value !== 'string' || !/^team_[a-f0-9]{64}$/.test(value))
          )
            fail('invalid_tool_input')
          if (['plan_revision', 'expected_revision'].includes(field) && !positive(value))
            fail('invalid_tool_input')
          if (
            field === 'member_subject' &&
            (typeof value !== 'string' || !/^[a-f0-9]{64}$/.test(value))
          )
            fail('invalid_tool_input')
          if (field === 'role' && !['reviewer', 'viewer'].includes(String(value)))
            fail('invalid_tool_input')
        }
        if (op === 'team_comment_add') {
          const c = input.comment
          if (
            !exact(c, ['id', 'targetKind', 'targetId', 'text']) ||
            !id(c.id) ||
            !id(c.targetId) ||
            !['slide', 'claim', 'source'].includes(String(c.targetKind)) ||
            typeof c.text !== 'string' ||
            !c.text.trim() ||
            c.text.length > 2000 ||
            Array.from(c.text).some(
              (x) => x.charCodeAt(0) < 32 || (x.charCodeAt(0) >= 127 && x.charCodeAt(0) <= 159),
            )
          )
            fail('invalid_tool_input')
        }
        const documentId = await options.documentId()
        check()
        const guard = async (s?: AbortSignal) => {
          check()
          if (s?.aborted) fail('cancelled')
          if ((await options.documentId()) !== documentId) fail('presentation_document_changed')
          check()
          if (s?.aborted) fail('cancelled')
        }
        const json = async (request: Options['request'], body: unknown, s?: AbortSignal) => {
          await guard(s)
          let response: Response, raw: string
          try {
            response = await request(body, s)
          } catch {
            await guard(s)
            return fail('presentation_team_unavailable')
          }
          await guard(s)
          try {
            raw = await response.text()
          } catch {
            await guard(s)
            return fail('presentation_response_invalid')
          }
          await guard(s)
          if (new TextEncoder().encode(raw).length > 800 * 1024)
            fail('presentation_response_invalid')
          let value: unknown
          try {
            value = JSON.parse(raw)
          } catch {
            fail('presentation_response_invalid')
          }
          if (!response.ok) {
            const error = exact(value, ['error']) ? value.error : undefined
            fail(
              [
                'access_denied',
                'revision_conflict',
                'not_found',
                'invalid_state',
                'quota_exceeded',
                'invalid_request',
              ].includes(String(error))
                ? String(error)
                : 'presentation_team_unavailable',
            )
          }
          return value
        }
        const teamRequest = options.teamRequest!
        const identity = async (s?: AbortSignal) => {
          const value = await json(teamRequest, { operation: 'team_identity' }, s)
          if (!exact(value, ['identity'])) fail('presentation_response_invalid')
          try {
            return parsePresentationTeamContext(value.identity)
          } catch {
            return fail('presentation_response_invalid')
          }
        }
        const actor = await identity(signal)
        const accountGuard = async (s?: AbortSignal) => {
          if (!same(await identity(s), actor)) fail('presentation_team_account_changed')
        }
        const parseTeam = (value: unknown) => {
          if (!exact(value, ['team'])) return fail('presentation_response_invalid')
          try {
            return parsePresentationTeamLedger(value.team)
          } catch {
            return fail('presentation_response_invalid')
          }
        }
        const readTeam = async (s?: AbortSignal) => {
          await accountGuard(s)
          const team = parseTeam(
            await json(teamRequest, { operation: 'team_project_read', teamId: input.team_id }, s),
          )
          await accountGuard(s)
          if (
            team.teamId !== input.team_id ||
            team.ownerSubject !== actor.pcSubject ||
            (actor.actorSubject !== team.ownerSubject &&
              !team.members.some((m) => m.subject === actor.actorSubject))
          )
            fail('access_denied')
          return team
        }
        if (op === 'team_identity')
          return {
            output: JSON.stringify({ identity: actor }),
            mutated: false,
            summary: '已读取实际团队账号身份',
          }
        const team = op === 'team_project_create' ? undefined : await readTeam(signal)
        if (op === 'team_project_read')
          return {
            output: JSON.stringify({ team }),
            mutated: false,
            summary: '已读取已发布团队快照；不代表事实或 QA 通过',
          }
        const owner =
          actor.actorSubject === actor.pcSubject &&
          (team === undefined || team.ownerSubject === actor.actorSubject)
        if (
          [
            'team_project_create',
            'team_plan_publish',
            'team_member_set',
            'team_member_revoke',
          ].includes(op) &&
          !owner
        )
          fail('access_denied')
        if (team && input.expected_revision !== team.revision) fail('revision_conflict')
        if (op === 'team_comment_add' || op === 'team_comment_resolve') {
          const role =
            team!.ownerSubject === actor.actorSubject
              ? 'owner'
              : team!.members.find((m) => m.subject === actor.actorSubject)?.role
          if (role !== 'owner' && role !== 'reviewer') fail('access_denied')
          if (op === 'team_comment_resolve') {
            const c = team!.comments.find((c) => c.id === input.comment_id)
            if (!c || c.state !== 'open') fail('invalid_tool_input')
            if (role !== 'owner' && c.authorSubject !== actor.actorSubject) fail('access_denied')
          }
          if (op === 'team_comment_add') {
            const c = input.comment as Record<string, unknown>
            const list =
              c.targetKind === 'slide'
                ? team!.publishedPlan.plan.slides
                : c.targetKind === 'claim'
                  ? team!.publishedPlan.plan.claims
                  : team!.publishedPlan.plan.sources
            if (!list.some((x) => x.id === c.targetId) || team!.comments.some((x) => x.id === c.id))
              fail('invalid_tool_input')
          }
        }
        if (
          ['team_member_set', 'team_member_revoke'].includes(op) &&
          input.member_subject === actor.pcSubject
        )
          fail('invalid_tool_input')
        const projectId = team?.projectId ?? (input.project_id as string)
        if (op === 'team_plan_publish' && team!.documentId !== documentId)
          fail('presentation_document_changed')
        const privatePlan = async (s?: AbortSignal) => {
          await accountGuard(s)
          const value = await json(
            teamRequest,
            {
              operation: 'team_plan_read',
              documentId,
              projectId,
              planRevision: input.plan_revision,
            },
            s,
          )
          if (
            !exact(value, ['projectId', 'revision', 'plan']) ||
            value.projectId !== projectId ||
            value.revision !== input.plan_revision
          )
            fail('presentation_response_invalid')
          let plan
          try {
            plan = parsePresentationPlan(value.plan)
          } catch {
            return fail('presentation_response_invalid')
          }
          if (plan.projectId !== projectId) fail('presentation_response_invalid')
          await accountGuard(s)
          return { revision: value.revision as number, plan }
        }
        const published = ['team_project_create', 'team_plan_publish'].includes(op)
          ? await privatePlan(signal)
          : undefined
        const teamId =
          team?.teamId ?? (await presentationTeamId(actor.pcSubject, documentId, projectId))
        await accountGuard(signal)
        const body: Record<string, unknown> = {
          operation: op,
          expectedIdentity: actor,
          ...(op === 'team_project_create'
            ? { documentId, projectId, planRevision: input.plan_revision }
            : { teamId, expectedRevision: input.expected_revision }),
        }
        if (op === 'team_plan_publish') body.planRevision = input.plan_revision
        if (op === 'team_member_set' || op === 'team_member_revoke')
          body.memberSubject = input.member_subject
        if (op === 'team_member_set') body.role = input.role
        if (op === 'team_comment_add') {
          body.planRevision = team!.publishedPlan.revision
          body.comment = input.comment
        }
        if (op === 'team_comment_resolve') body.commentId = input.comment_id
        const fresh = async (s?: AbortSignal) => {
          await accountGuard(s)
          if (team && !same(await readTeam(s), team)) fail('revision_conflict')
          if (published && !same(await privatePlan(s), published))
            fail('presentation_team_source_changed')
          await guard(s)
        }
        let planSummary: Record<string, unknown> | undefined
        if (published) {
          const serialized = JSON.stringify(published.plan)
          const fullUtf8Bytes = new TextEncoder().encode(serialized).length
          const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(serialized))
          await accountGuard(signal)
          const planDigest = Array.from(new Uint8Array(digest), (n) =>
            n.toString(16).padStart(2, '0'),
          ).join('')
          const window = (text: string | undefined) => {
            const original = text ?? ''
            let summary = ''
            for (const c of original) {
              if (new TextEncoder().encode(summary + c).length > 120) break
              summary += c
            }
            return {
              summary,
              utf8Bytes: new TextEncoder().encode(original).length,
              truncated: summary !== original,
            }
          }
          let count = Math.min(16, published.plan.sources.length)
          do {
            planSummary = {
              title: published.plan.title,
              revision: published.revision,
              fullUtf8Bytes,
              planDigest,
              sourceCount: published.plan.sources.length,
              sourceIds: published.plan.sources.map((source) => source.id),
              sources: published.plan.sources.slice(0, count).map((source) => ({
                id: source.id,
                title: window(source.title),
                uri: window(source.uri),
                locator: window(source.locator),
                excerpt: window(source.excerpt),
              })),
              omittedSourceDetails: published.plan.sources.length - count,
              slides: published.plan.slides.map((slide) => ({
                id: slide.id,
                title: window(slide.title),
              })),
              excerptSummaryOnly: true,
              sharesFullSavedExcerpts: true,
              disclosure:
                '展示的是来源元数据和摘录摘要窗口，truncated 标记裁剪，omittedSourceDetails 标记未展示详情数量。完整来源 ID 均列出；确认会共享完整已保存计划及所有原始来源摘录，而非仅这些摘要。',
            }
            count--
          } while (
            new TextEncoder().encode(JSON.stringify(planSummary)).length > 32 * 1024 &&
            count >= 0
          )
          if (new TextEncoder().encode(JSON.stringify(planSummary)).length > 32 * 1024)
            fail('presentation_team_unavailable')
        }
        const proposal = options.proposals.propose({
          operation: call.name,
          toolName: call.name,
          title: '确认团队共享记录变更',
          preview: {
            operation: op,
            identity: actor,
            teamId,
            documentId,
            projectId,
            ...(published ? { planSummary, planRevision: published.revision } : {}),
            ...(team ? { expectedRevision: team.revision } : {}),
            ...input,
            note: '仅修改本机 PC 团队账本；不会编辑当前计划、宿主或批准事实/QA。',
          },
          impact: { host: 'local_team', targets: [teamId], count: 1 },
          fingerprint: selectionFingerprint(
            JSON.stringify([actor, documentId, body, published, team]),
          ),
          validate: async (s) => {
            try {
              await fresh(s)
              return true
            } catch {
              return false
            }
          },
          execute: async (s) => {
            await fresh(s)
            const receipt = parseTeam(await json(teamRequest, body, s))
            await accountGuard(s)
            if (
              receipt.teamId !== teamId ||
              receipt.ownerSubject !== actor.pcSubject ||
              receipt.documentId !== (team?.documentId ?? documentId) ||
              receipt.projectId !== projectId ||
              receipt.revision !== (team ? team.revision + 1 : 1)
            )
              fail('presentation_response_invalid')
            const expected: PresentationTeamLedger = {
              ...(team ?? receipt),
              revision: receipt.revision,
              updatedAt: receipt.updatedAt,
            }
            if (published) expected.publishedPlan = published
            if (op === 'team_member_set')
              expected.members = [
                ...team!.members.filter((m) => m.subject !== input.member_subject),
                {
                  subject: input.member_subject as string,
                  role: input.role as 'reviewer' | 'viewer',
                },
              ]
            if (op === 'team_member_revoke')
              expected.members = team!.members.filter((m) => m.subject !== input.member_subject)
            if (op === 'team_comment_add') {
              const c = receipt.comments.find(
                (c) => c.id === (input.comment as Record<string, unknown>).id,
              )
              if (
                !c ||
                c.authorSubject !== actor.actorSubject ||
                c.state !== 'open' ||
                c.planRevision !== team!.publishedPlan.revision ||
                !same(
                  { id: c.id, targetKind: c.targetKind, targetId: c.targetId, text: c.text },
                  input.comment,
                )
              )
                fail('presentation_response_invalid')
              expected.comments = [...team!.comments, c]
            }
            if (op === 'team_comment_resolve')
              expected.comments = team!.comments.map((c) =>
                c.id === input.comment_id
                  ? {
                      ...c,
                      state: 'resolved' as const,
                      updatedAt: receipt.comments.find((x) => x.id === c.id)?.updatedAt ?? '',
                    }
                  : c,
              )
            if (op === 'team_project_create' && (receipt.members.length || receipt.comments.length))
              fail('presentation_response_invalid')
            if (!same(receipt, expected)) fail('presentation_response_invalid')
          },
        })
        return {
          output: JSON.stringify({
            proposalId: proposal.id,
            status: 'awaiting_confirmation',
            teamId,
          }),
          mutated: false,
          summary: '团队变更提案等待明确确认',
        }
      } catch (error) {
        const code = error instanceof Error ? error.message : ''
        return {
          output: [
            'invalid_tool_input',
            'cancelled',
            'presentation_team_unavailable',
            'presentation_document_changed',
            'presentation_response_invalid',
            'presentation_team_account_changed',
            'presentation_team_source_changed',
            'access_denied',
            'revision_conflict',
            'not_found',
            'invalid_state',
            'quota_exceeded',
            'invalid_request',
          ].includes(code)
            ? code
            : 'presentation_team_unavailable',
          isError: true,
          mutated: false,
          summary: '团队操作未完成',
        }
      }
    },
  }
}
