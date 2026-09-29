import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import {
  parsePresentationManualObservation,
  parsePresentationManualObservationShape,
  presentationManualObservationDigest,
  MAX_PRESENTATION_MANUAL_OBSERVATIONS_BYTES,
  type PresentationManualObservation,
} from '@wiswork/pptx-engine/presentation-manual-observation'
import { parseSavedPresentationPreference } from '@wiswork/pptx-engine/presentation-preference'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import type { PresentationBaselineAdapter } from './browser-presentation-baseline-adapter.js'
interface Options {
  adapter: PresentationBaselineAdapter
  available(): boolean
  documentId(): Promise<string>
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  proposals: StructuredProposalController
}
const names = [
  'begin_presentation_edit_observation',
  'complete_presentation_edit_observation',
  'read_presentation_edit_observation',
  'list_presentation_edit_observations',
  'delete_presentation_edit_observation',
  'save_presentation_observed_preference',
]
const tools: AgentToolDef[] = names.map((name, index) => ({
  name,
  description:
    'Explicit local observation of text, geometry and aggregate font. Differences do not prove authorship. Delete and save require visible confirmation; no host writes.',
  inputSchema: {
    type: 'object',
    properties: {
      project_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,80}$' },
      ...(index === 0
        ? {
            slide_id: { type: 'string', minLength: 1, maxLength: 256 },
            shape_id: { type: 'string', minLength: 1, maxLength: 256 },
          }
        : index === 3
          ? {}
          : { observation_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,80}$' } }),
      ...(index === 5 ? { preference: { type: 'string', minLength: 1, maxLength: 240 } } : {}),
    },
    required: [
      'project_id',
      ...(index === 0 ? ['slide_id', 'shape_id'] : index === 3 ? [] : ['observation_id']),
      ...(index === 5 ? ['preference'] : []),
    ],
    additionalProperties: false,
  },
}))
const same = (a: unknown, b: unknown) =>
  canonicalPresentationValue(a) === canonicalPresentationValue(b)
export function createPresentationManualObservationSkill(
  options: Options,
): AgentSkill & { clear(): void } {
  let epoch = 0
  let pendingId: string | undefined
  return {
    id: 'office-presentation-manual-observations',
    clear() {
      epoch++
      if (pendingId && options.proposals.pending()?.id === pendingId) options.proposals.reject()
      pendingId = undefined
    },
    get tools() {
      return options.available() ? tools : []
    },
    systemPrompt:
      'Explicitly begin an exact object observation before editing and complete after editing or reopening. Differences are unattributed, not proof of human authorship. User-confirmed preferences are suggestions; brand rules prevail. Never monitor in background or write the host.',
    async executeTool(call, signal) {
      let ownedProposalId: string | undefined
      try {
        const captured = epoch
        const guard = () => {
          if (signal?.aborted || captured !== epoch) throw new Error('cancelled')
          if (!options.available()) throw new Error('presentation_unavailable')
        }
        guard()
        const name = call.name
        const index = names.indexOf(name)
        const input = call.input
        if (
          index < 0 ||
          call.inputError ||
          call.truncated ||
          !input ||
          typeof input !== 'object' ||
          Array.isArray(input)
        )
          throw new Error('invalid_tool_input')
        const value = structuredClone(input) as Record<string, unknown>
        const schema = tools[index]!.inputSchema as {
          required: string[]
          properties: Record<string, unknown>
        }
        if (
          Object.keys(value).some((k) => !Object.hasOwn(schema.properties, k)) ||
          schema.required.some((k) => !Object.hasOwn(value, k)) ||
          typeof value.project_id !== 'string' ||
          !/^[A-Za-z0-9_-]{1,80}$/.test(value.project_id)
        )
          throw new Error('invalid_tool_input')
        if (
          index !== 0 &&
          index !== 3 &&
          (typeof value.observation_id !== 'string' ||
            !/^[A-Za-z0-9_-]{1,80}$/.test(value.observation_id))
        )
          throw new Error('invalid_tool_input')
        if (
          index === 5 &&
          (typeof value.preference !== 'string' ||
            !value.preference.trim() ||
            value.preference.length > 240)
        )
          throw new Error('invalid_tool_input')
        const documentId = await options.documentId()
        guard()
        const check = async (s?: AbortSignal) => {
          guard()
          if (s?.aborted) throw new Error('cancelled')
          if ((await options.documentId()) !== documentId)
            throw new Error('presentation_document_changed')
          guard()
          if (s?.aborted) throw new Error('cancelled')
        }
        const scope = { documentId, projectId: value.project_id }
        const request = async (body: unknown, s?: AbortSignal) => {
          await check(s)
          const response = await options.request(body, s)
          await check(s)
          if (!response.ok) throw new Error('presentation_service_unavailable')
          const text = await response.text()
          await check(s)
          if (
            new TextEncoder().encode(text).length >
            MAX_PRESENTATION_MANUAL_OBSERVATIONS_BYTES + 64
          )
            throw new Error('presentation_response_invalid')
          try {
            return JSON.parse(text) as Record<string, unknown>
          } catch (cause) {
            throw new Error('presentation_response_invalid', { cause })
          }
        }
        const validateRecord = async (v: unknown, s?: AbortSignal) => {
          const record = parsePresentationManualObservation(v)
          if (record.documentId !== documentId || record.projectId !== scope.projectId)
            throw new Error('presentation_response_invalid')
          for (const snap of [record.before, record.after])
            if (snap && (await presentationManualObservationDigest(snap.shape)) !== snap.digest)
              throw new Error('presentation_response_invalid')
          await check(s)
          return record
        }
        const observation = async (body: unknown, s?: AbortSignal) => {
          const result = await request(body, s)
          if (Object.keys(result).join(',') !== 'observation')
            throw new Error('presentation_response_invalid')
          const record = await validateRecord(result.observation, s)
          if (value.observation_id !== undefined && record.observationId !== value.observation_id)
            throw new Error('presentation_response_invalid')
          return record
        }
        const stable = async (slideId: unknown, shapeId: unknown, s?: AbortSignal) => {
          if (
            typeof slideId !== 'string' ||
            !slideId ||
            slideId.length > 256 ||
            typeof shapeId !== 'string' ||
            !shapeId ||
            shapeId.length > 256
          )
            throw new Error('invalid_tool_input')
          const read = async () => {
            await check(s)
            const context = await options.adapter.readContext(s)
            await check(s)
            if (
              !Array.isArray(context.slideIds) ||
              context.slideIds.length > 500 ||
              new Set(context.slideIds).size !== context.slideIds.length ||
              !context.slideIds.includes(slideId)
            )
              throw new Error('presentation_observation_unavailable')
            const page = await options.adapter.readPage(slideId, s)
            await check(s)
            if (
              page.slideId !== slideId ||
              !Array.isArray(page.shapes) ||
              page.shapes.length > 100 ||
              new Set(page.shapes.map((x) => x.id)).size !== page.shapes.length
            )
              throw new Error('presentation_observation_unavailable')
            const target = page.shapes.find((x) => x.id === shapeId)
            if (!target) throw new Error('presentation_observation_unavailable')
            const { id, name, type, left, top, width, height, text, font } = target
            return {
              order: context.slideIds,
              shape: parsePresentationManualObservationShape({
                id,
                name,
                type,
                left,
                top,
                width,
                height,
                ...(text !== undefined ? { text } : {}),
                ...(font !== undefined ? { font } : {}),
              }),
            }
          }
          const before = await read(),
            after = await read()
          if (!same(before, after)) throw new Error('presentation_observation_unstable')
          return after.shape
        }
        const get = () =>
          observation(
            { operation: 'manual_observation_get', ...scope, observationId: value.observation_id },
            signal,
          )
        let output: unknown
        if (index === 0) {
          const shape = await stable(value.slide_id, value.shape_id, signal)
          const observationId = crypto.randomUUID()
          const record = await observation(
            {
              operation: 'manual_observation_begin',
              ...scope,
              observationId,
              slideId: value.slide_id,
              shape,
            },
            signal,
          )
          if (
            record.observationId !== observationId ||
            record.slideId !== value.slide_id ||
            record.after ||
            !same(record.before.shape, shape)
          )
            throw new Error('presentation_response_invalid')
          output = { observation: record }
        } else if (index === 1) {
          const before = await get()
          const shape = await stable(before.slideId, before.shapeId, signal)
          const record = await observation(
            {
              operation: 'manual_observation_complete',
              ...scope,
              observationId: before.observationId,
              expectedBeforeDigest: before.before.digest,
              shape,
            },
            signal,
          )
          if (
            !same(record.before, before.before) ||
            record.slideId !== before.slideId ||
            !record.after ||
            !same(record.after.shape, shape)
          )
            throw new Error('presentation_response_invalid')
          output = { observation: record }
        } else if (index === 2) output = { observation: await get() }
        else if (index === 3) {
          const result = await request({ operation: 'manual_observation_list', ...scope }, signal)
          if (
            Object.keys(result).join(',') !== 'observations' ||
            !Array.isArray(result.observations) ||
            result.observations.length > 32
          )
            throw new Error('presentation_response_invalid')
          const records: PresentationManualObservation[] = []
          for (const record of result.observations)
            records.push(await validateRecord(record, signal))
          if (new Set(records.map((x) => x.observationId)).size !== records.length)
            throw new Error('presentation_response_invalid')
          output = { observations: records }
        } else {
          const before = await get()
          if (index === 5) {
            if (!before.after || before.before.digest === before.after.digest)
              throw new Error('presentation_observation_unavailable')
            if (!same(await stable(before.slideId, before.shapeId, signal), before.after.shape))
              throw new Error('presentation_observation_stale')
          }
          const fresh = async (s?: AbortSignal) => {
            await check(s)
            const current = await observation(
              {
                operation: 'manual_observation_get',
                ...scope,
                observationId: before.observationId,
              },
              s,
            )
            if (!same(current, before)) return false
            if (
              index === 5 &&
              !same(await stable(before.slideId, before.shapeId, s), before.after!.shape)
            )
              return false
            await check(s)
            return true
          }
          const fingerprint = Array.from(
            new Uint8Array(
              await crypto.subtle.digest(
                'SHA-256',
                new TextEncoder().encode(canonicalPresentationValue(before)),
              ),
            ),
            (byte) => byte.toString(16).padStart(2, '0'),
          ).join('')
          await check(signal)
          const proposal = options.proposals.propose({
            operation: name,
            toolName: name,
            title: index === 5 ? '保存观察后的偏好' : '删除编辑观察',
            preview: {
              projectId: scope.projectId,
              observationId: before.observationId,
              slideId: before.slideId,
              shapeId: before.shapeId,
              before: {
                ...before.before,
                shape: {
                  ...before.before.shape,
                  ...(before.before.shape.text !== undefined
                    ? { text: before.before.shape.text.slice(0, 240) }
                    : {}),
                },
              },
              ...(before.after
                ? {
                    after: {
                      ...before.after,
                      shape: {
                        ...before.after.shape,
                        ...(before.after.shape.text !== undefined
                          ? { text: before.after.shape.text.slice(0, 240) }
                          : {}),
                      },
                    },
                  }
                : {}),
              textPreviewTruncated:
                (before.before.shape.text?.length ?? 0) > 240 ||
                (before.after?.shape.text?.length ?? 0) > 240,
              previewCoverage: '文字预览最多各240字符；完整快照请读取观察记录。',
              ...(index === 5 ? { preference: value.preference } : {}),
              disclosure: '差异不证明编辑者身份；偏好仅为用户批准建议，不修改品牌或演示文稿。',
            },
            impact: { host: 'local_preference', targets: [before.observationId], count: 1 },
            fingerprint,
            validate: fresh,
            execute: async (s) => {
              if (!(await fresh(s))) throw new Error('proposal_stale')
              const result = await request(
                {
                  operation:
                    index === 5 ? 'preference_save_observation' : 'manual_observation_delete',
                  ...scope,
                  observationId: before.observationId,
                  expectedBeforeDigest: before.before.digest,
                  expectedAfterDigest: before.after?.digest ?? null,
                  ...(index === 5 ? { text: value.preference } : {}),
                },
                s,
              )
              if (index === 4) {
                if (Object.keys(result).join(',') !== 'deleted' || result.deleted !== true)
                  throw new Error('presentation_response_invalid')
              } else {
                if (Object.keys(result).join(',') !== 'preference')
                  throw new Error('presentation_response_invalid')
                const preference = parseSavedPresentationPreference(result.preference)
                if (
                  preference.projectId !== scope.projectId ||
                  preference.text !== value.preference ||
                  preference.changeId !== `manual_${before.observationId}` ||
                  !same(preference.origin, {
                    version: 1,
                    observationId: before.observationId,
                    beforeDigest: before.before.digest,
                    afterDigest: before.after!.digest,
                  })
                )
                  throw new Error('presentation_response_invalid')
              }
            },
          })
          ownedProposalId = proposal.id
          pendingId = proposal.id
          output = { proposalId: proposal.id, mutated: false }
        }
        await check()
        return {
          output: JSON.stringify(output),
          mutated: false,
          summary: '已读取或提出本机编辑观察操作',
        }
      } catch (error) {
        if (ownedProposalId && options.proposals.pending()?.id === ownedProposalId)
          options.proposals.reject()
        const code = error instanceof Error ? error.message : ''
        return {
          output: JSON.stringify({
            error: [
              'cancelled',
              'invalid_tool_input',
              'presentation_unavailable',
              'presentation_document_changed',
              'presentation_observation_unstable',
              'presentation_observation_stale',
              'presentation_observation_unavailable',
            ].includes(code)
              ? code
              : 'presentation_response_invalid',
          }),
          isError: true,
          mutated: false,
          summary: '编辑观察操作未完成',
        }
      }
    },
  }
}
