import type { AgentSkill } from '@wiswork/agent-core'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import type {
  CompiledPresentationArtifact,
  PresentationDeliveryOptions,
  PresentationImportCheckpoint,
  PresentationImportRecord,
} from './presentation-delivery.js'
export interface PresentationImportProgress {
  total: number
  completed: number
  status: 'not_started' | 'partial' | 'uncertain' | 'complete'
  pages: {
    id: string
    title: string
    state: 'pending' | 'complete' | 'uncertain'
    slideId?: string
  }[]
}
export const validSourceSlideId = (value: unknown): value is string =>
  typeof value === 'string' &&
  /^[1-9][0-9]{0,9}#$/.test(value) &&
  Number(value.slice(0, -1)) >= 256 &&
  Number(value.slice(0, -1)) <= 4294967295
const hostIds = (value: unknown, max = 1000): value is string[] =>
  Array.isArray(value) &&
  value.length <= max &&
  value.every((id) => typeof id === 'string' && id.length > 0 && id.length <= 256) &&
  new Set(value).size === value.length
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
export function validPresentationImportRecord(value: unknown): value is PresentationImportRecord {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const r = value as PresentationImportRecord
  if (
    Object.keys(r).some((k) => !['state', 'documentId', 'slideIds', 'checkpoint'].includes(k)) ||
    !['pending', 'complete'].includes(r.state) ||
    typeof r.documentId !== 'string' ||
    !r.documentId ||
    r.documentId.length > 4096 ||
    (r.slideIds !== undefined && !hostIds(r.slideIds, 100)) ||
    (r.state === 'complete' && !r.slideIds)
  )
    return false
  const c = r.checkpoint
  if (c === undefined) return true
  if (
    !c ||
    typeof c !== 'object' ||
    Object.keys(c).some(
      (k) =>
        ![
          'version',
          'artifactDigest',
          'sourceSlideIds',
          'baselineSlideIds',
          'completed',
          'inFlight',
        ].includes(k),
    ) ||
    c.version !== 1 ||
    !/^[a-f0-9]{64}$/.test(c.artifactDigest) ||
    !Array.isArray(c.sourceSlideIds) ||
    c.sourceSlideIds.length < 1 ||
    c.sourceSlideIds.length > 32 ||
    !c.sourceSlideIds.every(validSourceSlideId) ||
    new Set(c.sourceSlideIds).size !== c.sourceSlideIds.length ||
    !hostIds(c.baselineSlideIds) ||
    !Array.isArray(c.completed) ||
    c.completed.length > c.sourceSlideIds.length
  )
    return false
  for (const [index, page] of c.completed.entries()) {
    if (
      !page ||
      typeof page !== 'object' ||
      Object.keys(page).some((k) => !['sourceSlideId', 'slideId'].includes(k)) ||
      page.sourceSlideId !== c.sourceSlideIds[index] ||
      typeof page.slideId !== 'string' ||
      !page.slideId ||
      page.slideId.length > 256
    )
      return false
  }
  if (!hostIds([...c.baselineSlideIds, ...c.completed.map((p) => p.slideId)], 1032)) return false
  if (
    c.inFlight !== undefined &&
    (!c.inFlight ||
      typeof c.inFlight !== 'object' ||
      Object.keys(c.inFlight).length !== 1 ||
      !validSourceSlideId(c.inFlight.sourceSlideId) ||
      c.inFlight.sourceSlideId !== c.sourceSlideIds[c.completed.length])
  )
    return false
  if (
    r.state === 'complete' &&
    (c.inFlight ||
      c.completed.length !== c.sourceSlideIds.length ||
      !same(
        r.slideIds,
        c.completed.map((p) => p.slideId),
      ))
  )
    return false
  return (
    r.state !== 'pending' ||
    (r.slideIds === undefined && c.completed.length < c.sourceSlideIds.length)
  )
}
function validPages(artifact: CompiledPresentationArtifact): boolean {
  return Boolean(
    artifact.pages &&
    artifact.pages.length === artifact.slideCount &&
    artifact.pages.length > 0 &&
    artifact.pages.length <= 32 &&
    artifact.pages.every(
      (p) =>
        p &&
        typeof p.id === 'string' &&
        /^[A-Za-z0-9_-]{1,80}$/.test(p.id) &&
        typeof p.title === 'string' &&
        p.title.length <= 300 &&
        validSourceSlideId(p.sourceSlideId),
    ) &&
    new Set(artifact.pages.map((p) => p.id)).size === artifact.pages.length &&
    new Set(artifact.pages.map((p) => p.sourceSlideId)).size === artifact.pages.length,
  )
}
export function summarizePresentationImport(
  artifact: CompiledPresentationArtifact,
  record: PresentationImportRecord | undefined,
): PresentationImportProgress | undefined {
  if (
    !validPages(artifact) ||
    (record &&
      (!validPresentationImportRecord(record) || record.documentId !== artifact.documentId))
  )
    return undefined
  if (record?.state === 'complete' && record.slideIds?.length !== artifact.slideCount)
    return undefined
  const checkpoint = record?.checkpoint
  if (
    checkpoint &&
    !same(
      checkpoint.sourceSlideIds,
      artifact.pages!.map((p) => p.sourceSlideId),
    )
  )
    return undefined
  const completed =
    record?.state === 'complete' ? artifact.slideCount : (checkpoint?.completed.length ?? 0)
  const uncertain = record?.state === 'pending' && (!checkpoint || Boolean(checkpoint.inFlight))
  return {
    total: artifact.slideCount,
    completed,
    status:
      record?.state === 'complete'
        ? 'complete'
        : uncertain
          ? 'uncertain'
          : completed
            ? 'partial'
            : 'not_started',
    pages: artifact.pages!.map((page, index) => ({
      id: page.id,
      title: page.title,
      state:
        index < completed ? 'complete' : uncertain && index === completed ? 'uncertain' : 'pending',
      ...(index < completed
        ? { slideId: checkpoint?.completed[index]?.slideId ?? record?.slideIds?.[index] }
        : {}),
    })),
  }
}
const tool = {
  name: 'read_presentation_import_status',
  description:
    'Read durable per-page import checkpoints for a compiled or restored presentation. Completed means page IDs were verified, not visual QA. An uncertain in-flight page must never be replayed automatically.',
  inputSchema: {
    type: 'object',
    properties: { project_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' } },
    additionalProperties: false,
  },
}
export function createPresentationPageDeliverySkill(
  options: PresentationDeliveryOptions,
): AgentSkill {
  return {
    id: 'office-presentation-page-delivery',
    tools: [tool],
    systemPrompt:
      'Read saved page import progress after interruption. Resume only confirmed completed pages with no uncertain page. Never delete or replay an uncertain page automatically.',
    async executeTool(call, signal) {
      try {
        const read = call.name === tool.name
        if (
          (!read && call.name !== 'import_generated_presentation') ||
          call.inputError ||
          call.truncated ||
          Object.keys(call.input).some((k) => k !== 'project_id') ||
          (call.input.project_id !== undefined &&
            (typeof call.input.project_id !== 'string' ||
              !/^[A-Za-z0-9_-]{1,128}$/.test(call.input.project_id)))
        )
          throw new Error('invalid_tool_input')
        if (signal?.aborted) throw new Error('cancelled')
        if (!options.available() || !options.adapter.available())
          throw new Error('presentation_unavailable')
        const artifact = options.artifact(call.input.project_id as string | undefined)
        if (!artifact) throw new Error('presentation_restore_required')
        if (!validPages(artifact)) throw new Error('presentation_import_state_invalid')
        const artifactBase64 = artifact.pptxBase64
        const artifactPages = JSON.stringify(artifact.pages)
        const documentId = await options.documentId()
        const current = async (checkSignal?: AbortSignal) => {
          if (
            artifact.pptxBase64 !== artifactBase64 ||
            JSON.stringify(artifact.pages) !== artifactPages
          )
            throw new Error('presentation_import_state_invalid')
          if (
            checkSignal?.aborted ||
            !options.available() ||
            options.artifact(artifact.projectId) !== artifact
          )
            throw new Error('cancelled')
          if ((await options.documentId()) !== documentId || documentId !== artifact.documentId)
            throw new Error('presentation_document_changed')
          if (
            checkSignal?.aborted ||
            !options.available() ||
            options.artifact(artifact.projectId) !== artifact
          )
            throw new Error('cancelled')
        }
        await current(signal)
        const digest = Array.from(
          new Uint8Array(
            await crypto.subtle.digest('SHA-256', new TextEncoder().encode(artifact.pptxBase64)),
          ),
          (b) => b.toString(16).padStart(2, '0'),
        ).join('')
        await current(signal)
        const key = `${artifact.projectId}/${artifact.requestId}`,
          previous = options.readReceipt(key),
          progress = summarizePresentationImport(artifact, previous)
        if (!progress || (previous?.checkpoint && previous.checkpoint.artifactDigest !== digest))
          throw new Error('presentation_import_state_invalid')
        if (read)
          return {
            output: JSON.stringify(progress),
            mutated: false,
            summary: `已保存 ${progress.completed}/${progress.total} 页导入进度；未完成视觉验收`,
          }
        if (previous?.state === 'complete')
          return {
            output: JSON.stringify({ status: 'already_imported', slideIds: previous.slideIds }),
            mutated: false,
            summary: '这份文稿已导入，未重复插入页面',
          }
        if (previous && (!previous.checkpoint || previous.checkpoint.inFlight))
          throw new Error('presentation_import_uncertain')
        if (!options.adapter.insertPage) throw new Error('presentation_unavailable')
        const before = await options.adapter.snapshot(signal)
        await current(signal)
        if (
          !hostIds(before.slideIds, previous ? 1032 : 1000) ||
          before.fingerprint !== JSON.stringify(before.slideIds)
        )
          throw new Error('office_state_uncertain')
        const initial: PresentationImportCheckpoint = previous?.checkpoint ?? {
          version: 1,
          artifactDigest: digest,
          sourceSlideIds: artifact.pages!.map((p) => p.sourceSlideId),
          baselineSlideIds: before.slideIds,
          completed: [],
        }
        const expected = (c: PresentationImportCheckpoint) => [
          ...c.baselineSlideIds,
          ...c.completed.map((p) => p.slideId),
        ]
        if (!same(before.slideIds, expected(initial)))
          throw new Error('presentation_import_host_changed')
        const previousRaw = JSON.stringify(previous)
        const unchanged = () => JSON.stringify(options.readReceipt(key)) === previousRaw
        let last: PresentationImportRecord | undefined
        const proposal = options.proposals.propose({
          operation: 'import_generated_presentation',
          toolName: 'import_generated_presentation',
          title: `添加剩余 ${artifact.slideCount - initial.completed.length} 页，保留已完成页面`,
          preview: {
            project: artifact.projectId,
            pages: artifact.slideCount,
            completed: initial.completed.length,
          },
          impact: {
            host: 'powerpoint',
            targets: before.slideIds.length ? [before.slideIds.at(-1)!] : ['presentation'],
            count: artifact.slideCount - initial.completed.length,
          },
          fingerprint: selectionFingerprint(`${documentId}:${before.fingerprint}:${key}:${digest}`),
          validate: async (s) => {
            try {
              await current(s)
              if (!unchanged()) return false
              const snapshot = await options.adapter.snapshot(s)
              await current(s)
              return (
                same(snapshot.slideIds, expected(initial)) &&
                snapshot.fingerprint === JSON.stringify(snapshot.slideIds)
              )
            } catch {
              return false
            }
          },
          execute: async (s) => {
            await current(s)
            if (!unchanged()) throw new Error('proposal_stale')
            let checkpoint = structuredClone(initial)
            const save = async (
              c: PresentationImportCheckpoint,
              state: 'pending' | 'complete' = 'pending',
            ) => {
              const record: PresentationImportRecord = {
                state,
                documentId,
                checkpoint: c,
                ...(state === 'complete' ? { slideIds: c.completed.map((p) => p.slideId) } : {}),
              }
              await current()
              await options.writeReceipt(key, record)
              await current()
              last = record
            }
            if (!previous) await save(checkpoint)
            while (checkpoint.completed.length < checkpoint.sourceSlideIds.length) {
              await current(s)
              if (last && !same(options.readReceipt(key), last)) throw new Error('proposal_stale')
              const slideIds = expected(checkpoint),
                baseline = { slideIds, fingerprint: JSON.stringify(slideIds) }
              const actual = await options.adapter.snapshot(s)
              await current(s)
              if (!same(actual.slideIds, slideIds) || actual.fingerprint !== baseline.fingerprint)
                throw new Error('proposal_stale')
              const sourceSlideId = checkpoint.sourceSlideIds[checkpoint.completed.length]!
              await save({ ...checkpoint, inFlight: { sourceSlideId } })
              let receipt
              try {
                await current(s)
                receipt = await options.adapter.insertPage!(
                  artifact.pptxBase64,
                  sourceSlideId,
                  baseline,
                  s,
                )
              } catch (error) {
                if (
                  error instanceof Error &&
                  [
                    'cancelled',
                    'proposal_stale',
                    'office_api_unsupported',
                    'invalid_tool_input',
                  ].includes(error.message)
                ) {
                  await current()
                  await save(checkpoint)
                }
                throw error
              }
              await current()
              if (
                !hostIds(receipt.slideIds, 1) ||
                receipt.slideIds.length !== 1 ||
                slideIds.includes(receipt.slideIds[0]!) ||
                !(await options.adapter.verify(receipt, baseline))
              )
                throw new Error('office_state_uncertain')
              await current()
              checkpoint = {
                ...checkpoint,
                completed: [
                  ...checkpoint.completed,
                  { sourceSlideId, slideId: receipt.slideIds[0]! },
                ],
              }
              await save(
                checkpoint,
                checkpoint.completed.length === checkpoint.sourceSlideIds.length
                  ? 'complete'
                  : 'pending',
              )
            }
          },
          verify: async () => {
            await current()
            if (!last || last.state !== 'complete' || !same(options.readReceipt(key), last))
              throw new Error('office_state_uncertain')
            const snapshot = await options.adapter.snapshot()
            await current()
            if (!same(snapshot.slideIds, expected(last.checkpoint!)))
              throw new Error('office_state_uncertain')
          },
        })
        return {
          output: JSON.stringify({
            proposalId: proposal.id,
            status: 'awaiting_confirmation',
            completed: initial.completed.length,
            total: artifact.slideCount,
          }),
          mutated: false,
          summary: '已准备继续添加剩余页面，等待确认',
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : ''
        return {
          output: /^(presentation_[a-z_]+|office_[a-z_]+|invalid_tool_input|cancelled)$/.test(
            message,
          )
            ? message
            : 'presentation_operation_failed',
          isError: true,
          mutated: false,
          summary: '页面导入未完成，已保存的进度和现有页面已保留',
        }
      }
    },
  }
}
