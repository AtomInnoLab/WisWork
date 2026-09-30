import type { AgentSkill } from '@wiswork/agent-core'
import { presentationImportContent } from '@wiswork/project-store/presentation-import-source'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import { MAX_PPTX_IMPORT_PAGE_BYTES, presentationPackageDigest } from './powerpoint-package.js'
import { equivalentNativePresentationPage } from './presentation-page-equivalence.js'
import { validatePresentationImportSourcePage } from './presentation-import-source-package.js'
import type {
  CompiledPresentationArtifact,
  PresentationDeliveryOptions,
  PresentationImportCheckpoint,
  PresentationImportRecord,
} from './presentation-delivery.js'
export interface PresentationImportProgress {
  source?: 'production'
  projectId?: string
  requestId?: string
  total: number
  completed: number
  status: 'not_started' | 'partial' | 'uncertain' | 'complete'
  pages: {
    id: string
    title: string
    state: 'pending' | 'complete' | 'uncertain'
    slideId?: string
    completedAt?: string
    startedAt?: string
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
const validTimestamp = (value: unknown): value is string =>
  typeof value === 'string' &&
  /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(value) &&
  Number.isFinite(Date.parse(value)) &&
  new Date(value).toISOString() === value
export function validPresentationImportRecord(value: unknown): value is PresentationImportRecord {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const r = value as PresentationImportRecord
  if (
    Object.keys(r).some(
      (k) =>
        !['state', 'documentId', 'toolCallId', 'agentRunId', 'slideIds', 'checkpoint'].includes(k),
    ) ||
    !['pending', 'complete'].includes(r.state) ||
    typeof r.documentId !== 'string' ||
    !r.documentId ||
    r.documentId.length > 4096 ||
    (r.toolCallId !== undefined &&
      (typeof r.toolCallId !== 'string' || r.toolCallId.length < 1 || r.toolCallId.length > 256)) ||
    (r.agentRunId !== undefined &&
      (typeof r.agentRunId !== 'string' ||
        !/^[A-Za-z0-9_-]{1,128}$/.test(r.agentRunId) ||
        !r.toolCallId)) ||
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
          ...(c.version === 2 ? ['pageIds'] : []),
        ].includes(k),
    ) ||
    ![1, 2].includes(c.version) ||
    !/^[a-f0-9]{64}$/.test(c.artifactDigest) ||
    !Array.isArray(c.sourceSlideIds) ||
    c.sourceSlideIds.length < 1 ||
    c.sourceSlideIds.length > 32 ||
    !Array.from(c.sourceSlideIds).every(validSourceSlideId) ||
    (c.version === 1 && new Set(c.sourceSlideIds).size !== c.sourceSlideIds.length) ||
    (c.version === 2 &&
      (!Array.isArray(c.pageIds) ||
        c.pageIds.length !== c.sourceSlideIds.length ||
        Array.from(c.pageIds).some(
          (id) => typeof id !== 'string' || !/^[A-Za-z0-9_-]{1,80}$/.test(id),
        ) ||
        new Set(c.pageIds).size !== c.pageIds.length)) ||
    !hostIds(c.baselineSlideIds) ||
    !Array.isArray(c.completed) ||
    c.completed.length > c.sourceSlideIds.length
  )
    return false
  for (const [index, page] of c.completed.entries()) {
    if (
      !page ||
      typeof page !== 'object' ||
      Object.keys(page).some(
        (k) => !['sourceSlideId', 'slideId', 'completedAt', 'startedAt'].includes(k),
      ) ||
      page.sourceSlideId !== c.sourceSlideIds[index] ||
      typeof page.slideId !== 'string' ||
      !page.slideId ||
      page.slideId.length > 256 ||
      (index > 0 &&
        Boolean(c.completed[index - 1]?.completedAt) &&
        page.completedAt === undefined) ||
      (Object.hasOwn(page, 'startedAt') &&
        (!validTimestamp(page.startedAt) ||
          !validTimestamp(page.completedAt) ||
          page.startedAt > page.completedAt ||
          (index > 0 &&
            c.completed[index - 1]?.completedAt &&
            page.startedAt < c.completed[index - 1]!.completedAt!))) ||
      (page.completedAt !== undefined &&
        (!validTimestamp(page.completedAt) ||
          (index > 0 &&
            c.completed[index - 1]?.completedAt &&
            page.completedAt < c.completed[index - 1]!.completedAt!)))
    )
      return false
  }
  if (!hostIds([...c.baselineSlideIds, ...c.completed.map((p) => p.slideId)], 1032)) return false
  if (
    c.inFlight !== undefined &&
    (!c.inFlight ||
      typeof c.inFlight !== 'object' ||
      Object.keys(c.inFlight).some((key) => !['sourceSlideId', 'startedAt'].includes(key)) ||
      !validSourceSlideId(c.inFlight.sourceSlideId) ||
      c.inFlight.sourceSlideId !== c.sourceSlideIds[c.completed.length] ||
      (c.inFlight.startedAt !== undefined &&
        (!validTimestamp(c.inFlight.startedAt) ||
          (c.completed.at(-1)?.completedAt &&
            c.inFlight.startedAt < c.completed.at(-1)!.completedAt!))))
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
  const production = artifact.pagePptxBase64 !== undefined
  if (production) {
    if (
      !Array.isArray(artifact.pagePptxBase64) ||
      artifact.pagePptxBase64.length !== artifact.slideCount ||
      artifact.pptxBase64 !== '' ||
      !Number.isSafeInteger(artifact.planRevision) ||
      artifact.planRevision! < 1
    )
      return false
    let bytes = 0
    for (const base64 of artifact.pagePptxBase64) {
      if (
        typeof base64 !== 'string' ||
        !base64.length ||
        base64.length > Math.ceil(MAX_PPTX_IMPORT_PAGE_BYTES / 3) * 4
      )
        return false
      try {
        const binary = atob(base64)
        if (btoa(binary) !== base64) return false
        bytes += binary.length
      } catch {
        return false
      }
      if (bytes > MAX_PPTX_IMPORT_PAGE_BYTES) return false
    }
  }
  return Boolean(
    artifact.pages &&
    artifact.pages.length === artifact.slideCount &&
    artifact.pages.length > 0 &&
    artifact.pages.length <= 32 &&
    Array.from(artifact.pages).every(
      (p) =>
        p &&
        typeof p.id === 'string' &&
        /^[A-Za-z0-9_-]{1,80}$/.test(p.id) &&
        typeof p.title === 'string' &&
        p.title.length <= 300 &&
        validSourceSlideId(p.sourceSlideId),
    ) &&
    new Set(artifact.pages.map((p) => p.id)).size === artifact.pages.length &&
    (production ||
      new Set(artifact.pages.map((p) => p.sourceSlideId)).size === artifact.pages.length),
  )
}
// Keep v1 byte hashing and v2 bundle serialization identical to persisted import receipts.
export function presentationArtifactContent(artifact: CompiledPresentationArtifact): string {
  if (artifact.pagePptxBase64 === undefined) return artifact.pptxBase64
  if (!validPages(artifact)) throw new Error('presentation_import_state_invalid')
  presentationImportKey(artifact)
  return presentationImportContent(artifact)
}
export function presentationPageMapping(
  artifact: CompiledPresentationArtifact,
  record: PresentationImportRecord | undefined,
  pageId: string,
): { sourceSlideId: string; slideId: string } | undefined {
  if (!record?.checkpoint || !summarizePresentationImport(artifact, record)) return undefined
  const index = artifact.pages!.findIndex((page) => page.id === pageId)
  const completed = index < 0 ? undefined : record.checkpoint.completed[index]
  return completed
    ? { sourceSlideId: completed.sourceSlideId, slideId: completed.slideId }
    : undefined
}
export function presentationImportKey(artifact: CompiledPresentationArtifact): string {
  if (
    typeof artifact.projectId !== 'string' ||
    typeof artifact.requestId !== 'string' ||
    !/^[A-Za-z0-9_-]{1,128}$/.test(artifact.projectId) ||
    !/^[A-Za-z0-9_-]{1,128}$/.test(artifact.requestId)
  )
    throw new Error('presentation_import_state_invalid')
  return `${artifact.pagePptxBase64 !== undefined ? 'production/' : ''}${artifact.projectId}/${artifact.requestId}`
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
  const production = artifact.pagePptxBase64 !== undefined
  if (record && (production ? checkpoint?.version !== 2 : checkpoint?.version === 2))
    return undefined
  if (
    checkpoint?.version === 2 &&
    !same(
      checkpoint.pageIds,
      artifact.pages!.map((page) => page.id),
    )
  )
    return undefined
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
    ...(production ? { source: 'production' as const } : {}),
    projectId: artifact.projectId,
    requestId: artifact.requestId,
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
      ...(index < completed && checkpoint?.completed[index]?.completedAt
        ? { completedAt: checkpoint.completed[index].completedAt }
        : {}),
      ...(index < completed && checkpoint?.completed[index]?.startedAt
        ? { startedAt: checkpoint.completed[index].startedAt }
        : {}),
      ...(uncertain && index === completed && checkpoint?.inFlight?.startedAt
        ? { startedAt: checkpoint.inFlight.startedAt }
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
  return createPageDelivery(options, false)
}
export function createPresentationProductionDeliverySkill(
  options: PresentationDeliveryOptions,
): AgentSkill {
  return createPageDelivery(options, true)
}
function createPageDelivery(options: PresentationDeliveryOptions, production: boolean): AgentSkill {
  const importName = production ? 'import_presentation_production' : 'import_generated_presentation'
  const reconcileName = 'reconcile_presentation_production_import'
  const readTool = {
    ...tool,
    name: production ? 'read_presentation_production_import_status' : tool.name,
  }
  return {
    id: production
      ? 'office-presentation-production-delivery'
      : 'office-presentation-page-delivery',
    get tools() {
      if (production && (!options.available() || !options.adapter.available())) return []
      if (
        production &&
        (!options.adapter.exportPage || options.adapter.supportsPageExport?.() === false)
      )
        return [readTool]
      return production
        ? [
            {
              ...readTool,
              name: importName,
              description:
                'Propose appending the prepared ordered single-page PPTX production outputs. Confirm once; interrupted pages use durable checkpoints. Import is not QA.',
            },
            readTool,
            {
              ...readTool,
              name: reconcileName,
              description:
                'Read the host page and complete an interrupted production-page checkpoint only after exact-package or conservative native text/shape/picture equivalence is proven.',
            },
          ]
        : [readTool]
    },
    systemPrompt:
      'Read saved page import progress after interruption. For an uncertain production page, try reconcile_presentation_production_import to read the appended host page and compare it with the prepared source. Exact package equality or conservative native text/shape/picture equivalence may prove the page; unsupported structures remain uncertain for human inspection. Resume only the remaining pages after the confirmed completed prefix, and only when there is no uncertain page. Never repeat completed pages. Never delete or replay an uncertain page automatically.',
    async executeTool(call, signal) {
      try {
        const read = call.name === readTool.name
        const reconcile = production && call.name === reconcileName
        if (
          (!read && !reconcile && call.name !== importName) ||
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
        if (
          production &&
          !read &&
          (!options.adapter.exportPage || options.adapter.supportsPageExport?.() === false)
        )
          throw new Error('presentation_unavailable')
        const artifact = options.artifact(call.input.project_id as string | undefined)
        if (!artifact) throw new Error('presentation_restore_required')
        if (!validPages(artifact) || production !== (artifact.pagePptxBase64 !== undefined))
          throw new Error('presentation_import_state_invalid')
        const artifactBase64 = artifact.pptxBase64
        const artifactPages = JSON.stringify(artifact.pages)
        const pageBytes = artifact.pagePptxBase64?.slice()
        const identity = JSON.stringify([
          artifact.documentId,
          artifact.projectId,
          artifact.requestId,
          artifact.planRevision,
          artifact.slideCount,
        ])
        const documentId = await options.documentId()
        const checkArtifact = () => {
          if (
            artifact.pptxBase64 !== artifactBase64 ||
            JSON.stringify(artifact.pages) !== artifactPages ||
            (production &&
              (JSON.stringify([
                artifact.documentId,
                artifact.projectId,
                artifact.requestId,
                artifact.planRevision,
                artifact.slideCount,
              ]) !== identity ||
                !Array.isArray(artifact.pagePptxBase64) ||
                artifact.pagePptxBase64.length !== pageBytes!.length ||
                pageBytes!.some((bytes, index) => bytes !== artifact.pagePptxBase64![index])))
          )
            throw new Error('presentation_import_state_invalid')
        }
        const current = async (checkSignal?: AbortSignal) => {
          checkArtifact()
          if (
            checkSignal?.aborted ||
            !options.available() ||
            options.artifact(artifact.projectId) !== artifact
          )
            throw new Error('cancelled')
          if ((await options.documentId()) !== documentId || documentId !== artifact.documentId)
            throw new Error('presentation_document_changed')
          checkArtifact()
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
            await crypto.subtle.digest(
              'SHA-256',
              new TextEncoder().encode(presentationArtifactContent(artifact)),
            ),
          ),
          (b) => b.toString(16).padStart(2, '0'),
        ).join('')
        await current(signal)
        const key = presentationImportKey(artifact),
          previous = options.readReceipt(key),
          progress = summarizePresentationImport(artifact, previous)
        if (!progress || (previous?.checkpoint && previous.checkpoint.artifactDigest !== digest))
          throw new Error('presentation_import_state_invalid')
        if (reconcile) {
          const checkpoint = previous?.checkpoint
          if (!checkpoint?.inFlight || previous?.state !== 'pending' || !options.adapter.exportPage)
            throw new Error('presentation_import_uncertain')
          const expectedIds = [
            ...checkpoint.baselineSlideIds,
            ...checkpoint.completed.map((page) => page.slideId),
          ]
          const before = await options.adapter.snapshot(signal)
          await current(signal)
          if (
            before.fingerprint !== JSON.stringify(before.slideIds) ||
            before.slideIds.length !== expectedIds.length + 1 ||
            expectedIds.some((id, index) => before.slideIds[index] !== id)
          )
            throw new Error('presentation_import_uncertain')
          const candidate = before.slideIds.at(-1)!
          const source = pageBytes?.[checkpoint.completed.length]
          if (!source) throw new Error('presentation_import_state_invalid')
          const exported = await options.adapter.exportPage(candidate, signal)
          await current(signal)
          if (!(await equivalentNativePresentationPage(source, exported)))
            throw new Error('presentation_import_uncertain')
          const after = await options.adapter.snapshot(signal)
          await current(signal)
          if (!same(after, before) || !same(options.readReceipt(key), previous))
            throw new Error('presentation_import_uncertain')
          const completed = [
            ...checkpoint.completed,
            {
              sourceSlideId: checkpoint.inFlight.sourceSlideId,
              slideId: candidate,
              ...(checkpoint.inFlight.startedAt
                ? { startedAt: checkpoint.inFlight.startedAt }
                : {}),
              completedAt: new Date(
                Math.max(
                  Date.now(),
                  Date.parse(checkpoint.inFlight.startedAt ?? '') || 0,
                  Date.parse(checkpoint.completed.at(-1)?.completedAt ?? '') || 0,
                ),
              ).toISOString(),
            },
          ]
          const next: PresentationImportRecord = {
            ...previous,
            state: completed.length === checkpoint.sourceSlideIds.length ? 'complete' : 'pending',
            checkpoint: { ...checkpoint, completed, inFlight: undefined },
            ...(completed.length === checkpoint.sourceSlideIds.length
              ? { slideIds: completed.map((page) => page.slideId) }
              : {}),
          }
          await options.writeReceipt(key, next)
          await current()
          if (!same(options.readReceipt(key), next))
            throw new Error('presentation_import_uncertain')
          return {
            output: JSON.stringify(summarizePresentationImport(artifact, next)),
            mutated: false,
            summary: `已核对并认领第 ${completed.length} 页导入；未完成视觉验收`,
          }
        }
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
        if (production) {
          for (
            let index = previous?.checkpoint?.completed.length ?? 0;
            index < pageBytes!.length;
            index++
          ) {
            try {
              await presentationPackageDigest(pageBytes![index]!, signal, 'import_page')
              await validatePresentationImportSourcePage(
                pageBytes![index]!,
                artifact.pages![index]!.sourceSlideId,
              )
            } catch (error) {
              if (error instanceof Error && error.message === 'cancelled') throw error
              throw new Error('presentation_import_state_invalid', { cause: error })
            }
          }
          await current(signal)
        }
        const before = await options.adapter.snapshot(signal)
        await current(signal)
        if (
          !hostIds(before.slideIds, previous ? 1032 : 1000) ||
          before.fingerprint !== JSON.stringify(before.slideIds)
        )
          throw new Error('office_state_uncertain')
        const initial: PresentationImportCheckpoint = previous?.checkpoint ?? {
          version: production ? 2 : 1,
          ...(production ? { pageIds: artifact.pages!.map((page) => page.id) } : {}),
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
          operation: importName,
          toolName: importName,
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
                toolCallId: call.id,
                checkpoint: c,
                ...(state === 'complete' ? { slideIds: c.completed.map((p) => p.slideId) } : {}),
              }
              await current()
              await options.writeReceipt(key, record)
              await current()
              last = options.readReceipt(key)
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
              const startedAt = new Date(
                Math.max(
                  Date.now(),
                  Date.parse(checkpoint.completed.at(-1)?.completedAt ?? '') || 0,
                ),
              ).toISOString()
              await save({ ...checkpoint, inFlight: { sourceSlideId, startedAt } })
              let receipt
              try {
                await current(s)
                receipt = await options.adapter.insertPage!(
                  pageBytes?.[checkpoint.completed.length] ?? artifact.pptxBase64,
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
                  let unchangedHost = false
                  try {
                    const after = await options.adapter.snapshot()
                    unchangedHost =
                      same(after.slideIds, slideIds) &&
                      after.fingerprint === baseline.fingerprint &&
                      same(options.readReceipt(key), last)
                  } catch {
                    // A failed read cannot prove that Office did not append the page.
                  }
                  if (unchangedHost) {
                    await current()
                    await save(checkpoint)
                  }
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
              // A host ACK proves the appended ID, not that native media and objects survived.
              // Keep the in-flight checkpoint if the exported production page differs.
              if (production) {
                try {
                  const exported = await options.adapter.exportPage!(receipt.slideIds[0]!)
                  if (
                    !(await equivalentNativePresentationPage(
                      pageBytes![checkpoint.completed.length]!,
                      exported,
                    ))
                  )
                    throw new Error('office_state_uncertain')
                } catch (error) {
                  throw new Error('office_state_uncertain', { cause: error })
                }
                const afterExport = await options.adapter.snapshot()
                if (!same(afterExport.slideIds, [...slideIds, receipt.slideIds[0]!]))
                  throw new Error('office_state_uncertain')
              }
              await current()
              checkpoint = {
                ...checkpoint,
                completed: [
                  ...checkpoint.completed,
                  {
                    sourceSlideId,
                    slideId: receipt.slideIds[0]!,
                    startedAt,
                    completedAt: new Date(
                      Math.max(
                        Date.now(),
                        Date.parse(startedAt),
                        Date.parse(checkpoint.completed.at(-1)?.completedAt ?? '') || 0,
                      ),
                    ).toISOString(),
                  },
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
