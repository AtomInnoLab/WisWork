import type {
  CompiledPresentationArtifact,
  PresentationImportRecord,
} from './presentation-delivery.js'
import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import { parsePresentationDeck, PRESENTATION_DECK_SCHEMA } from '@wiswork/pptx-engine/presentation'
import type { PresentationGenerationOptions } from './presentation-generation.js'
import {
  presentationArtifactContent,
  presentationImportKey,
  presentationPageMapping,
  validPresentationImportRecord,
} from './presentation-page-delivery.js'
const id = (v: unknown): v is string => typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
const integer = (v: unknown, min = 0) => Number.isSafeInteger(v) && Number(v) >= min
const errors = [
  'compile_failed',
  'invalid_deck',
  'aborted',
  'output_too_large',
  'asset_unavailable',
] as const
export interface PresentationProductionStatus {
  revision?: { parentRequestId: string; pageId: string; parentInputDigest: string }
  projectId: string
  requestId: string
  planRevision: number
  status: 'pending' | 'building' | 'partial' | 'compiled'
  compiledCount: number
  total: number
  pages: {
    id: string
    title: string
    state: 'pending' | 'building' | 'compiled' | 'failed'
    attempt: number
    error?: (typeof errors)[number]
  }[]
}
function object(v: unknown, keys: string[]): v is Record<string, unknown> {
  return (
    !!v &&
    typeof v === 'object' &&
    !Array.isArray(v) &&
    Object.keys(v).every((k) => keys.includes(k))
  )
}
export function parsePresentationProductionStatus(value: unknown): PresentationProductionStatus {
  if (
    !object(value, [
      'revision',
      'projectId',
      'requestId',
      'planRevision',
      'status',
      'compiledCount',
      'total',
      'pages',
    ]) ||
    !id(value.projectId) ||
    !id(value.requestId) ||
    !integer(value.planRevision, 1) ||
    !integer(value.total, 1) ||
    Number(value.total) > 32 ||
    !integer(value.compiledCount) ||
    !Array.isArray(value.pages) ||
    value.pages.length !== value.total
  )
    throw new Error('presentation_response_invalid')
  for (const page of value.pages) {
    if (
      !object(page, ['id', 'title', 'state', 'attempt', 'error']) ||
      !id(page.id) ||
      typeof page.title !== 'string' ||
      !page.title ||
      page.title.length > 300 ||
      !['pending', 'building', 'compiled', 'failed'].includes(String(page.state)) ||
      !integer(page.attempt) ||
      (page.state === 'pending' ? page.attempt !== 0 : Number(page.attempt) < 1) ||
      (page.state === 'failed'
        ? !errors.includes(page.error as (typeof errors)[number])
        : page.error !== undefined)
    )
      throw new Error('presentation_response_invalid')
  }
  const p = value as unknown as PresentationProductionStatus
  if (
    p.revision !== undefined &&
    (!object(p.revision, ['parentRequestId', 'pageId', 'parentInputDigest']) ||
      !id(p.revision.parentRequestId) ||
      p.revision.parentRequestId === p.requestId ||
      !id(p.revision.pageId) ||
      !p.pages.some((page) => page.id === p.revision!.pageId) ||
      typeof p.revision.parentInputDigest !== 'string' ||
      !/^[a-f0-9]{64}$/.test(p.revision.parentInputDigest))
  )
    throw new Error('presentation_response_invalid')
  const count = p.pages.filter((x) => x.state === 'compiled').length
  const state =
    count === p.total
      ? 'compiled'
      : p.pages.some((x) => x.state === 'building')
        ? 'building'
        : count || p.pages.some((x) => x.state === 'failed')
          ? 'partial'
          : 'pending'
  if (
    new Set(p.pages.map((x) => x.id)).size !== p.total ||
    p.compiledCount !== count ||
    p.status !== state
  )
    throw new Error('presentation_response_invalid')
  return structuredClone(p)
}
export function parsePresentationPageArtifact(
  value: unknown,
  projectId: string,
  requestId: string,
  pageId: string,
) {
  if (
    !object(value, [
      'projectId',
      'requestId',
      'pageId',
      'planRevision',
      'status',
      'pptxBase64',
      'sourceSlideId',
      'report',
    ]) ||
    value.projectId !== projectId ||
    value.requestId !== requestId ||
    value.pageId !== pageId ||
    value.status !== 'compiled' ||
    !integer(value.planRevision, 1) ||
    typeof value.sourceSlideId !== 'string' ||
    !/^[1-9][0-9]{0,9}#$/.test(value.sourceSlideId) ||
    Number(value.sourceSlideId.slice(0, -1)) < 256 ||
    Number(value.sourceSlideId.slice(0, -1)) > 4294967295 ||
    typeof value.pptxBase64 !== 'string' ||
    value.pptxBase64.length > 14 * 1024 * 1024 ||
    value.pptxBase64.length % 4 !== 0 ||
    !/^[A-Za-z0-9+/]+={0,2}$/.test(value.pptxBase64) ||
    !value.report ||
    typeof value.report !== 'object' ||
    Array.isArray(value.report)
  )
    throw new Error('presentation_response_invalid')
  const report = value.report as Record<string, unknown>
  if (
    report.deckId !== projectId ||
    report.slideCount !== 1 ||
    new TextEncoder().encode(JSON.stringify(report)).byteLength > 48 * 1024
  )
    throw new Error('presentation_response_invalid')
  const binary = atob(value.pptxBase64)
  if (
    binary.slice(0, 4) !== 'PK\u0003\u0004' ||
    binary.length > 10 * 1024 * 1024 ||
    btoa(binary) !== value.pptxBase64
  )
    throw new Error('presentation_response_invalid')
  return {
    report,
    binary,
    pptxBase64: value.pptxBase64,
    sourceSlideId: value.sourceSlideId,
    planRevision: value.planRevision as number,
  }
}
const operations = {
  rebuild_presentation_page: 'production_rebuild_page',
  prepare_presentation_production_import: 'production_status',
  start_presentation_production: 'production_begin',
  run_presentation_production: 'production_run',
  read_presentation_production: 'production_status',
  read_presentation_page_artifact: 'production_page',
} as const
const idSchema = { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' }
const tools: AgentToolDef[] = Object.keys(operations).map((name) => ({
  name,
  description:
    name === 'rebuild_presentation_page'
      ? 'Create a derived production task by changing one SlideIR page from a fully compiled parent. Reuse the frozen plan, title, claims, style and registered assets. Does not run compilation or replace a host page. Derived tasks cannot be bulk imported; download the changed page for inspection.'
      : name === 'prepare_presentation_production_import'
        ? 'Prepare all compiled pages of one exact production request for separately confirmed import. All pages must be compiled; preserves order and verifies the shared plan revision. Keeps only the latest prepared project in memory, within a 10 MiB decoded budget. Does not insert slides or perform QA.'
        : name === 'start_presentation_production'
          ? 'Freeze a saved plan revision and SlideIR as a durable page compilation task. Reuse request_id for unchanged retries. Does not import slides.'
          : name === 'run_presentation_production'
            ? 'Compile remaining pages of a saved task. Failed pages do not discard successful pages; retry the same request. This is PC preparation, not host delivery or QA.'
            : name === 'read_presentation_production'
              ? 'Read persisted page compilation states and failures; omit request_id for latest.'
              : 'Download one compiled page and its report to session files. Does not import the page or mark QA passed.',
  inputSchema: {
    type: 'object',
    properties:
      name === 'rebuild_presentation_page'
        ? {
            project_id: idSchema,
            parent_request_id: idSchema,
            request_id: idSchema,
            page_id: idSchema,
            slide: (
              PRESENTATION_DECK_SCHEMA as unknown as {
                properties: { slides: { items: Record<string, unknown> } }
              }
            ).properties.slides.items,
          }
        : name === 'start_presentation_production'
          ? {
              request_id: idSchema,
              deck: PRESENTATION_DECK_SCHEMA,
              plan_revision: { type: 'integer', minimum: 1 },
            }
          : {
              project_id: idSchema,
              request_id: idSchema,
              ...(name === 'read_presentation_page_artifact' ? { page_id: idSchema } : {}),
            },
    required:
      name === 'rebuild_presentation_page'
        ? ['project_id', 'parent_request_id', 'request_id', 'page_id', 'slide']
        : name === 'start_presentation_production'
          ? ['request_id', 'deck', 'plan_revision']
          : name === 'read_presentation_production'
            ? ['project_id']
            : name === 'read_presentation_page_artifact'
              ? ['project_id', 'request_id', 'page_id']
              : ['project_id', 'request_id'],
    additionalProperties: false,
  },
}))
export function createPresentationProductionSkill(
  options: PresentationGenerationOptions & {
    readReceipt?(key: string): PresentationImportRecord | undefined
  },
): AgentSkill & {
  clear(): void
  artifact(projectId?: string): CompiledPresentationArtifact | undefined
} {
  let epoch = 0
  let prepareSequence = 0
  let artifact: CompiledPresentationArtifact | undefined
  return {
    id: 'office-presentation-production',
    artifact: (projectId) =>
      !projectId || artifact?.projectId === projectId ? artifact : undefined,
    clear() {
      epoch++
      artifact = undefined
    },
    get tools() {
      return options.available() ? tools : []
    },
    systemPrompt:
      'For page production first save the presentation plan, then start_presentation_production with that plan_revision and matching SlideIR. Run remaining pages with run_presentation_production; inspect failed states and reuse the same request for unchanged retries. Already compiled pages are preserved. Use prepare_presentation_production_import only after all pages compile to prepare a bounded ordered collection for separately confirmed import; it replaces the previous prepared collection but never inserts slides. Download individual page artifacts only as files: these are not imported, visually reviewed, source-verified or round-trip checked. Never claim the deck is delivered from compiled counts. Do not invent project/request/page IDs. rebuild_presentation_page creates a derived task only; run it separately to compile the changed page. Use the confirmed page replacement tools for host replacement. Preparing a derived task requires its already committed complete business mapping; it never authorizes bulk append. After commit prepare the child; after undo prepare the parent before editing or QA.',
    async executeTool(call, signal) {
      const captured = epoch
      let preparation: number | undefined
      let references = false
      const check = () => {
        if (
          captured !== epoch ||
          signal?.aborted ||
          (preparation !== undefined && preparation !== prepareSequence)
        )
          throw new Error('cancelled')
        if (!options.available()) throw new Error('presentation_unavailable')
        if (references && !options.assetsAvailable?.())
          throw new Error('presentation_assets_unavailable')
      }
      try {
        check()
        const op = operations[call.name as keyof typeof operations],
          input = { ...call.input },
          begin = op === 'production_begin',
          rebuild = op === 'production_rebuild_page',
          page = op === 'production_page',
          prepare = call.name === 'prepare_presentation_production_import'
        const allowed = rebuild
          ? ['project_id', 'parent_request_id', 'request_id', 'page_id', 'slide']
          : begin
            ? ['request_id', 'deck', 'plan_revision']
            : page
              ? ['project_id', 'request_id', 'page_id']
              : ['project_id', 'request_id']
        if (
          !op ||
          call.inputError ||
          call.truncated ||
          Object.keys(input).some((k) => !allowed.includes(k)) ||
          ((op !== 'production_status' || prepare || input.request_id !== undefined) &&
            !id(input.request_id)) ||
          (page && !id(input.page_id)) ||
          (rebuild &&
            (!id(input.parent_request_id) ||
              input.parent_request_id === input.request_id ||
              !id(input.page_id) ||
              !object(input.slide, ['id', 'title', 'elements', 'claimIds', 'notes']) ||
              input.slide.id !== input.page_id)) ||
          (begin && !integer(input.plan_revision, 1))
        )
          throw new Error('invalid_tool_input')
        const deck = begin ? parsePresentationDeck(input.deck) : undefined,
          projectId = deck?.id ?? input.project_id
        if (!id(projectId)) throw new Error('invalid_tool_input')
        if (prepare) preparation = ++prepareSequence
        const slide = rebuild ? structuredClone(input.slide) : undefined
        references = !!deck?.assets.some((a) => 'attachmentId' in a)
        check()
        const documentId = await options.documentId()
        check()
        const current = async () => {
          check()
          if ((await options.documentId()) !== documentId)
            throw new Error('presentation_document_changed')
          check()
        }
        const body = {
          operation: op,
          documentId,
          projectId,
          ...(input.request_id ? { requestId: input.request_id } : {}),
          ...(page ? { pageId: input.page_id } : {}),
          ...(deck ? { deck, planRevision: input.plan_revision } : {}),
          ...(rebuild
            ? { parentRequestId: input.parent_request_id, pageId: input.page_id, slide }
            : {}),
        }
        const fetchResponse = async (requestBody: unknown, isPage: boolean) => {
          if (new TextEncoder().encode(JSON.stringify(requestBody)).byteLength > 256 * 1024)
            throw new Error('presentation_request_too_large')
          await current()
          const response = await options.request(requestBody, signal)
          await current()
          const text = await response.text()
          await current()
          if (new TextEncoder().encode(text).byteLength > (isPage ? 15 * 1024 * 1024 : 64 * 1024))
            throw new Error('presentation_response_invalid')
          const value = JSON.parse(text)
          if (value?.error) {
            if (['invalid_request', 'unsupported', 'unsupported_operation'].includes(value.error))
              throw new Error('presentation_upgrade_required')
            if (
              [
                'page_not_ready',
                'not_found',
                'document_mismatch',
                'request_conflict',
                'plan_mismatch',
                'revision_conflict',
                'invalid_deck',
                'aborted',
                'output_too_large',
                'compile_failed',
                'asset_unavailable',
                'invalid_state',
              ].includes(value.error)
            )
              throw new Error(`presentation_${value.error}`)
            throw new Error('presentation_response_invalid')
          }
          if (!response.ok) throw new Error('presentation_service_unavailable')
          return value
        }
        const value = await fetchResponse(body, page)
        await current()
        let output: unknown, files: [string, string | Uint8Array][] | undefined
        let prepared: CompiledPresentationArtifact | undefined
        let preparedReceiptJson: string | undefined
        let preparedImported = false
        if (prepare) {
          const status = parsePresentationProductionStatus(value)
          if (status.projectId !== projectId || status.requestId !== input.request_id)
            throw new Error('presentation_response_invalid')
          const receiptKey = `production/${projectId}/${status.requestId}`
          const savedReceipt = options.readReceipt?.(receiptKey)
          const savedReceiptJson = JSON.stringify(savedReceipt)
          preparedReceiptJson = savedReceiptJson
          preparedImported = savedReceipt?.state === 'complete'
          if (
            status.revision &&
            (!validPresentationImportRecord(savedReceipt) ||
              savedReceipt.state !== 'complete' ||
              savedReceipt.checkpoint?.version !== 2)
          )
            throw new Error('presentation_page_replacement_required')
          if (status.status !== 'compiled') throw new Error('presentation_production_not_ready')
          const pages: NonNullable<CompiledPresentationArtifact['pages']> = [],
            pagePptxBase64: string[] = []
          let decodedBytes = 0
          for (const planned of status.pages) {
            const response = await fetchResponse(
              {
                operation: 'production_page',
                documentId,
                projectId,
                requestId: status.requestId,
                pageId: planned.id,
              },
              true,
            )
            await current()
            const result = parsePresentationPageArtifact(
              response,
              projectId,
              status.requestId,
              planned.id,
            )
            if (result.planRevision !== status.planRevision)
              throw new Error('presentation_response_invalid')
            decodedBytes += result.binary.length
            if (decodedBytes > 10 * 1024 * 1024) throw new Error('presentation_output_too_large')
            pagePptxBase64.push(result.pptxBase64)
            pages.push({
              id: planned.id,
              title: planned.title,
              sourceSlideId: result.sourceSlideId,
            })
          }
          prepared = {
            documentId,
            projectId,
            requestId: status.requestId,
            slideCount: status.total,
            pptxBase64: '',
            pages,
            pagePptxBase64,
            planRevision: status.planRevision,
          }
          if (savedReceipt !== undefined) {
            const artifactDigest = Array.from(
              new Uint8Array(
                await crypto.subtle.digest(
                  'SHA-256',
                  new TextEncoder().encode(presentationArtifactContent(prepared)),
                ),
              ),
              (b) => b.toString(16).padStart(2, '0'),
            ).join('')
            await current()
            if (
              !validPresentationImportRecord(savedReceipt) ||
              savedReceipt.documentId !== documentId ||
              savedReceipt.checkpoint?.version !== 2 ||
              savedReceipt.checkpoint.artifactDigest !== artifactDigest ||
              JSON.stringify(savedReceipt.checkpoint.pageIds) !==
                JSON.stringify(pages.map((p) => p.id)) ||
              JSON.stringify(savedReceipt.checkpoint.sourceSlideIds) !==
                JSON.stringify(pages.map((p) => p.sourceSlideId)) ||
              (savedReceipt.state === 'complete' &&
                pages.some((p) => !presentationPageMapping(prepared!, savedReceipt, p.id)))
            )
              throw new Error('presentation_page_binding_invalid')
          }
          await current()
          if (
            JSON.stringify(options.readReceipt?.(presentationImportKey(prepared))) !==
            savedReceiptJson
          )
            throw new Error('presentation_page_binding_invalid')
          pages.forEach(Object.freeze)
          Object.freeze(pages)
          Object.freeze(pagePptxBase64)
          Object.freeze(prepared)
          output = {
            projectId,
            requestId: status.requestId,
            planRevision: status.planRevision,
            status: 'prepared',
            slideCount: status.total,
            pages,
          }
        } else if (page) {
          const { report, binary, sourceSlideId, planRevision } = parsePresentationPageArtifact(
            value,
            projectId,
            input.request_id as string,
            input.page_id as string,
          )
          const hash = Array.from(
            new Uint8Array(
              await crypto.subtle.digest(
                'SHA-256',
                new TextEncoder().encode(
                  JSON.stringify([documentId, projectId, input.request_id, input.page_id]),
                ),
              ),
            ),
            (b) => b.toString(16).padStart(2, '0'),
          ).join('')
          await current()
          const path = `/home/user/generated/page-${hash}.pptx`,
            reportPath = `/home/user/generated/page-${hash}.report.json`
          files = [
            [path, Uint8Array.from(binary, (c) => c.charCodeAt(0))],
            [reportPath, JSON.stringify(report, null, 2)],
          ]
          output = {
            projectId,
            requestId: input.request_id,
            pageId: input.page_id,
            planRevision,
            status: 'compiled',
            sourceSlideId,
            path,
            reportPath,
            report,
          }
        } else {
          const parsed = parsePresentationProductionStatus(value)
          if (
            parsed.projectId !== projectId ||
            (rebuild &&
              (!parsed.revision ||
                parsed.revision.parentRequestId !== input.parent_request_id ||
                parsed.revision.pageId !== input.page_id)) ||
            (input.request_id && parsed.requestId !== input.request_id) ||
            (deck &&
              (parsed.planRevision !== input.plan_revision ||
                JSON.stringify(parsed.pages.map((p) => ({ id: p.id, title: p.title }))) !==
                  JSON.stringify(deck.slides.map((p) => ({ id: p.id, title: p.title })))))
          )
            throw new Error('presentation_response_invalid')
          output = parsed
        }
        await current()
        await options.rememberProject(projectId)
        await current()
        if (files) options.vfs.writeBatch(files)
        if (
          prepared &&
          JSON.stringify(options.readReceipt?.(presentationImportKey(prepared))) !==
            preparedReceiptJson
        )
          throw new Error('presentation_page_binding_invalid')
        if (prepared) artifact = prepared // Publish only the complete collection; one project keeps the total cache bounded.
        return {
          output: JSON.stringify(output),
          mutated: false,
          summary: rebuild
            ? '已创建单页派生任务；尚未运行编译或替换宿主页'
            : prepare
              ? preparedImported
                ? '已恢复已导入页面的产物与映射；验收需另行执行'
                : '已准备逐页导入成果；尚未插入文稿或验收'
              : page
                ? '已下载单页编译成果；尚未导入或验收'
                : '已读取页级编译进度；编译成功不代表导入或验收完成',
        }
      } catch (error) {
        const raw = error instanceof Error ? error.message : '',
          code =
            raw === 'cancelled' ||
            raw === 'invalid_tool_input' ||
            /^presentation_[a-z_]{1,80}$/.test(raw)
              ? raw
              : 'presentation_operation_failed'
        return {
          output: code,
          isError: true,
          mutated: false,
          summary:
            code === 'presentation_page_replacement_required'
              ? '当前修订尚无已提交页面映射，不能整批追加导入；请先完成页面替换'
              : code === 'presentation_upgrade_required'
                ? '当前 PC 尚不支持页级生产，请升级 WisWork PC 后重试'
                : '页级生产操作未完成；已保存成果保留，可刷新查看',
        }
      }
    },
  }
}
