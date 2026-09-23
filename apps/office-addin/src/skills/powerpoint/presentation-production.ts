import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import { parsePresentationDeck, PRESENTATION_DECK_SCHEMA } from '@wiswork/pptx-engine/presentation'
import type { PresentationGenerationOptions } from './presentation-generation.js'
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
const operations = {
  start_presentation_production: 'production_begin',
  run_presentation_production: 'production_run',
  read_presentation_production: 'production_status',
  read_presentation_page_artifact: 'production_page',
} as const
const idSchema = { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' }
const tools: AgentToolDef[] = Object.keys(operations).map((name) => ({
  name,
  description:
    name === 'start_presentation_production'
      ? 'Freeze a saved plan revision and SlideIR as a durable page compilation task. Reuse request_id for unchanged retries. Does not import slides.'
      : name === 'run_presentation_production'
        ? 'Compile remaining pages of a saved task. Failed pages do not discard successful pages; retry the same request. This is PC preparation, not host delivery or QA.'
        : name === 'read_presentation_production'
          ? 'Read persisted page compilation states and failures; omit request_id for latest.'
          : 'Download one compiled page and its report to session files. Does not import the page or mark QA passed.',
  inputSchema: {
    type: 'object',
    properties:
      name === 'start_presentation_production'
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
      name === 'start_presentation_production'
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
  options: PresentationGenerationOptions,
): AgentSkill & { clear(): void } {
  let epoch = 0
  return {
    id: 'office-presentation-production',
    clear() {
      epoch++
    },
    get tools() {
      return options.available() ? tools : []
    },
    systemPrompt:
      'For page production first save the presentation plan, then start_presentation_production with that plan_revision and matching SlideIR. Run remaining pages with run_presentation_production; inspect failed states and reuse the same request for unchanged retries. Already compiled pages are preserved. Download individual page artifacts only as files: these are not imported, visually reviewed, source-verified or round-trip checked. Never claim the deck is delivered from compiled counts. Do not invent project/request/page IDs.',
    async executeTool(call, signal) {
      const captured = epoch
      let references = false
      const check = () => {
        if (captured !== epoch || signal?.aborted) throw new Error('cancelled')
        if (!options.available()) throw new Error('presentation_unavailable')
        if (references && !options.assetsAvailable?.())
          throw new Error('presentation_assets_unavailable')
      }
      try {
        check()
        const op = operations[call.name as keyof typeof operations],
          input = { ...call.input },
          begin = op === 'production_begin',
          page = op === 'production_page'
        const allowed = begin
          ? ['request_id', 'deck', 'plan_revision']
          : page
            ? ['project_id', 'request_id', 'page_id']
            : ['project_id', 'request_id']
        if (
          !op ||
          call.inputError ||
          call.truncated ||
          Object.keys(input).some((k) => !allowed.includes(k)) ||
          ((op !== 'production_status' || input.request_id !== undefined) &&
            !id(input.request_id)) ||
          (page && !id(input.page_id)) ||
          (begin && !integer(input.plan_revision, 1))
        )
          throw new Error('invalid_tool_input')
        const deck = begin ? parsePresentationDeck(input.deck) : undefined,
          projectId = deck?.id ?? input.project_id
        if (!id(projectId)) throw new Error('invalid_tool_input')
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
        }
        if (new TextEncoder().encode(JSON.stringify(body)).byteLength > 256 * 1024)
          throw new Error('presentation_request_too_large')
        await current()
        const response = await options.request(body, signal)
        await current()
        const text = await response.text()
        await current()
        if (new TextEncoder().encode(text).byteLength > (page ? 15 * 1024 * 1024 : 64 * 1024))
          throw new Error('presentation_response_invalid')
        const value = JSON.parse(text)
        if (value?.error) {
          if (['invalid_request', 'unsupported', 'unsupported_operation'].includes(value.error))
            throw new Error('presentation_upgrade_required')
          if (
            [
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
        let output: unknown, files: [string, string | Uint8Array][] | undefined
        if (page) {
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
            value.requestId !== input.request_id ||
            value.pageId !== input.page_id ||
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
            planRevision: value.planRevision,
            status: 'compiled',
            sourceSlideId: value.sourceSlideId,
            path,
            reportPath,
            report,
          }
        } else {
          const parsed = parsePresentationProductionStatus(value)
          if (
            parsed.projectId !== projectId ||
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
        return {
          output: JSON.stringify(output),
          mutated: false,
          summary: page
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
            code === 'presentation_upgrade_required'
              ? '当前 PC 尚不支持页级生产，请升级 WisWork PC 后重试'
              : '页级生产操作未完成；已保存成果保留，可刷新查看',
        }
      }
    },
  }
}
