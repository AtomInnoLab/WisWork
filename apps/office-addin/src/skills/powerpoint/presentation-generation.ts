import type { CompiledPresentationArtifact } from './presentation-delivery.js'
import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import { parsePresentationDeck, PRESENTATION_DECK_SCHEMA } from '@wiswork/pptx-engine/presentation'
import type { InMemoryVfs } from '../shared/vfs.js'

const validId = (value: unknown): value is string =>
  typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const tools: AgentToolDef[] = [
  {
    name: 'compile_deck_with_pptxgenjs',
    description:
      'For a saved plan supply its plan_revision and exact style, slide order/titles and claim mapping. Compile a planned 16:9 presentation into editable PPTX using the paired PC. Coordinates are inches on a 13.333333 x 7.5 canvas; claimed slides reserve the bottom 0.55 inches for sources. Colors are hex without #. Entire request including inline images must be <=256 KiB; use compact prepared images. Returns downloadable PPTX and a report; it does not modify the open document. Reuse the request_id for unchanged retries; use a new request_id only when the deck changes.',
    inputSchema: {
      type: 'object',
      properties: {
        request_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
        deck: PRESENTATION_DECK_SCHEMA,
        plan_revision: { type: 'integer', minimum: 1 },
      },
      required: ['request_id', 'deck'],
      additionalProperties: false,
    },
  },
  {
    name: 'restore_presentation_project',
    description:
      'Restore the last compiled PPTX for this document from the paired PC after interruption or reopening the Taskpane. An optional project_id selects a known project bound to this document.',
    inputSchema: {
      type: 'object',
      properties: { project_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' } },
      additionalProperties: false,
    },
  },
  {
    name: 'resume_presentation_project',
    description:
      'Resume a persisted compilation request on the paired PC using its original saved input and request ID. This does not import or modify slides. Completed requests return the original artifact. Use the project status to choose the request, never invent an ID.',
    inputSchema: {
      type: 'object',
      properties: {
        project_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
        request_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
      },
      required: ['project_id', 'request_id'],
      additionalProperties: false,
    },
  },
]
export interface PresentationGenerationOptions {
  vfs: InMemoryVfs
  available(): boolean
  assetsAvailable?(): boolean
  attachmentsAvailable?(): boolean
  attachmentsRequest?(body: unknown, signal?: AbortSignal): Promise<Response>
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  documentId(): Promise<string>
  lastProject(): string | undefined
  rememberProject(id: string): Promise<void>
  selectedProduction?(projectId: string, documentId: string): string | undefined
  rememberSelectedProduction?(projectId: string, documentId: string, requestId: string): Promise<void>
}
const abort = (signal?: AbortSignal) => {
  if (signal?.aborted) throw new Error('cancelled')
}

export function createPresentationGenerationSkill(
  options: PresentationGenerationOptions,
): AgentSkill & {
  artifact(projectId?: string): CompiledPresentationArtifact | undefined
  clear(): void
} {
  const artifacts = new Map<string, CompiledPresentationArtifact>()
  let epoch = 0
  return {
    artifact: (projectId?: string) =>
      artifacts.get(projectId ?? [...artifacts.keys()].at(-1) ?? ''),
    clear: () => {
      epoch += 1
      artifacts.clear()
    },
    id: 'office-presentation-generation',
    systemPrompt: `For a new presentation, read user materials first, establish evidence and the per-page story, then choose a consistent visual style. When presentation compilation is available, use compile_deck_with_pptxgenjs with validated SlideIR to create a downloadable native PPTX. Prefer native charts/tables for factual data. Assets can be prepared inline PNG/JPEG images, never paths or external URLs. When presentation-assets.v1 is available, prefer compact attachment references from list_presentation_attachments; the PC resolves cached image data. Do not invent sources. Preserve project ID and request ID on unchanged retries. Compiled is not visually reviewed: explain checks marked not_run or not_verified and use existing Office tools for subsequent editing and host verification. The PPTX and report are available in Session attachments. Restore a prior compiled result with restore_presentation_project.`,
    get tools() {
      return options.available() ? tools : []
    },
    buildContext: () =>
      options.available()
        ? `Presentation compilation available.${options.lastProject() ? ` Last project: ${options.lastProject()}.` : ''}`
        : '',
    async executeTool(call, signal) {
      const captured = epoch
      let usesAssetReferences = false
      const check = () => {
        abort(signal)
        if (captured !== epoch) throw new Error('cancelled')
        if (!options.available()) throw new Error('presentation_unavailable')
        if (usesAssetReferences && !options.assetsAvailable?.())
          throw new Error('presentation_assets_unavailable')
      }
      try {
        check()
        if (!options.available()) throw new Error('presentation_unavailable')
        if (call.inputError || call.truncated) throw new Error('invalid_tool_input')
        const value = call.input
        let projectId: string
        let requestId: string | undefined
        let deck: ReturnType<typeof parsePresentationDeck> | undefined
        if (call.name === 'compile_deck_with_pptxgenjs') {
          if (
            Object.keys(value).some(
              (key) => !['request_id', 'deck', 'plan_revision'].includes(key),
            ) ||
            !validId(value.request_id) ||
            (value.plan_revision !== undefined &&
              (!Number.isSafeInteger(value.plan_revision) || Number(value.plan_revision) < 1))
          )
            throw new Error('invalid_tool_input')
          deck = parsePresentationDeck(value.deck)
          usesAssetReferences = deck.assets.some((asset) => 'attachmentId' in asset)
          check()
          projectId = deck.id
          requestId = value.request_id
        } else if (call.name === 'resume_presentation_project') {
          if (
            Object.keys(value).some((key) => !['project_id', 'request_id'].includes(key)) ||
            !validId(value.project_id) ||
            !validId(value.request_id)
          )
            throw new Error('invalid_tool_input')
          projectId = value.project_id
          requestId = value.request_id
        } else if (call.name === 'restore_presentation_project') {
          if (Object.keys(value).some((key) => key !== 'project_id'))
            throw new Error('invalid_tool_input')
          const id = value.project_id ?? options.lastProject()
          if (!validId(id)) throw new Error('presentation_project_missing')
          projectId = id
        } else throw new Error('invalid_tool_input')
        const documentId = await options.documentId()
        check()
        const body = deck
          ? {
              operation: 'compile',
              documentId,
              projectId,
              requestId,
              deck,
              ...(value.plan_revision !== undefined ? { planRevision: value.plan_revision } : {}),
            }
          : requestId
            ? { operation: 'resume', documentId, projectId, requestId }
            : { operation: 'get', documentId, projectId }
        if (new TextEncoder().encode(JSON.stringify(body)).byteLength > 256 * 1024)
          throw new Error('presentation_request_too_large')
        // Remember before dispatch: a lost first response must not orphan durable PC work.
        if (deck) await options.rememberProject(projectId)
        check()
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed')
        check()
        const response = await options.request(body, signal)
        check()
        if (!response.ok) throw new Error('presentation_service_unavailable')
        const text = await response.text()
        if (text.length > 16 * 1024 * 1024) throw new Error('presentation_response_invalid')
        const result = JSON.parse(text)
        if (
          result &&
          [
            'invalid_request',
            'invalid_plan',
            'revision_conflict',
            'plan_mismatch',
            'invalid_deck',
            'invalid_state',
            'document_mismatch',
            'request_conflict',
            'not_found',
            'aborted',
            'output_too_large',
            'compile_failed',
          ].includes(result.error)
        )
          throw new Error(`presentation_${result.error}`)
        if (
          !result ||
          result.status !== 'compiled' ||
          result.projectId !== projectId ||
          !validId(result.requestId) ||
          (requestId && result.requestId !== requestId) ||
          result.report?.deckId !== projectId ||
          !Number.isSafeInteger(result.report?.slideCount) ||
          result.report.slideCount < 1 ||
          (deck && result.report.slideCount !== deck.slides.length) ||
          typeof result.pptxBase64 !== 'string' ||
          result.pptxBase64.length > 14 * 1024 * 1024 ||
          !/^[A-Za-z0-9+/]+={0,2}$/.test(result.pptxBase64)
        )
          throw new Error('presentation_response_invalid')
        let pages: Array<{ id: string; title: string; sourceSlideId: string }> | undefined
        if (result.pages !== undefined) {
          if (
            !Array.isArray(result.pages) ||
            result.pages.length !== result.report.slideCount ||
            result.pages.length > 32
          )
            throw new Error('presentation_response_invalid')
          const parsedPages: NonNullable<typeof pages> = result.pages.map(
            (page: unknown, index: number) => {
              if (!page || typeof page !== 'object' || Array.isArray(page))
                throw new Error('presentation_response_invalid')
              const value = page as { id: string; title: string; sourceSlideId: string }
              if (
                Object.keys(value).some((key) => !['id', 'title', 'sourceSlideId'].includes(key)) ||
                !validId(value.id) ||
                typeof value.title !== 'string' ||
                !value.title ||
                value.title.length > 300 ||
                typeof value.sourceSlideId !== 'string' ||
                !/^[1-9][0-9]{0,9}#$/.test(value.sourceSlideId) ||
                Number(value.sourceSlideId.slice(0, -1)) < 256 ||
                Number(value.sourceSlideId.slice(0, -1)) > 4294967295 ||
                (deck &&
                  (value.id !== deck.slides[index]?.id ||
                    value.title !== deck.slides[index]?.title))
              )
                throw new Error('presentation_response_invalid')
              return Object.freeze({ ...value })
            },
          )
          pages = parsedPages
          if (
            new Set(pages.map((page) => page.id)).size !== pages.length ||
            new Set(pages.map((page) => page.sourceSlideId)).size !== pages.length
          )
            throw new Error('presentation_response_invalid')
          Object.freeze(pages)
        }
        const binary = atob(result.pptxBase64)
        if (binary.length < 4 || binary.slice(0, 4) !== 'PK\u0003\u0004')
          throw new Error('presentation_response_invalid')
        const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0))
        check()
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed')
        check()
        if (!deck) {
          await options.rememberProject(projectId)
          check()
          if ((await options.documentId()) !== documentId)
            throw new Error('presentation_document_changed')
          check()
        }
        const path = `/home/user/generated/${projectId}.pptx`
        const reportPath = `/home/user/generated/${projectId}.report.json`
        options.vfs.writeBatch([
          [path, bytes],
          [reportPath, JSON.stringify(result.report, null, 2)],
        ])
        artifacts.delete(projectId)
        artifacts.set(
          projectId,
          Object.freeze({
            documentId,
            projectId,
            requestId: result.requestId,
            pptxBase64: result.pptxBase64,
            slideCount: result.report.slideCount,
            ...(pages ? { pages } : {}),
          }),
        )
        // Keep the in-memory import cache bounded; persisted projects remain restorable.
        while (artifacts.size > 4) artifacts.delete(artifacts.keys().next().value!)
        return {
          output: JSON.stringify({
            projectId,
            requestId: result.requestId,
            status: 'compiled',
            path,
            reportPath,
            report: result.report,
          }),
          mutated: false,
          summary: `已生成 ${result.report.slideCount} 页可编辑 PPTX，可在附件中下载；请查看待验收项`,
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : ''
        const code =
          message === 'cancelled' ||
          message === 'invalid_tool_input' ||
          /^presentation_[a-z_]{1,80}$/.test(message)
            ? message
            : 'presentation_operation_failed'
        return {
          output: code,
          isError: true,
          mutated: false,
          summary:
            code === 'presentation_assets_unavailable'
              ? '请更新并连接支持图片素材的 PC 端后重试'
              : 'PPT 生成未完成，已有成果已保留',
        }
      }
    },
  }
}
