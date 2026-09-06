import type { AgentSkill, ToolExecution } from '@wiswork/agent-core'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { exactObject, integerField, optionalField, stringField } from '../../agent/tool-schema.js'
import { readBoundedImage, validateBoundedImageBytes } from '../shared/import-media.js'
import type { InMemoryVfs } from '../shared/vfs.js'

export interface PowerPointImageAdapter {
  snapshotSlide(
    slideIndex: number,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; fingerprint: string }>
  insertImage(
    slideIndex: number,
    base64: string,
    geometry: ImageGeometry,
    signal?: AbortSignal,
  ): Promise<{ id: string }>
  verifyImage(
    slideIndex: number,
    id: string,
    geometry: ImageGeometry,
    signal?: AbortSignal,
  ): Promise<boolean>
  removeImage(slideIndex: number, id: string, signal?: AbortSignal): Promise<void>
  verifyImageAbsent(slideIndex: number, id: string, signal?: AbortSignal): Promise<boolean>
}
export interface ImageGeometry {
  left: number
  top: number
  width: number
  height: number
}
const point = (value: unknown) => {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > 2_000)
    throw new Error('invalid_tool_input')
  return value
}
const input = exactObject({
  path: stringField({ minLength: 1, maxLength: 512 }),
  slide_index: integerField({ min: 0, max: 100_000 }),
  left: point,
  top: point,
  width: point,
  height: point,
  explanation: optionalField(stringField({ maxLength: 100 })),
})
const webInput = exactObject({
  url: stringField({ minLength: 1, maxLength: 2_048 }),
  slide_index: integerField({ min: 0, max: 100_000 }),
  left: point,
  top: point,
  width: point,
  height: point,
  explanation: optionalField(stringField({ maxLength: 100 })),
})
const tool = {
  name: 'insert-image',
  description: 'Propose inserting a bounded VFS PNG or JPEG on a slide.',
  inputSchema: {
    type: 'object',
    properties: {
      path: { type: 'string', maxLength: 512 },
      slide_index: { type: 'integer', minimum: 0, maximum: 100_000 },
      left: { type: 'number', minimum: 0, maximum: 2_000 },
      top: { type: 'number', minimum: 0, maximum: 2_000 },
      width: { type: 'number', minimum: 0, maximum: 2_000 },
      height: { type: 'number', minimum: 0, maximum: 2_000 },
      explanation: { type: 'string', maxLength: 100 },
    },
    required: ['path', 'slide_index', 'left', 'top', 'width', 'height'],
    additionalProperties: false,
  },
}
const webTool = {
  name: 'insert_web_image',
  description: 'Fetch and propose inserting a bounded PNG or JPEG URL returned by image_search.',
  inputSchema: {
    type: 'object',
    properties: {
      url: { type: 'string', maxLength: 2_048 },
      slide_index: { type: 'integer', minimum: 0, maximum: 100_000 },
      left: { type: 'number', minimum: 0, maximum: 2_000 },
      top: { type: 'number', minimum: 0, maximum: 2_000 },
      width: { type: 'number', minimum: 0, maximum: 2_000 },
      height: { type: 'number', minimum: 0, maximum: 2_000 },
      explanation: { type: 'string', maxLength: 100 },
    },
    required: ['url', 'slide_index', 'left', 'top', 'width', 'height'],
    additionalProperties: false,
  },
}
function failed(name: string, error: unknown): ToolExecution {
  const message = error instanceof Error ? error.message : ''
  const code = [
    'invalid_tool_input',
    'image_limit',
    'image_mime_unsupported',
    'invalid_image',
    'vfs_not_found',
    'image_fetch_unavailable',
    'office_api_unsupported',
    'cancelled',
  ].includes(message)
    ? message
    : 'office_operation_failed'
  return { output: code, isError: true, mutated: false, summary: name }
}
export function createPowerPointImportMediaSkill(options: {
  adapter: PowerPointImageAdapter
  proposals: StructuredProposalController
  vfs: InMemoryVfs
  fetchImage?: (url: string, signal?: AbortSignal) => Promise<Uint8Array>
}): AgentSkill {
  return {
    id: 'office-powerpoint-import-media',
    systemPrompt:
      'When insert_web_image is available, use it for an HTTPS image_url returned by image_search; use insert-image only for an attached VFS path. Image insertions use bounded media, the PC-managed PowerPoint session policy, stale-state checks, and semantic verification.',
    tools: options.fetchImage ? [tool, webTool] : [tool],
    async executeTool(call, signal) {
      if (call.inputError || call.truncated)
        return failed(call.name, new Error('invalid_tool_input'))
      try {
        if (signal?.aborted) throw new Error('cancelled')
        if (call.name !== 'insert-image' && call.name !== 'insert_web_image')
          throw new Error('invalid_tool_input')
        const local = call.name === 'insert-image' ? input(call.input) : undefined
        const remote = call.name === 'insert_web_image' ? webInput(call.input) : undefined
        const value = local ?? remote!
        if (value.width < 1 || value.height < 1) throw new Error('invalid_tool_input')
        const image =
          local !== undefined
            ? await readBoundedImage(options.vfs, local.path)
            : await validateBoundedImageBytes(await options.fetchImage!(remote!.url, signal))
        const geometry = {
          left: value.left,
          top: value.top,
          width: value.width,
          height: value.height,
        }
        const before = await options.adapter.snapshotSlide(value.slide_index, signal)
        let id: string | undefined
        const recover = async () => {
          if (!id) throw new Error('office_recovery_failed')
          try {
            await options.adapter.removeImage(value.slide_index, id)
            if (!(await options.adapter.verifyImageAbsent(value.slide_index, id)))
              throw new Error('office_recovery_failed')
          } catch (error) {
            throw new Error('office_recovery_failed', { cause: error })
          }
        }
        const proposal = options.proposals.propose({
          operation: call.name,
          toolName: call.name,
          title: value.explanation || 'Insert image',
          preview: {
            source: local?.path ?? remote!.url,
            mime: image.mime,
            bytes: image.bytes,
            sourceWidth: image.width,
            sourceHeight: image.height,
            ...geometry,
          },
          impact: { host: 'powerpoint', targets: [before.slideId], count: 1 },
          fingerprint: `${before.fingerprint}:${image.fingerprint}`,
          before,
          validate: async (s) =>
            (await options.adapter.snapshotSlide(value.slide_index, s)).fingerprint ===
            before.fingerprint,
          execute: async (s) => {
            if (s?.aborted) throw new Error('cancelled')
            try {
              id = (await options.adapter.insertImage(value.slide_index, image.base64, geometry, s))
                .id
              if (s?.aborted) {
                await recover()
                throw new Error('cancelled')
              }
            } catch (error) {
              if (error instanceof Error && error.message === 'office_recovery_failed') throw error
              throw new Error(s?.aborted ? 'cancelled' : 'office_write_failed', {
                cause: error,
              })
            }
          },
          verify: async (s) => {
            try {
              if (s?.aborted) throw new Error('cancelled')
              if (!id || !(await options.adapter.verifyImage(value.slide_index, id, geometry, s)))
                throw new Error('office_verify_failed')
              if (s?.aborted) throw new Error('cancelled')
            } catch (error) {
              await recover()
              throw new Error(s?.aborted ? 'cancelled' : 'office_verify_failed', {
                cause: error,
              })
            }
          },
        })
        return {
          output: JSON.stringify({ proposalId: proposal.id, mutated: false }),
          mutated: false,
          summary: 'Proposed PowerPoint image insertion',
        }
      } catch (error) {
        return failed(call.name, error)
      }
    },
  }
}
