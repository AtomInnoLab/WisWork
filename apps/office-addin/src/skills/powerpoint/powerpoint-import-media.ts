import type { AgentSkill, ToolExecution } from '@wiswork/agent-core'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { exactObject, integerField, optionalField, stringField } from '../../agent/tool-schema.js'
import { readBoundedImage } from '../shared/import-media.js'
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
function failed(name: string, error: unknown): ToolExecution {
  const message = error instanceof Error ? error.message : ''
  const code = [
    'invalid_tool_input',
    'image_limit',
    'image_mime_unsupported',
    'invalid_image',
    'vfs_not_found',
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
}): AgentSkill {
  return {
    id: 'office-powerpoint-import-media',
    systemPrompt:
      'Image insertions use bounded VFS media, confirmation, stale-state checks, and semantic verification.',
    tools: [tool],
    async executeTool(call, signal) {
      if (call.inputError || call.truncated)
        return failed(call.name, new Error('invalid_tool_input'))
      try {
        if (signal?.aborted) throw new Error('cancelled')
        if (call.name !== 'insert-image') throw new Error('invalid_tool_input')
        const value = input(call.input)
        if (value.width < 1 || value.height < 1) throw new Error('invalid_tool_input')
        const image = readBoundedImage(options.vfs, value.path)
        const geometry = {
          left: value.left,
          top: value.top,
          width: value.width,
          height: value.height,
        }
        const before = await options.adapter.snapshotSlide(value.slide_index, signal)
        let id: string | undefined
        const proposal = options.proposals.propose({
          operation: call.name,
          toolName: call.name,
          title: value.explanation || 'Insert image',
          preview: {
            path: value.path,
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
            } catch {
              if (s?.aborted) throw new Error('cancelled')
              throw new Error('office_write_failed')
            }
          },
          verify: async (s) => {
            if (!id || !(await options.adapter.verifyImage(value.slide_index, id, geometry, s)))
              throw new Error('office_verify_failed')
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
