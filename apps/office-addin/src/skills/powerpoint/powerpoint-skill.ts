import type { AgentSkill, ToolExecution } from '@wiswork/agent-core'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { exactObject, integerField, optionalField, stringField } from '../../agent/tool-schema.js'
import type { PowerPointAdapter } from './browser-powerpoint-adapter.js'
import { MAX_POWERPOINT_RESULT_BYTES } from './browser-powerpoint-adapter.js'

const MAX_SLIDE_INDEX = 100_000
const MAX_CODE = 32 * 1024
const MAX_SCREENSHOT_BYTES = 4 * 1024 * 1024
const slideInput = exactObject({
  slide_index: integerField({ min: 0, max: MAX_SLIDE_INDEX }),
  explanation: optionalField(stringField({ maxLength: 50 })),
})
const shapeInput = exactObject({
  slide_index: integerField({ min: 0, max: MAX_SLIDE_INDEX }),
  shape_id: stringField({ minLength: 1, maxLength: 256 }),
  explanation: optionalField(stringField({ maxLength: 50 })),
})
const verifyInput = exactObject({ explanation: optionalField(stringField({ maxLength: 50 })) })
const textEditInput = exactObject({
  slide_index: integerField({ min: 0, max: MAX_SLIDE_INDEX }),
  shape_id: stringField({ minLength: 1, maxLength: 256 }),
  text: stringField({ maxLength: 12_000 }),
  explanation: optionalField(stringField({ maxLength: 50 })),
})
const codeInput = exactObject({
  code: stringField({ minLength: 1, maxLength: MAX_CODE }),
  explanation: optionalField(stringField({ maxLength: 100 })),
})
const masterCodeInput = exactObject({
  code: stringField({ minLength: 1, maxLength: MAX_CODE }),
  explanation: optionalField(stringField({ maxLength: 50 })),
})
const slideCodeInput = exactObject({
  slide_index: integerField({ min: 0, max: MAX_SLIDE_INDEX }),
  code: stringField({ minLength: 1, maxLength: MAX_CODE }),
  explanation: optionalField(stringField({ maxLength: 50 })),
})

const slideProperties = {
  slide_index: { type: 'integer', minimum: 0, maximum: MAX_SLIDE_INDEX },
  explanation: { type: 'string', maxLength: 50 },
} as const
const tools = [
  {
    name: 'screenshot_slide',
    description:
      'Take a bounded PNG screenshot for the task-pane UI and return model-visible MIME, byte count, and fingerprint metadata.',
    inputSchema: {
      type: 'object',
      properties: slideProperties,
      required: ['slide_index'],
      additionalProperties: false,
    },
  },
  {
    name: 'list_slide_shapes',
    description: 'List stable shape IDs, types, and geometry on one slide.',
    inputSchema: {
      type: 'object',
      properties: slideProperties,
      required: ['slide_index'],
      additionalProperties: false,
    },
  },
  {
    name: 'read_slide_text',
    description: 'Read bounded text from a shape selected by stable ID.',
    inputSchema: {
      type: 'object',
      properties: {
        ...slideProperties,
        shape_id: { type: 'string', minLength: 1, maxLength: 256 },
      },
      required: ['slide_index', 'shape_id'],
      additionalProperties: false,
    },
  },
  {
    name: 'verify_slides',
    description: 'Check bounded slides for negative, out-of-bounds, and overlapping geometry.',
    inputSchema: {
      type: 'object',
      properties: { explanation: { type: 'string', maxLength: 50 } },
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: 'execute_office_js',
    description:
      'Raw Office.js compatibility entry; disabled until an audited hardened evaluator is available.',
    inputSchema: {
      type: 'object',
      properties: {
        code: { type: 'string', minLength: 1, maxLength: MAX_CODE },
        explanation: { type: 'string', maxLength: 100 },
      },
      required: ['code'],
      additionalProperties: false,
    },
  },
  {
    name: 'edit_slide_text',
    description: 'Propose replacing the text of one shape.',
    inputSchema: {
      type: 'object',
      properties: {
        ...slideProperties,
        shape_id: { type: 'string', minLength: 1, maxLength: 256 },
        text: { type: 'string', maxLength: 12_000 },
      },
      required: ['slide_index', 'shape_id', 'text'],
      additionalProperties: false,
    },
  },
  {
    name: 'edit_slide_xml',
    description: 'OOXML compatibility entry; disabled pending audited ZIP support.',
    inputSchema: {
      type: 'object',
      properties: {
        ...slideProperties,
        code: { type: 'string', minLength: 1, maxLength: MAX_CODE },
      },
      required: ['slide_index', 'code'],
      additionalProperties: false,
    },
  },
  {
    name: 'edit_slide_chart',
    description: 'Chart OOXML compatibility entry; disabled pending audited ZIP support.',
    inputSchema: {
      type: 'object',
      properties: {
        ...slideProperties,
        code: { type: 'string', minLength: 1, maxLength: MAX_CODE },
      },
      required: ['slide_index', 'code'],
      additionalProperties: false,
    },
  },
  {
    name: 'edit_slide_master',
    description: 'Master OOXML compatibility entry; disabled pending audited ZIP support.',
    inputSchema: {
      type: 'object',
      properties: {
        code: { type: 'string', minLength: 1, maxLength: MAX_CODE },
        explanation: { type: 'string', maxLength: 50 },
      },
      required: ['code'],
      additionalProperties: false,
    },
  },
  {
    name: 'duplicate_slide',
    description: 'Propose duplicating a slide immediately after its source.',
    inputSchema: {
      type: 'object',
      properties: slideProperties,
      required: ['slide_index'],
      additionalProperties: false,
    },
  },
] as const

function failure(name: string, code: string): ToolExecution {
  return { output: code, isError: true, mutated: false, summary: name }
}
function assertNotCancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new Error('cancelled')
}
function boundedJson(value: unknown): string {
  const result = JSON.stringify(value)
  if (new TextEncoder().encode(result).byteLength > MAX_POWERPOINT_RESULT_BYTES)
    throw new Error('office_read_failed')
  return result
}
function errorCode(error: unknown, write = false): string {
  const code = error instanceof Error ? error.message : ''
  if (['invalid_tool_input', 'office_api_unsupported', 'cancelled'].includes(code)) return code
  if (code === 'office_verify_failed') return code
  return write ? 'office_write_failed' : 'office_read_failed'
}
function validPng(value: unknown): value is string {
  if (
    typeof value !== 'string' ||
    value.length === 0 ||
    value.length > Math.ceil(MAX_SCREENSHOT_BYTES / 3) * 4 ||
    value.length % 4 !== 0 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(value) ||
    !value.startsWith('iVBORw0KGgo')
  )
    return false
  const padding = value.endsWith('==') ? 2 : value.endsWith('=') ? 1 : 0
  return (value.length / 4) * 3 - padding <= MAX_SCREENSHOT_BYTES
}

function base64Bytes(value: string): number {
  const padding = value.endsWith('==') ? 2 : value.endsWith('=') ? 1 : 0
  return (value.length / 4) * 3 - padding
}

function fingerprint(value: string): string {
  let result = 0x811c9dc5
  for (let index = 0; index < value.length; index += 1) {
    result ^= value.charCodeAt(index)
    result = Math.imul(result, 0x01000193)
  }
  return `${value.length}:${(result >>> 0).toString(16).padStart(8, '0')}`
}

export function createPowerPointSkill(options: {
  adapter: PowerPointAdapter
  proposals: StructuredProposalController
}): AgentSkill {
  return {
    id: 'office-powerpoint',
    systemPrompt:
      'PowerPoint reads are bounded. Every supported write creates an explicit proposal and is verified after confirmation. Raw JavaScript and unaudited OOXML operations fail closed.',
    tools: [...tools],
    async executeTool(call, signal) {
      if (call.inputError || call.truncated) return failure(call.name, 'invalid_tool_input')
      try {
        assertNotCancelled(signal)
        if (call.name === 'screenshot_slide') {
          const input = slideInput(call.input)
          const result = await options.adapter.screenshotSlide(input.slide_index, signal)
          assertNotCancelled(signal)
          if (result.mime !== 'image/png' || !validPng(result.base64))
            throw new Error('office_read_failed')
          return {
            output: boundedJson({
              mime: result.mime,
              bytes: base64Bytes(result.base64),
              fingerprint: fingerprint(result.base64),
              visualAvailableToModel: false,
            }),
            display: {
              kind: 'images',
              items: [{ url: `data:${result.mime};base64,${result.base64}` }],
            },
            mutated: false,
            summary: 'Rendered PowerPoint slide',
          }
        }
        if (call.name === 'list_slide_shapes') {
          const input = slideInput(call.input)
          return {
            output: boundedJson(await options.adapter.listSlideShapes(input.slide_index, signal)),
            mutated: false,
            summary: 'Listed PowerPoint shapes',
          }
        }
        if (call.name === 'read_slide_text') {
          const input = shapeInput(call.input)
          return {
            output: boundedJson(
              await options.adapter.readSlideText(input.slide_index, input.shape_id, signal),
            ),
            mutated: false,
            summary: 'Read PowerPoint text',
          }
        }
        if (call.name === 'verify_slides') {
          verifyInput(call.input)
          return {
            output: boundedJson(await options.adapter.verifySlides(signal)),
            mutated: false,
            summary: 'Verified PowerPoint slides',
          }
        }
        if (call.name === 'edit_slide_text') {
          const input = textEditInput(call.input)
          await options.adapter.verifySlides(signal)
          const before = await options.adapter.readSlideText(
            input.slide_index,
            input.shape_id,
            signal,
          )
          assertNotCancelled(signal)
          const snapshot = await options.adapter.snapshotSlide(input.slide_index, signal)
          const proposal = options.proposals.propose({
            operation: 'edit_slide_text',
            toolName: call.name,
            title: input.explanation || 'Edit slide text',
            preview: { shapeId: input.shape_id, before: before.text, after: input.text },
            impact: {
              host: 'powerpoint',
              targets: [`${snapshot.slideId}/${input.shape_id}`],
              count: 1,
            },
            fingerprint: snapshot.fingerprint,
            before: before.text,
            after: input.text,
            validate: async (confirmSignal) => {
              const currentSnapshot = await options.adapter.snapshotSlide(
                input.slide_index,
                confirmSignal,
              )
              if (currentSnapshot.fingerprint !== snapshot.fingerprint) return false
              const currentText = await options.adapter.readSlideText(
                input.slide_index,
                input.shape_id,
                confirmSignal,
              )
              return currentText.slideId === before.slideId && currentText.text === before.text
            },
            execute: (confirmSignal) =>
              options.adapter.editSlideText(
                input.slide_index,
                input.shape_id,
                input.text,
                confirmSignal,
              ),
            verify: async (confirmSignal) => {
              const result = await options.adapter.readSlideText(
                input.slide_index,
                input.shape_id,
                confirmSignal,
              )
              if (result.slideId !== before.slideId || result.text !== input.text)
                throw new Error('office_verify_failed')
              await options.adapter.verifySlides(confirmSignal)
            },
          })
          return {
            output: boundedJson(proposal),
            mutated: false,
            summary: 'Proposed PowerPoint text edit',
          }
        }
        if (call.name === 'duplicate_slide') {
          const input = slideInput(call.input)
          await options.adapter.verifySlides(signal)
          const snapshot = await options.adapter.snapshotSlide(input.slide_index, signal)
          let insertedSlideId: string | undefined
          const proposal = options.proposals.propose({
            operation: 'duplicate_slide',
            toolName: call.name,
            title: input.explanation || 'Duplicate slide',
            preview: { slideIndex: input.slide_index, slideId: snapshot.slideId },
            impact: { host: 'powerpoint', targets: [snapshot.slideId], count: 1 },
            fingerprint: snapshot.fingerprint,
            before: snapshot,
            validate: async (confirmSignal) =>
              (await options.adapter.snapshotSlide(input.slide_index, confirmSignal))
                .fingerprint === snapshot.fingerprint,
            execute: async (confirmSignal) => {
              insertedSlideId = (
                await options.adapter.duplicateSlide(input.slide_index, confirmSignal)
              ).slideId
            },
            verify: async (confirmSignal) => {
              if (!insertedSlideId) throw new Error('office_verify_failed')
              const inserted = await options.adapter.listSlideShapes(
                input.slide_index + 1,
                confirmSignal,
              )
              if (inserted.slideId !== insertedSlideId) throw new Error('office_verify_failed')
              await options.adapter.verifySlides(confirmSignal)
            },
          })
          return {
            output: boundedJson(proposal),
            mutated: false,
            summary: 'Proposed PowerPoint slide duplication',
          }
        }
        if (call.name === 'execute_office_js') {
          codeInput(call.input)
          return failure(call.name, 'office_api_unsupported')
        }
        if (call.name === 'edit_slide_master') {
          masterCodeInput(call.input)
          return failure(call.name, 'office_api_unsupported')
        }
        if (call.name === 'edit_slide_xml' || call.name === 'edit_slide_chart') {
          slideCodeInput(call.input)
          return failure(call.name, 'office_api_unsupported')
        }
        return failure(call.name, 'invalid_tool_input')
      } catch (error) {
        return failure(
          call.name,
          errorCode(error, ['edit_slide_text', 'duplicate_slide'].includes(call.name)),
        )
      }
    },
  }
}
