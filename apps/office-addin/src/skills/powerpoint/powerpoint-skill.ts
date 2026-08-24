import type { AgentSkill, ToolExecution } from '@wiswork/agent-core'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { exactObject, integerField, optionalField, stringField } from '../../agent/tool-schema.js'
import { parseDeclarativeProgram } from '../shared/declarative-program.js'
import { readUntilConverged } from '../shared/office-write-transaction.js'
import type { PowerPointAdapter } from './browser-powerpoint-adapter.js'
import {
  MAX_POWERPOINT_RESULT_BYTES,
  type PowerPointDeclarativeOperation,
} from './browser-powerpoint-adapter.js'
import {
  editPowerPointPackage,
  verifyImportedPowerPointPackage,
  verifyPowerPointPackageInputs,
  type PackageEditKind,
  type XmlReplacement,
} from './powerpoint-package.js'

const MAX_SLIDE_INDEX = 100_000
const MAX_CODE = 32 * 1024
const MAX_SCREENSHOT_BYTES = 4 * 1024 * 1024
const POWERPOINT_GEOMETRY_EPSILON = 0.01
const PROGRAM_TOOLS = new Set([
  'execute_office_js',
  'edit_slide_xml',
  'edit_slide_chart',
  'edit_slide_master',
])
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
const slideProperties = {
  slide_index: { type: 'integer', minimum: 0, maximum: MAX_SLIDE_INDEX },
  explanation: { type: 'string', maxLength: 50 },
} as const
const operationSlideIndex = { type: 'integer', minimum: 0, maximum: MAX_SLIDE_INDEX } as const
const operationShapeId = { type: 'string', minLength: 1, maxLength: 256 } as const
const geometryProperties = {
  left: { type: 'number' },
  top: { type: 'number' },
  width: { type: 'number', exclusiveMinimum: 0 },
  height: { type: 'number', exclusiveMinimum: 0 },
} as const
const exactOperation = (
  properties: Readonly<Record<string, unknown>>,
  required: readonly string[],
) => ({ type: 'object', properties, required, additionalProperties: false }) as const
const declarativeProgramSchema = {
  type: 'object',
  properties: {
    version: { type: 'integer', enum: [1] },
    operations: {
      type: 'array',
      minItems: 1,
      maxItems: 32,
      items: {
        anyOf: [
          exactOperation(
            {
              op: { type: 'string', enum: ['set_shape_text'] },
              slide_index: operationSlideIndex,
              shape_id: operationShapeId,
              text: { type: 'string', maxLength: 12_000 },
            },
            ['op', 'slide_index', 'shape_id', 'text'],
          ),
          exactOperation(
            {
              op: { type: 'string', enum: ['set_shape_geometry'] },
              slide_index: operationSlideIndex,
              shape_id: operationShapeId,
              ...geometryProperties,
            },
            ['op', 'slide_index', 'shape_id', 'left', 'top', 'width', 'height'],
          ),
          exactOperation(
            {
              op: { type: 'string', enum: ['add_text_box'] },
              slide_index: operationSlideIndex,
              name: { type: 'string', minLength: 1, maxLength: 256 },
              text: { type: 'string', maxLength: 12_000 },
              ...geometryProperties,
            },
            ['op', 'slide_index', 'name', 'text', 'left', 'top', 'width', 'height'],
          ),
          exactOperation(
            {
              op: { type: 'string', enum: ['delete_shape'] },
              slide_index: operationSlideIndex,
              shape_id: operationShapeId,
            },
            ['op', 'slide_index', 'shape_id'],
          ),
          exactOperation(
            {
              op: { type: 'string', enum: ['duplicate_slide'] },
              slide_index: operationSlideIndex,
            },
            ['op', 'slide_index'],
          ),
        ],
      },
    },
  },
  required: ['version', 'operations'],
  additionalProperties: false,
} as const
const xmlProgramSchema = {
  type: 'object',
  properties: {
    version: { type: 'integer', enum: [1] },
    operations: {
      type: 'array',
      minItems: 1,
      maxItems: 32,
      items: {
        type: 'object',
        properties: {
          op: { type: 'string', enum: ['replace_xml'] },
          path: { type: 'string', minLength: 1, maxLength: 256 },
          xml: { type: 'string', minLength: 1, maxLength: MAX_CODE },
        },
        required: ['op', 'path', 'xml'],
        additionalProperties: false,
      },
    },
  },
  required: ['version', 'operations'],
  additionalProperties: false,
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
      'Execute a confirmation-gated bounded declarative PowerPoint program. Pass program directly as an object with version 1 and an operations array; do not stringify it and do not send JavaScript. Use snake_case fields. Supported operations are set_shape_text (slide_index, shape_id, text), set_shape_geometry (slide_index, shape_id, left, top, width, height), add_text_box (slide_index, name, text, left, top, width, height), delete_shape (slide_index, shape_id), and duplicate_slide (slide_index; it must be the only operation).',
    inputSchema: {
      type: 'object',
      properties: {
        program: declarativeProgramSchema,
        explanation: { type: 'string', maxLength: 100 },
      },
      required: ['program'],
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
    description: 'Propose bounded allowlisted slide XML replacements in an exported slide package.',
    inputSchema: {
      type: 'object',
      properties: {
        ...slideProperties,
        program: xmlProgramSchema,
      },
      required: ['slide_index', 'program'],
      additionalProperties: false,
    },
  },
  {
    name: 'edit_slide_chart',
    description:
      'Propose bounded allowlisted chart XML replacements while preserving package relationships.',
    inputSchema: {
      type: 'object',
      properties: {
        ...slideProperties,
        program: xmlProgramSchema,
      },
      required: ['slide_index', 'program'],
      additionalProperties: false,
    },
  },
  {
    name: 'edit_slide_master',
    description: 'Propose bounded allowlisted master, layout, or theme XML replacements.',
    inputSchema: {
      type: 'object',
      properties: {
        program: xmlProgramSchema,
        explanation: { type: 'string', maxLength: 50 },
      },
      required: ['program'],
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

function failure(name: string, code: string, diagnosticError?: unknown): ToolExecution {
  return {
    output: code,
    isError: true,
    mutated: false,
    summary: name,
    ...(diagnosticError === undefined ? {} : { diagnosticError }),
  }
}
function invalidToolInput(location: 'program' | 'program.operations'): Error {
  return Object.assign(new Error('invalid_tool_input'), {
    code: 'InvalidToolInput',
    debugInfo: { errorLocation: location },
  })
}
function assertNotCancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new Error('cancelled')
}

async function verifyPowerPointReadback(
  verify: () => Promise<boolean>,
  signal?: AbortSignal,
): Promise<void> {
  const verified = await readUntilConverged({ read: verify, accept: Boolean, signal })
  if (!verified) throw new Error('office_verify_failed')
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

function exactRecord(value: unknown, keys: string[]): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw invalidToolInput('program')
  const record = value as Record<string, unknown>
  if (Object.keys(record).some((key) => !keys.includes(key))) throw new Error('invalid_tool_input')
  return record
}

function declarativeInput(
  value: unknown,
  options: { slide: boolean; explanationMax: number },
): { code: string; explanation?: string; slide_index?: number } {
  const keys = options.slide
    ? ['slide_index', 'program', 'code', 'explanation']
    : ['program', 'code', 'explanation']
  const input = exactRecord(value, keys)
  if ((input.program === undefined) === (input.code === undefined))
    throw invalidToolInput('program')
  if (
    input.explanation !== undefined &&
    (typeof input.explanation !== 'string' || input.explanation.length > options.explanationMax)
  )
    throw invalidToolInput('program')
  if (
    options.slide &&
    (!Number.isInteger(input.slide_index) ||
      (input.slide_index as number) < 0 ||
      (input.slide_index as number) > MAX_SLIDE_INDEX)
  )
    throw new Error('invalid_tool_input')
  let code: string
  if (input.code !== undefined) {
    if (typeof input.code !== 'string' || !input.code || input.code.length > MAX_CODE)
      throw invalidToolInput('program')
    code = input.code
  } else {
    try {
      code = JSON.stringify(input.program)
    } catch {
      throw invalidToolInput('program')
    }
    if (!code || new TextEncoder().encode(code).byteLength > MAX_CODE)
      throw invalidToolInput('program')
  }
  return {
    code,
    ...(typeof input.explanation === 'string' ? { explanation: input.explanation } : {}),
    ...(options.slide ? { slide_index: input.slide_index as number } : {}),
  }
}

function sameGeometry(actual: number, expected: number): boolean {
  return Math.abs(actual - expected) <= POWERPOINT_GEOMETRY_EPSILON
}

function parseXmlProgram(code: string): XmlReplacement[] {
  return parseDeclarativeProgram(code, (value) => {
    const operation = exactRecord(value, ['op', 'path', 'xml'])
    if (
      operation.op !== 'replace_xml' ||
      typeof operation.path !== 'string' ||
      !operation.path ||
      operation.path.length > 256 ||
      typeof operation.xml !== 'string' ||
      !operation.xml
    )
      throw new Error('invalid_tool_input')
    return { path: operation.path, xml: operation.xml }
  }).operations
}

function parsePowerPointOperation(value: unknown): PowerPointDeclarativeOperation {
  const root = exactRecord(value, [
    'op',
    'slide_index',
    'shape_id',
    'name',
    'text',
    'left',
    'top',
    'width',
    'height',
  ])
  const operation = root
  if (
    !Number.isInteger(operation.slide_index) ||
    (operation.slide_index as number) < 0 ||
    (operation.slide_index as number) > MAX_SLIDE_INDEX
  )
    throw new Error('invalid_tool_input')
  if (operation.op === 'duplicate_slide') {
    if (Object.keys(operation).some((key) => !['op', 'slide_index'].includes(key)))
      throw new Error('invalid_tool_input')
    return { op: 'duplicate_slide', slide_index: operation.slide_index as number }
  }
  const finiteGeometry = () => {
    for (const key of ['left', 'top', 'width', 'height'] as const)
      if (typeof operation[key] !== 'number' || !Number.isFinite(operation[key]))
        throw new Error('invalid_tool_input')
    if ((operation.width as number) <= 0 || (operation.height as number) <= 0)
      throw new Error('invalid_tool_input')
  }
  if (operation.op === 'set_shape_geometry') {
    if (
      Object.keys(operation).some(
        (key) => !['op', 'slide_index', 'shape_id', 'left', 'top', 'width', 'height'].includes(key),
      ) ||
      typeof operation.shape_id !== 'string' ||
      !operation.shape_id ||
      operation.shape_id.length > 256
    )
      throw new Error('invalid_tool_input')
    finiteGeometry()
    return {
      op: 'set_shape_geometry',
      slide_index: operation.slide_index as number,
      shape_id: operation.shape_id,
      left: operation.left as number,
      top: operation.top as number,
      width: operation.width as number,
      height: operation.height as number,
    }
  }
  if (operation.op === 'add_text_box') {
    if (
      Object.keys(operation).some(
        (key) =>
          !['op', 'slide_index', 'name', 'text', 'left', 'top', 'width', 'height'].includes(key),
      ) ||
      typeof operation.name !== 'string' ||
      !operation.name ||
      operation.name.length > 256 ||
      typeof operation.text !== 'string' ||
      operation.text.length > 12_000
    )
      throw new Error('invalid_tool_input')
    finiteGeometry()
    return {
      op: 'add_text_box',
      slide_index: operation.slide_index as number,
      name: operation.name,
      text: operation.text,
      left: operation.left as number,
      top: operation.top as number,
      width: operation.width as number,
      height: operation.height as number,
    }
  }
  if (operation.op === 'delete_shape') {
    if (
      Object.keys(operation).some((key) => !['op', 'slide_index', 'shape_id'].includes(key)) ||
      typeof operation.shape_id !== 'string' ||
      !operation.shape_id ||
      operation.shape_id.length > 256
    )
      throw new Error('invalid_tool_input')
    return {
      op: 'delete_shape',
      slide_index: operation.slide_index as number,
      shape_id: operation.shape_id,
    }
  }
  if (
    operation.op !== 'set_shape_text' ||
    Object.keys(operation).some(
      (key) => !['op', 'slide_index', 'shape_id', 'text'].includes(key),
    ) ||
    typeof operation.shape_id !== 'string' ||
    !operation.shape_id ||
    operation.shape_id.length > 256 ||
    typeof operation.text !== 'string' ||
    operation.text.length > 12_000
  )
    throw new Error('invalid_tool_input')
  return {
    op: 'set_shape_text',
    slide_index: operation.slide_index as number,
    shape_id: operation.shape_id,
    text: operation.text,
  }
}

export function createPowerPointSkill(options: {
  adapter: PowerPointAdapter
  proposals: StructuredProposalController
}): AgentSkill {
  async function proposePackageEdit(
    toolName: string,
    kind: PackageEditKind,
    slideIndex: number,
    replacements: XmlReplacement[],
    explanation: string | undefined,
    signal?: AbortSignal,
  ): Promise<ToolExecution> {
    const deck = await options.adapter.verifySlides(signal)
    const before = await options.adapter.exportSlidePackage(slideIndex, signal)
    const edited = await editPowerPointPackage(before.base64, kind, replacements, signal)
    let applied: Awaited<ReturnType<typeof editPowerPointPackage>> | undefined
    const proposal = options.proposals.propose({
      operation: toolName,
      toolName,
      title: explanation || `Edit PowerPoint ${kind} XML`,
      preview: {
        kind,
        slideIndex,
        changedPaths: edited.changedPaths,
        beforeHashes: edited.beforeHashes,
        afterHashes: edited.afterHashes,
      },
      impact: {
        host: 'powerpoint',
        targets:
          kind === 'master'
            ? deck.slides.map((slide) => `slide:${slide.slideId}`)
            : edited.changedPaths,
        count: kind === 'master' ? deck.slides.length : edited.changedPaths.length,
      },
      fingerprint: before.fingerprint,
      before: { slideId: before.slideId, hashes: edited.beforeHashes },
      after: { hashes: edited.afterHashes },
      code: JSON.stringify({
        version: 1,
        operations: replacements.map((item) => ({ op: 'replace_xml', ...item })),
      }),
      validate: async (confirmSignal) => {
        const current = await options.adapter.exportSlidePackage(slideIndex, confirmSignal)
        return verifyPowerPointPackageInputs(current.base64, edited.beforeHashes, confirmSignal)
      },
      execute: async (confirmSignal) => {
        const current = await options.adapter.exportSlidePackage(slideIndex, confirmSignal)
        if (
          !(await verifyPowerPointPackageInputs(current.base64, edited.beforeHashes, confirmSignal))
        )
          throw new Error('proposal_stale')
        applied = await editPowerPointPackage(current.base64, kind, replacements, confirmSignal)
        await options.adapter.replaceSlidePackage(
          slideIndex,
          applied.base64,
          kind === 'master',
          applied,
          confirmSignal,
        )
      },
      verify: async (confirmSignal) => {
        if (!applied) throw new Error('office_verify_failed')
        const current = await options.adapter.exportSlidePackage(slideIndex, confirmSignal)
        if (!(await verifyImportedPowerPointPackage(current.base64, applied, confirmSignal)))
          throw new Error('office_verify_failed')
        await options.adapter.verifySlides(confirmSignal)
      },
    })
    return {
      output: boundedJson(proposal),
      mutated: false,
      summary: `Proposed PowerPoint ${kind} XML edit`,
    }
  }

  return {
    id: 'office-powerpoint',
    systemPrompt:
      'PowerPoint reads are bounded. Every write creates an explicit proposal and is semantically verified after confirmation. execute_office_js accepts only a versioned declarative JSON program; JavaScript and ambient browser authority are rejected. XML tools accept only allowlisted bounded package parts.',
    tools: [...tools],
    async executeTool(call, signal) {
      if (call.inputError || call.truncated)
        return failure(
          call.name,
          'invalid_tool_input',
          PROGRAM_TOOLS.has(call.name) ? invalidToolInput('program') : undefined,
        )
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
              visualAvailableToModel: true,
            }),
            modelContent: [{ type: 'image', image: { mime: result.mime, base64: result.base64 } }],
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
          const stableTextFingerprint = fingerprint(
            JSON.stringify([before.slideId, before.shapeId, before.text, before.paragraphs]),
          )
          const proposal = options.proposals.propose({
            operation: 'edit_slide_text',
            toolName: call.name,
            title: input.explanation || 'Edit slide text',
            preview: { shapeId: input.shape_id, before: before.text, after: input.text },
            impact: {
              host: 'powerpoint',
              targets: [`${before.slideId}/${input.shape_id}`],
              count: 1,
            },
            fingerprint: stableTextFingerprint,
            before: before.text,
            after: input.text,
            validate: async (confirmSignal) => {
              const currentText = await options.adapter.readSlideText(
                input.slide_index,
                input.shape_id,
                confirmSignal,
              )
              return (
                fingerprint(
                  JSON.stringify([
                    currentText.slideId,
                    currentText.shapeId,
                    currentText.text,
                    currentText.paragraphs,
                  ]),
                ) === stableTextFingerprint
              )
            },
            execute: (confirmSignal) =>
              options.adapter.editSlideText(
                input.slide_index,
                input.shape_id,
                input.text,
                confirmSignal,
              ),
            verify: async (confirmSignal) => {
              await verifyPowerPointReadback(async () => {
                const result = await options.adapter.readSlideText(
                  input.slide_index,
                  input.shape_id,
                  confirmSignal,
                )
                return result.slideId === before.slideId && result.text === input.text
              }, confirmSignal)
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
              await verifyPowerPointReadback(async () => {
                const inserted = await options.adapter.listSlideShapes(
                  input.slide_index + 1,
                  confirmSignal,
                )
                return inserted.slideId === insertedSlideId
              }, confirmSignal)
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
          const input = declarativeInput(call.input, { slide: false, explanationMax: 100 })
          let program
          try {
            program = parseDeclarativeProgram(input.code, parsePowerPointOperation)
          } catch (error) {
            if (error instanceof Error && error.message === 'invalid_tool_input')
              throw invalidToolInput('program.operations')
            throw error
          }
          if (
            program.operations.some((operation) => operation.op === 'duplicate_slide') &&
            program.operations.length !== 1
          )
            throw new Error('invalid_tool_input')
          const shapeTargets = new Map<string, PowerPointDeclarativeOperation[]>()
          for (const operation of program.operations) {
            if (!('shape_id' in operation)) continue
            const key = `${operation.slide_index}/${operation.shape_id}`
            const related = shapeTargets.get(key) ?? []
            related.push(operation)
            shapeTargets.set(key, related)
          }
          if (
            [...shapeTargets.values()].some(
              (related) =>
                related.some((operation) => operation.op === 'delete_shape') && related.length > 1,
            )
          )
            throw new Error('invalid_tool_input')
          await options.adapter.verifySlides(signal)
          const slideIndexes = [
            ...new Set(program.operations.map((operation) => operation.slide_index)),
          ]
          if (slideIndexes.length > 8) throw new Error('invalid_tool_input')
          const snapshots = await Promise.all(
            slideIndexes.map((index) => options.adapter.snapshotSlide(index, signal)),
          )
          const beforeTexts = await Promise.all(
            program.operations.flatMap((operation) =>
              operation.op === 'set_shape_text'
                ? [options.adapter.readSlideText(operation.slide_index, operation.shape_id, signal)]
                : [],
            ),
          )
          const combined = snapshots.map((item) => item.fingerprint).join('|')
          let declarativeResult: { createdShapeIds: string[]; insertedSlideId?: string } | undefined
          const proposal = options.proposals.propose({
            operation: call.name,
            toolName: call.name,
            title: input.explanation || 'Execute declarative PowerPoint operations',
            preview: { version: 1, operations: program.operations },
            impact: {
              host: 'powerpoint',
              targets: program.operations.map((operation) =>
                operation.op === 'set_shape_text'
                  ? `${operation.slide_index}/${operation.shape_id}`
                  : `${operation.slide_index}`,
              ),
              count: program.operations.length,
            },
            fingerprint: fingerprint(combined),
            code: input.code,
            before: {
              slides: snapshots.map(({ slideId, fingerprint: value }) => ({
                slideId,
                fingerprint: value,
              })),
              texts: beforeTexts.map((item) => ({
                slideId: item.slideId,
                shapeId: item.shapeId,
                text: item.text,
              })),
            },
            after: { operations: program.operations },
            validate: async (confirmSignal) => {
              const current = await Promise.all(
                slideIndexes.map((index) => options.adapter.snapshotSlide(index, confirmSignal)),
              )
              return current.every(
                (item, index) => item.fingerprint === snapshots[index].fingerprint,
              )
            },
            execute: async (confirmSignal) => {
              declarativeResult = await options.adapter.executeDeclarative(
                program.operations,
                confirmSignal,
              )
            },
            verify: async (confirmSignal) => {
              let createdShapeIndex = 0
              for (const [operationIndex, operation] of program.operations.entries()) {
                const superseded = program.operations.slice(operationIndex + 1).some((later) => {
                  if (
                    !('shape_id' in operation) ||
                    !('shape_id' in later) ||
                    later.slide_index !== operation.slide_index ||
                    later.shape_id !== operation.shape_id
                  )
                    return false
                  return (
                    later.op === 'delete_shape' ||
                    (operation.op === 'set_shape_text' && later.op === 'set_shape_text') ||
                    (operation.op === 'set_shape_geometry' && later.op === 'set_shape_geometry')
                  )
                })
                if (superseded) continue
                if (operation.op === 'set_shape_text') {
                  await verifyPowerPointReadback(async () => {
                    const current = await options.adapter.readSlideText(
                      operation.slide_index,
                      operation.shape_id,
                      confirmSignal,
                    )
                    return current.text === operation.text
                  }, confirmSignal)
                } else if (operation.op !== 'duplicate_slide') {
                  if (operation.op === 'add_text_box') {
                    const createdShapeId = declarativeResult?.createdShapeIds[createdShapeIndex++]
                    await verifyPowerPointReadback(async () => {
                      const current = await options.adapter.listSlideShapes(
                        operation.slide_index,
                        confirmSignal,
                      )
                      const shape = current.shapes.find((item) => item.id === createdShapeId)
                      if (
                        shape &&
                        sameGeometry(shape.left, operation.left) &&
                        sameGeometry(shape.top, operation.top) &&
                        sameGeometry(shape.width, operation.width) &&
                        sameGeometry(shape.height, operation.height)
                      ) {
                        const text = await options.adapter.readSlideText(
                          operation.slide_index,
                          shape.id,
                          confirmSignal,
                        )
                        if (text.text === operation.text) return true
                      }
                      return false
                    }, confirmSignal)
                    continue
                  }
                  await verifyPowerPointReadback(async () => {
                    const current = await options.adapter.listSlideShapes(
                      operation.slide_index,
                      confirmSignal,
                    )
                    const shape = current.shapes.find((item) => item.id === operation.shape_id)
                    if (operation.op === 'delete_shape') return !shape
                    return Boolean(
                      shape &&
                      sameGeometry(shape.left, operation.left) &&
                      sameGeometry(shape.top, operation.top) &&
                      sameGeometry(shape.width, operation.width) &&
                      sameGeometry(shape.height, operation.height),
                    )
                  }, confirmSignal)
                }
              }
              if (program.operations[0]?.op === 'duplicate_slide') {
                const operation = program.operations[0]
                const insertedSlideId = declarativeResult?.insertedSlideId
                if (!insertedSlideId) throw new Error('office_verify_failed')
                await verifyPowerPointReadback(async () => {
                  const inserted = await options.adapter.listSlideShapes(
                    operation.slide_index + 1,
                    confirmSignal,
                  )
                  return inserted.slideId === insertedSlideId
                }, confirmSignal)
              }
              await options.adapter.verifySlides(confirmSignal)
            },
          })
          return {
            output: boundedJson(proposal),
            mutated: false,
            summary: 'Proposed declarative PowerPoint execution',
          }
        }
        if (call.name === 'edit_slide_master') {
          const input = declarativeInput(call.input, { slide: false, explanationMax: 50 })
          return await proposePackageEdit(
            call.name,
            'master',
            0,
            (() => {
              try {
                return parseXmlProgram(input.code)
              } catch (error) {
                if (error instanceof Error && error.message === 'invalid_tool_input')
                  throw invalidToolInput('program.operations')
                throw error
              }
            })(),
            input.explanation,
            signal,
          )
        }
        if (call.name === 'edit_slide_xml' || call.name === 'edit_slide_chart') {
          const input = declarativeInput(call.input, { slide: true, explanationMax: 50 })
          return await proposePackageEdit(
            call.name,
            call.name === 'edit_slide_chart' ? 'chart' : 'slide',
            input.slide_index!,
            (() => {
              try {
                return parseXmlProgram(input.code)
              } catch (error) {
                if (error instanceof Error && error.message === 'invalid_tool_input')
                  throw invalidToolInput('program.operations')
                throw error
              }
            })(),
            input.explanation,
            signal,
          )
        }
        return failure(call.name, 'invalid_tool_input')
      } catch (error) {
        const code = errorCode(error, ['edit_slide_text', 'duplicate_slide'].includes(call.name))
        return failure(call.name, code, code === 'invalid_tool_input' ? error : undefined)
      }
    },
  }
}
