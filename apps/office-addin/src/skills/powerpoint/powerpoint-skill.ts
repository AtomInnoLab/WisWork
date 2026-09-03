import type { AgentSkill, ToolExecution } from '@wiswork/agent-core'
import type { PresentationVerificationFlags } from '@wiswork/presentation-verification'
import type { PresentationTelemetryEvent } from '@wiswork/presentation-verification'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { exactObject, integerField, optionalField, stringField } from '../../agent/tool-schema.js'
import { parseDeclarativeProgram } from '../shared/declarative-program.js'
import { readUntilConverged } from '../shared/office-write-transaction.js'
import { readBoundedImage } from '../shared/import-media.js'
import type { InMemoryVfs } from '../shared/vfs.js'
import type { PowerPointAdapter } from './browser-powerpoint-adapter.js'
import {
  createOfficePowerPointVerification,
  canonicalPowerPointVerificationBinding,
  powerPointProposalFingerprint,
  type OfficePowerPointVerificationAuthority,
  type OfficePowerPointVisualReviewer,
} from './powerpoint-verification.js'
import {
  MAX_POWERPOINT_RESULT_BYTES,
  type PowerPointDeclarativeOperation,
  type PowerPointMasterOperation,
  type PowerPointMasterState,
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

function arrayField<T>(
  parser: (value: unknown) => T,
  options: { minItems?: number; maxItems: number },
): (value: unknown) => T[] {
  return (value) => {
    if (
      !Array.isArray(value) ||
      value.length < (options.minItems ?? 0) ||
      value.length > options.maxItems
    )
      throw new Error('invalid_tool_input')
    return value.map((item) => parser(item))
  }
}
const MASTER_PATTERN_TYPES = [
  'Percent5',
  'Percent10',
  'Percent20',
  'Percent25',
  'Percent30',
  'Percent40',
  'Percent50',
  'Percent60',
  'Percent70',
  'Percent75',
  'Percent80',
  'Percent90',
  'Horizontal',
  'Vertical',
  'LightHorizontal',
  'LightVertical',
  'DarkHorizontal',
  'DarkVertical',
  'NarrowHorizontal',
  'NarrowVertical',
  'DashedHorizontal',
  'DashedVertical',
  'Cross',
  'DownwardDiagonal',
  'UpwardDiagonal',
  'LightDownwardDiagonal',
  'LightUpwardDiagonal',
  'DarkDownwardDiagonal',
  'DarkUpwardDiagonal',
  'WideDownwardDiagonal',
  'WideUpwardDiagonal',
  'DashedDownwardDiagonal',
  'DashedUpwardDiagonal',
  'DiagonalCross',
  'SmallCheckerBoard',
  'LargeCheckerBoard',
  'SmallGrid',
  'LargeGrid',
  'DottedGrid',
  'SmallConfetti',
  'LargeConfetti',
  'HorizontalBrick',
  'DiagonalBrick',
  'SolidDiamond',
  'OutlinedDiamond',
  'DottedDiamond',
  'Plaid',
  'Sphere',
  'Weave',
  'Divot',
  'Shingle',
  'Wave',
  'Trellis',
  'ZigZag',
] as const
const PROGRAM_TOOLS = new Set([
  'execute_office_js',
  'edit_slide_xml',
  'edit_slide_chart',
  'edit_slide_master',
  'edit_slide_master_xml',
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
const planDeckInput = exactObject({
  core_hook: stringField({ minLength: 1, maxLength: 500 }),
  style: stringField({ minLength: 1, maxLength: 1_000 }),
  pages: arrayField(
    exactObject({
      title: stringField({ minLength: 1, maxLength: 300 }),
      type: optionalField(stringField({ maxLength: 50 })),
      brief: stringField({ minLength: 1, maxLength: 2_000 }),
      layout: stringField({ minLength: 1, maxLength: 100 }),
      image_queries: optionalField(
        arrayField(stringField({ minLength: 1, maxLength: 200 }), { maxItems: 4 }),
      ),
    }),
    { minItems: 1, maxItems: 20 },
  ),
})
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
              op: { type: 'string', enum: ['set_shape_text_style'] },
              slide_index: operationSlideIndex,
              shape_id: operationShapeId,
              color: { type: 'string', pattern: '^#[0-9A-Fa-f]{6}$' },
              fontFamily: { type: 'string', minLength: 1, maxLength: 128 },
              fontSize: { type: 'number', minimum: 1, maximum: 400 },
              bold: { type: 'boolean' },
              italic: { type: 'boolean' },
            },
            ['op', 'slide_index', 'shape_id'],
          ),
          exactOperation(
            {
              op: { type: 'string', enum: ['set_shape_geometry'] },
              slide_index: operationSlideIndex,
              shape_id: operationShapeId,
              reference_slide_index: operationSlideIndex,
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
const masterProgramSchema = {
  type: 'object',
  properties: {
    version: { type: 'integer', enum: [2] },
    operations: {
      type: 'array',
      minItems: 1,
      maxItems: 32,
      items: {
        anyOf: [
          exactOperation(
            {
              op: { type: 'string', enum: ['set_master_background'] },
              master_id: { type: 'string', minLength: 1, maxLength: 256 },
              fill: {
                anyOf: [
                  exactOperation(
                    {
                      type: { type: 'string', enum: ['solid'] },
                      color: { type: 'string', pattern: '^#[0-9A-Fa-f]{6}$' },
                      transparency: { type: 'number', minimum: 0, maximum: 1 },
                    },
                    ['type', 'color', 'transparency'],
                  ),
                  exactOperation(
                    {
                      type: { type: 'string', enum: ['gradient'] },
                      gradient_type: {
                        type: 'string',
                        enum: ['Linear', 'Radial', 'Rectangular', 'Path', 'ShadeFromTitle'],
                      },
                    },
                    ['type', 'gradient_type'],
                  ),
                  exactOperation(
                    {
                      type: { type: 'string', enum: ['pattern'] },
                      pattern: { type: 'string', enum: MASTER_PATTERN_TYPES },
                      foreground_color: { type: 'string', pattern: '^#[0-9A-Fa-f]{6}$' },
                      background_color: { type: 'string', pattern: '^#[0-9A-Fa-f]{6}$' },
                    },
                    ['type', 'pattern', 'foreground_color', 'background_color'],
                  ),
                  exactOperation(
                    {
                      type: { type: 'string', enum: ['picture_or_texture'] },
                      path: { type: 'string', minLength: 1, maxLength: 1024 },
                      transparency: { type: 'number', minimum: 0, maximum: 1 },
                    },
                    ['type', 'path', 'transparency'],
                  ),
                ],
              },
            },
            ['op', 'master_id', 'fill'],
          ),
          exactOperation(
            {
              op: { type: 'string', enum: ['set_master_theme_color'] },
              master_id: { type: 'string', minLength: 1, maxLength: 256 },
              theme_color: {
                type: 'string',
                enum: [
                  'Accent1',
                  'Accent2',
                  'Accent3',
                  'Accent4',
                  'Accent5',
                  'Accent6',
                  'Dark1',
                  'Dark2',
                  'Light1',
                  'Light2',
                  'Hyperlink',
                  'FollowedHyperlink',
                ],
              },
              color: { type: 'string', pattern: '^#[0-9A-Fa-f]{6}$' },
            },
            ['op', 'master_id', 'theme_color', 'color'],
          ),
          exactOperation(
            {
              op: { type: 'string', enum: ['set_layout_background_following'] },
              master_id: { type: 'string', minLength: 1, maxLength: 256 },
              layout_id: { type: 'string', minLength: 1, maxLength: 256 },
              follow_master: { type: 'boolean' },
              show_master_graphics: { type: 'boolean' },
            },
            ['op', 'master_id', 'layout_id', 'follow_master', 'show_master_graphics'],
          ),
        ],
      },
    },
  },
  required: ['version', 'operations'],
  additionalProperties: false,
} as const
const tools = [
  {
    name: 'inspect_slide_masters',
    description: 'Inspect bounded native slide masters, layouts, backgrounds, and theme colors.',
    inputSchema: {
      type: 'object',
      properties: { explanation: { type: 'string', maxLength: 50 } },
      required: [],
      additionalProperties: false,
    },
  },
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
    name: 'plan_deck',
    description:
      'Record the complete narrative and visual plan before creating or substantially rebuilding a presentation. This tool never edits PowerPoint.',
    inputSchema: {
      type: 'object',
      properties: {
        core_hook: { type: 'string', minLength: 1, maxLength: 500 },
        style: { type: 'string', minLength: 1, maxLength: 1_000 },
        pages: {
          type: 'array',
          minItems: 1,
          maxItems: 20,
          items: {
            type: 'object',
            properties: {
              title: { type: 'string', minLength: 1, maxLength: 300 },
              type: { type: 'string', maxLength: 50 },
              brief: { type: 'string', minLength: 1, maxLength: 2_000 },
              layout: { type: 'string', minLength: 1, maxLength: 100 },
              image_queries: {
                type: 'array',
                maxItems: 4,
                items: { type: 'string', minLength: 1, maxLength: 200 },
              },
            },
            required: ['title', 'brief', 'layout'],
            additionalProperties: false,
          },
        },
      },
      required: ['core_hook', 'style', 'pages'],
      additionalProperties: false,
    },
  },
  {
    name: 'execute_office_js',
    description:
      'Execute a confirmation-gated bounded declarative PowerPoint program. Pass program directly as an object with version 1 and an operations array; do not stringify it and do not send JavaScript. Use snake_case fields except the bounded text-style properties. Supported operations are set_shape_text, set_shape_text_style (color/fontFamily/fontSize/bold/italic), set_shape_geometry, add_text_box, delete_shape, and duplicate_slide.',
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
    description:
      'Propose native PowerPointApi 1.10 master background, theme color, and layout inheritance edits using a version 2 declarative program.',
    inputSchema: {
      type: 'object',
      properties: {
        program: masterProgramSchema,
        explanation: { type: 'string', maxLength: 100 },
      },
      required: ['program'],
      additionalProperties: false,
    },
  },
  {
    name: 'edit_slide_master_xml',
    description:
      'Propose bounded allowlisted master, layout, or theme XML replacements on hosts with reliable package import.',
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
  return powerPointProposalFingerprint(value)
}

function parseMasterProgram(value: unknown): PowerPointMasterOperation[] {
  const program = exactRecord(value, ['version', 'operations'])
  if (
    program.version !== 2 ||
    !Array.isArray(program.operations) ||
    program.operations.length < 1 ||
    program.operations.length > 32
  )
    throw invalidToolInput('program.operations')
  return program.operations.map((raw) => {
    const operation = exactRecord(raw, [
      'op',
      'master_id',
      'layout_id',
      'fill',
      'theme_color',
      'color',
      'follow_master',
      'show_master_graphics',
    ])
    if (
      typeof operation.master_id !== 'string' ||
      !operation.master_id ||
      operation.master_id.length > 256
    )
      throw invalidToolInput('program.operations')
    if (operation.op === 'set_master_background') {
      const fill = exactRecord(operation.fill, [
        'type',
        'color',
        'transparency',
        'gradient_type',
        'pattern',
        'foreground_color',
        'background_color',
        'image_base64',
      ])
      if (
        fill.type === 'solid' &&
        typeof fill.color === 'string' &&
        /^#[0-9A-Fa-f]{6}$/.test(fill.color) &&
        typeof fill.transparency === 'number' &&
        fill.transparency >= 0 &&
        fill.transparency <= 1
      )
        return {
          op: operation.op,
          master_id: operation.master_id,
          fill: { type: 'solid', color: fill.color.toUpperCase(), transparency: fill.transparency },
        }
      if (
        fill.type === 'picture_or_texture' &&
        typeof fill.image_base64 === 'string' &&
        fill.image_base64.length > 0 &&
        typeof fill.transparency === 'number' &&
        fill.transparency >= 0 &&
        fill.transparency <= 1
      )
        return {
          op: operation.op,
          master_id: operation.master_id,
          fill: {
            type: 'picture_or_texture',
            image_base64: fill.image_base64,
            transparency: fill.transparency,
          },
        }
      if (
        fill.type === 'gradient' &&
        typeof fill.gradient_type === 'string' &&
        ['Linear', 'Radial', 'Rectangular', 'Path', 'ShadeFromTitle'].includes(fill.gradient_type)
      )
        return {
          op: operation.op,
          master_id: operation.master_id,
          fill: { type: 'gradient', gradient_type: fill.gradient_type },
        }
      if (
        fill.type === 'pattern' &&
        typeof fill.pattern === 'string' &&
        (MASTER_PATTERN_TYPES as readonly string[]).includes(fill.pattern) &&
        typeof fill.foreground_color === 'string' &&
        /^#[0-9A-Fa-f]{6}$/.test(fill.foreground_color) &&
        typeof fill.background_color === 'string' &&
        /^#[0-9A-Fa-f]{6}$/.test(fill.background_color)
      )
        return {
          op: operation.op,
          master_id: operation.master_id,
          fill: {
            type: 'pattern',
            pattern: fill.pattern,
            foreground_color: fill.foreground_color.toUpperCase(),
            background_color: fill.background_color.toUpperCase(),
          },
        }
      throw invalidToolInput('program.operations')
    }
    if (operation.op === 'set_master_theme_color') {
      const slots = new Set([
        'Accent1',
        'Accent2',
        'Accent3',
        'Accent4',
        'Accent5',
        'Accent6',
        'Dark1',
        'Dark2',
        'Light1',
        'Light2',
        'Hyperlink',
        'FollowedHyperlink',
      ])
      if (
        typeof operation.theme_color !== 'string' ||
        !slots.has(operation.theme_color) ||
        typeof operation.color !== 'string' ||
        !/^#[0-9A-Fa-f]{6}$/.test(operation.color)
      )
        throw invalidToolInput('program.operations')
      return {
        op: operation.op,
        master_id: operation.master_id,
        theme_color: operation.theme_color,
        color: operation.color.toUpperCase(),
      }
    }
    if (
      operation.op === 'set_layout_background_following' &&
      typeof operation.layout_id === 'string' &&
      operation.layout_id &&
      typeof operation.follow_master === 'boolean' &&
      typeof operation.show_master_graphics === 'boolean'
    )
      return {
        op: operation.op,
        master_id: operation.master_id,
        layout_id: operation.layout_id,
        follow_master: operation.follow_master,
        show_master_graphics: operation.show_master_graphics,
      }
    throw invalidToolInput('program.operations')
  })
}

async function prepareMasterProgram(value: unknown, vfs?: InMemoryVfs): Promise<unknown> {
  const copy = structuredClone(value) as {
    operations?: Array<{ fill?: { type?: unknown; path?: unknown; transparency?: unknown } }>
  }
  if (!Array.isArray(copy?.operations)) return copy
  for (const operation of copy.operations) {
    const fill = operation?.fill
    if (fill?.type !== 'picture_or_texture') continue
    if (!vfs || typeof fill.path !== 'string') throw invalidToolInput('program.operations')
    const image = await readBoundedImage(vfs, fill.path)
    operation.fill = {
      type: 'picture_or_texture',
      transparency: fill.transparency,
      image_base64: image.base64,
    } as typeof fill
  }
  return copy
}

function projectedMasterState(
  before: PowerPointMasterState,
  operations: PowerPointMasterOperation[],
): PowerPointMasterState {
  const value = structuredClone(before)
  for (const operation of operations) {
    const master = value.masters.find((item) => item.id === operation.master_id)
    if (!master) throw new Error('invalid_tool_input')
    if (operation.op === 'set_master_background') {
      if (operation.fill.type === 'solid')
        master.background = {
          type: 'Solid',
          color: operation.fill.color,
          transparency: operation.fill.transparency,
        }
      else if (operation.fill.type === 'gradient')
        master.background = {
          type: 'Gradient',
          gradientType: operation.fill.gradient_type,
        } as PowerPointMasterState['masters'][number]['background']
      else if (operation.fill.type === 'pattern')
        master.background = {
          type: 'Pattern',
          pattern: operation.fill.pattern,
          foregroundColor: operation.fill.foreground_color,
          backgroundColor: operation.fill.background_color,
        } as PowerPointMasterState['masters'][number]['background']
      else
        master.background = {
          type: 'PictureOrTexture',
          pictureTransparency: operation.fill.transparency,
        }
    } else if (operation.op === 'set_master_theme_color')
      master.themeColors[operation.theme_color] = operation.color
    else {
      const layout = master.layouts.find((item) => item.id === operation.layout_id)
      if (!layout) throw new Error('invalid_tool_input')
      layout.isMasterBackgroundFollowed = operation.follow_master
      layout.areBackgroundGraphicsHidden = !operation.show_master_graphics
    }
  }
  return value
}

function masterOperationKey(operation: PowerPointMasterOperation): string {
  if (operation.op === 'set_master_background') return `${operation.master_id}:background`
  if (operation.op === 'set_master_theme_color')
    return `${operation.master_id}:theme:${operation.theme_color}`
  return `${operation.master_id}:layout:${operation.layout_id}:background-following`
}

function normalizedColor(value: unknown): unknown {
  return typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value) ? value.toUpperCase() : value
}

function masterOperationValue(
  state: PowerPointMasterState,
  operation: PowerPointMasterOperation,
): unknown {
  const master = state.masters.find((item) => item.id === operation.master_id)
  if (!master) return undefined
  if (operation.op === 'set_master_background') {
    const background = master.background
    return {
      type: background.type.toLowerCase(),
      ...(background.color === undefined ? {} : { color: normalizedColor(background.color) }),
      ...(background.transparency === undefined ? {} : { transparency: background.transparency }),
      ...(background.gradientType === undefined
        ? {}
        : { gradientType: background.gradientType.toLowerCase() }),
      ...(background.pattern === undefined ? {} : { pattern: background.pattern.toLowerCase() }),
      ...(background.foregroundColor === undefined
        ? {}
        : { foregroundColor: normalizedColor(background.foregroundColor) }),
      ...(background.backgroundColor === undefined
        ? {}
        : { backgroundColor: normalizedColor(background.backgroundColor) }),
      ...(background.pictureTransparency === undefined
        ? {}
        : { pictureTransparency: background.pictureTransparency }),
    }
  }
  if (operation.op === 'set_master_theme_color')
    return normalizedColor(master.themeColors[operation.theme_color])
  const layout = master.layouts.find((item) => item.id === operation.layout_id)
  return layout
    ? {
        follow_master: layout.isMasterBackgroundFollowed,
        show_master_graphics: !layout.areBackgroundGraphicsHidden,
      }
    : undefined
}

function affectedMasterFingerprint(
  state: PowerPointMasterState,
  operations: PowerPointMasterOperation[],
): string {
  return fingerprint(
    JSON.stringify(
      operations.map((operation) => [
        masterOperationKey(operation),
        masterOperationValue(state, operation),
      ]),
    ),
  )
}

function sameMasterOperationValue(
  actual: PowerPointMasterState,
  expected: PowerPointMasterState,
  operation: PowerPointMasterOperation,
): boolean {
  return (
    JSON.stringify(masterOperationValue(actual, operation)) ===
    JSON.stringify(masterOperationValue(expected, operation))
  )
}

function inverseMasterOperation(
  before: PowerPointMasterState,
  operation: PowerPointMasterOperation,
): PowerPointMasterOperation {
  const master = before.masters.find((item) => item.id === operation.master_id)
  if (!master) throw new Error('invalid_tool_input')
  if (operation.op === 'set_master_background') {
    const type = master.background.type.toLowerCase()
    if (
      type === 'solid' &&
      typeof master.background.color === 'string' &&
      typeof master.background.transparency === 'number'
    )
      return {
        op: 'set_master_background',
        master_id: operation.master_id,
        fill: {
          type: 'solid',
          color: master.background.color,
          transparency: master.background.transparency,
        },
      }
    if (type === 'gradient' && typeof master.background.gradientType === 'string')
      return {
        op: 'set_master_background',
        master_id: operation.master_id,
        fill: { type: 'gradient', gradient_type: master.background.gradientType },
      }
    if (
      type === 'pattern' &&
      typeof master.background.pattern === 'string' &&
      typeof master.background.foregroundColor === 'string' &&
      typeof master.background.backgroundColor === 'string'
    )
      return {
        op: 'set_master_background',
        master_id: operation.master_id,
        fill: {
          type: 'pattern',
          pattern: master.background.pattern,
          foreground_color: master.background.foregroundColor,
          background_color: master.background.backgroundColor,
        },
      }
    throw new Error('office_api_unsupported')
  }
  if (operation.op === 'set_master_theme_color') {
    const color = master.themeColors[operation.theme_color]
    if (!color) throw new Error('office_api_unsupported')
    return {
      op: operation.op,
      master_id: operation.master_id,
      theme_color: operation.theme_color,
      color,
    }
  }
  const layout = master.layouts.find((item) => item.id === operation.layout_id)
  if (!layout) throw new Error('invalid_tool_input')
  return {
    op: operation.op,
    master_id: operation.master_id,
    layout_id: operation.layout_id,
    follow_master: layout.isMasterBackgroundFollowed,
    show_master_graphics: !layout.areBackgroundGraphicsHidden,
  }
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
    'color',
    'fontFamily',
    'fontSize',
    'bold',
    'italic',
    'reference_slide_index',
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
        (key) =>
          ![
            'op',
            'slide_index',
            'shape_id',
            'left',
            'top',
            'width',
            'height',
            'reference_slide_index',
          ].includes(key),
      ) ||
      typeof operation.shape_id !== 'string' ||
      !operation.shape_id ||
      operation.shape_id.length > 256
    )
      throw new Error('invalid_tool_input')
    if (
      operation.reference_slide_index !== undefined &&
      (!Number.isInteger(operation.reference_slide_index) ||
        (operation.reference_slide_index as number) < 0 ||
        (operation.reference_slide_index as number) > MAX_SLIDE_INDEX)
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
      ...(Number.isInteger(operation.reference_slide_index)
        ? { reference_slide_index: operation.reference_slide_index as number }
        : {}),
    }
  }
  if (operation.op === 'set_shape_text_style') {
    const allowed = [
      'op',
      'slide_index',
      'shape_id',
      'color',
      'fontFamily',
      'fontSize',
      'bold',
      'italic',
    ]
    if (
      Object.keys(operation).some((key) => !allowed.includes(key)) ||
      typeof operation.shape_id !== 'string' ||
      !operation.shape_id ||
      (Object.hasOwn(operation, 'color') &&
        (typeof operation.color !== 'string' || !/^#[0-9A-Fa-f]{6}$/.test(operation.color))) ||
      (Object.hasOwn(operation, 'fontFamily') &&
        (typeof operation.fontFamily !== 'string' ||
          !operation.fontFamily ||
          operation.fontFamily.length > 128)) ||
      (Object.hasOwn(operation, 'fontSize') &&
        (typeof operation.fontSize !== 'number' ||
          !Number.isFinite(operation.fontSize) ||
          operation.fontSize < 1 ||
          operation.fontSize > 400)) ||
      (Object.hasOwn(operation, 'bold') && typeof operation.bold !== 'boolean') ||
      (Object.hasOwn(operation, 'italic') && typeof operation.italic !== 'boolean')
    )
      throw new Error('invalid_tool_input')
    const style = {
      ...(typeof operation.color === 'string' && /^#[0-9A-Fa-f]{6}$/.test(operation.color)
        ? { color: operation.color.toUpperCase() }
        : {}),
      ...(typeof operation.fontFamily === 'string' && operation.fontFamily.length <= 128
        ? { fontFamily: operation.fontFamily }
        : {}),
      ...(typeof operation.fontSize === 'number' &&
      operation.fontSize >= 1 &&
      operation.fontSize <= 400
        ? { fontSize: operation.fontSize }
        : {}),
      ...(typeof operation.bold === 'boolean' ? { bold: operation.bold } : {}),
      ...(typeof operation.italic === 'boolean' ? { italic: operation.italic } : {}),
    }
    if (!Object.keys(style).length) throw new Error('invalid_tool_input')
    return {
      op: 'set_shape_text_style',
      slide_index: operation.slide_index as number,
      shape_id: operation.shape_id,
      ...style,
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
  platform?: string
  vfs?: InMemoryVfs
  nativeMasterEditingSupported?: boolean
  verificationAuthority?: OfficePowerPointVerificationAuthority
  visualReviewer?: OfficePowerPointVisualReviewer
  presentationFlags?: PresentationVerificationFlags
  presentationTelemetry?: (event: PresentationTelemetryEvent) => void
}): AgentSkill {
  const masterXmlEditingSupported = options.platform?.toLowerCase() !== 'mac'
  const nativeMasterEditingSupported = options.nativeMasterEditingSupported !== false
  const presentation =
    options.verificationAuthority && options.presentationFlags?.verifiedCompletion !== false
      ? createOfficePowerPointVerification({
          authority: options.verificationAuthority,
          platform: options.platform,
          reviewer: options.visualReviewer,
          flags: options.presentationFlags,
          telemetry: options.presentationTelemetry,
        })
      : undefined
  options.proposals.subscribeAudit?.((event) => {
    if (!presentation) return
    if (event.kind === 'proposed') presentation.recordProposal(event)
    else if (event.kind === 'settled') presentation.recordSettlement(event)
  })
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
      'Follow the WisWork Slides workflow for presentation tasks: inspect the presentation before planning; ask the user only when a material audience, purpose, or scope decision cannot be inferred; use plan_deck before the first mutation when creating or substantially rebuilding a deck; research facts and images when needed; apply bounded edits through proposals; and call verify_slides after the approved build before reporting completion. ' +
      'PowerPoint reads are bounded. Every write creates an explicit proposal and is semantically verified after confirmation. execute_office_js accepts only a versioned declarative JSON program; JavaScript and ambient browser authority are rejected. XML tools accept only allowlisted bounded package parts.' +
      ' Prefer inspect_slide_masters and native edit_slide_master for backgrounds, theme colors, and layout inheritance. PowerPoint for Mac must never use edit_slide_master_xml.',
    tools: tools.filter(
      (tool) =>
        (masterXmlEditingSupported || tool.name !== 'edit_slide_master_xml') &&
        (nativeMasterEditingSupported ||
          !['inspect_slide_masters', 'edit_slide_master'].includes(tool.name)),
    ),
    ...(presentation ? { presentation } : {}),
    async executeTool(call, signal) {
      if (call.inputError || call.truncated)
        return failure(
          call.name,
          'invalid_tool_input',
          PROGRAM_TOOLS.has(call.name) ? invalidToolInput('program') : undefined,
        )
      try {
        assertNotCancelled(signal)
        if (call.name === 'plan_deck') {
          const plan = planDeckInput(call.input)
          return {
            output: boundedJson({
              status: 'planned',
              coreHook: plan.core_hook,
              style: plan.style,
              pages: plan.pages.map((page, index) => ({
                page: index + 1,
                title: page.title,
                type: page.type ?? 'content',
                brief: page.brief,
                layout: page.layout,
                imageQueries: page.image_queries ?? [],
              })),
            }),
            mutated: false,
            summary: `Planned ${plan.pages.length} slides`,
          }
        }
        if (presentation?.shouldSkip(call))
          return {
            output: JSON.stringify({ status: 'unchanged' }),
            mutated: false,
            summary: 'PowerPoint state already matched',
          }
        if (call.name === 'edit_slide_master_xml' && !masterXmlEditingSupported)
          return failure(call.name, 'office_api_unsupported')
        if (
          ['inspect_slide_masters', 'edit_slide_master'].includes(call.name) &&
          !nativeMasterEditingSupported
        )
          return failure(call.name, 'office_api_unsupported')
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
        if (call.name === 'inspect_slide_masters') {
          verifyInput(call.input)
          return {
            output: boundedJson(await options.adapter.inspectSlideMasters(signal)),
            mutated: false,
            summary: 'Inspected PowerPoint slide masters',
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
            verificationBinding: canonicalPowerPointVerificationBinding(call, [
              `${before.slideId}/${input.shape_id}`,
            ]),
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
          const slideIds = new Map(
            slideIndexes.map((slideIndex, index) => [slideIndex, snapshots[index]!.slideId]),
          )
          const beforeTexts = await Promise.all(
            program.operations.flatMap((operation) =>
              operation.op === 'set_shape_text'
                ? [options.adapter.readSlideText(operation.slide_index, operation.shape_id, signal)]
                : [],
            ),
          )
          const combined = snapshots.map((item) => item.fingerprint).join('|')
          const normalizedTargets = program.operations.map((operation) => {
            const slideId = slideIds.get(operation.slide_index)!
            return 'shape_id' in operation ? `${slideId}/${operation.shape_id}` : slideId
          })
          let declarativeResult: { createdShapeIds: string[]; insertedSlideId?: string } | undefined
          const proposal = options.proposals.propose({
            operation: call.name,
            toolName: call.name,
            title: input.explanation || 'Execute declarative PowerPoint operations',
            preview: { version: 1, operations: program.operations },
            impact: {
              host: 'powerpoint',
              targets: normalizedTargets,
              count: program.operations.length,
            },
            verificationBinding: canonicalPowerPointVerificationBinding(call, normalizedTargets),
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
                } else if (operation.op === 'set_shape_text_style') {
                  if (!options.adapter.readShapeTextStyle) throw new Error('office_api_unsupported')
                  const expectedStyle = program.operations
                    .filter(
                      (
                        later,
                      ): later is Extract<
                        PowerPointDeclarativeOperation,
                        { op: 'set_shape_text_style' }
                      > =>
                        later.op === 'set_shape_text_style' &&
                        later.slide_index === operation.slide_index &&
                        later.shape_id === operation.shape_id,
                    )
                    .reduce(
                      (value, later) => ({
                        ...value,
                        ...(later.color !== undefined ? { color: later.color } : {}),
                        ...(later.fontFamily !== undefined ? { fontFamily: later.fontFamily } : {}),
                        ...(later.fontSize !== undefined ? { fontSize: later.fontSize } : {}),
                        ...(later.bold !== undefined ? { bold: later.bold } : {}),
                        ...(later.italic !== undefined ? { italic: later.italic } : {}),
                      }),
                      {} as Extract<PowerPointDeclarativeOperation, { op: 'set_shape_text_style' }>,
                    )
                  await verifyPowerPointReadback(async () => {
                    const current = await options.adapter.readShapeTextStyle!(
                      operation.slide_index,
                      operation.shape_id,
                      confirmSignal,
                    )
                    return (
                      (expectedStyle.color === undefined ||
                        current.color?.toUpperCase() === expectedStyle.color.toUpperCase()) &&
                      (expectedStyle.fontFamily === undefined ||
                        current.fontFamily === expectedStyle.fontFamily) &&
                      (expectedStyle.fontSize === undefined ||
                        current.fontSize === expectedStyle.fontSize) &&
                      (expectedStyle.bold === undefined || current.bold === expectedStyle.bold) &&
                      (expectedStyle.italic === undefined ||
                        current.italic === expectedStyle.italic)
                    )
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
            },
          })
          return {
            output: boundedJson(proposal),
            mutated: false,
            summary: 'Proposed declarative PowerPoint execution',
          }
        }
        if (call.name === 'edit_slide_master') {
          const input = exactRecord(call.input, ['program', 'explanation'])
          if (
            input.explanation !== undefined &&
            (typeof input.explanation !== 'string' || input.explanation.length > 100)
          )
            throw invalidToolInput('program')
          const operations = parseMasterProgram(
            await prepareMasterProgram(input.program, options.vfs),
          )
          const operationKeys = operations.map(masterOperationKey)
          if (new Set(operationKeys).size !== operationKeys.length)
            throw invalidToolInput('program.operations')
          const before = await options.adapter.inspectSlideMasters(signal)
          const after = projectedMasterState(before, operations)
          for (const operation of operations) inverseMasterOperation(before, operation)
          const targets = [
            ...new Set(operations.map((operation) => `master:${operation.master_id}`)),
          ]
          const proposal = options.proposals.propose({
            operation: call.name,
            toolName: call.name,
            title: (input.explanation as string | undefined) || 'Edit PowerPoint slide master',
            preview: {
              operations: operations.map((operation) => ({
                ...operation,
                ...(operation.op === 'set_master_background' &&
                operation.fill.type === 'picture_or_texture'
                  ? { fill: { ...operation.fill, image_base64: '[image]' } }
                  : {}),
              })),
            },
            impact: { host: 'powerpoint', targets, count: targets.length },
            fingerprint: affectedMasterFingerprint(before, operations),
            before,
            after,
            validate: async (s) =>
              affectedMasterFingerprint(
                await options.adapter.inspectSlideMasters(s),
                operations,
              ) === affectedMasterFingerprint(before, operations),
            execute: async (s) => {
              let currentExpected = before
              const applied: Array<{
                before: PowerPointMasterState
                after: PowerPointMasterState
                inverse: PowerPointMasterOperation
                operation: PowerPointMasterOperation
              }> = []
              try {
                for (const operation of operations) {
                  const stepBefore = currentExpected
                  const nextExpected = projectedMasterState(currentExpected, [operation])
                  const inverse = inverseMasterOperation(stepBefore, operation)
                  try {
                    await options.adapter.executeMasterOperations([operation], s)
                  } catch (error) {
                    const actual = await options.adapter.inspectSlideMasters()
                    if (sameMasterOperationValue(actual, nextExpected, operation)) {
                      applied.push({ before: stepBefore, after: nextExpected, inverse, operation })
                      currentExpected = nextExpected
                    } else if (!sameMasterOperationValue(actual, currentExpected, operation)) {
                      throw new Error('office_state_uncertain', { cause: error })
                    }
                    throw error
                  }
                  const actual = await options.adapter.inspectSlideMasters(s)
                  if (!sameMasterOperationValue(actual, nextExpected, operation)) {
                    if (!sameMasterOperationValue(actual, stepBefore, operation))
                      applied.push({ before: stepBefore, after: nextExpected, inverse, operation })
                    throw new Error('office_verify_failed')
                  }
                  applied.push({ before: stepBefore, after: nextExpected, inverse, operation })
                  currentExpected = nextExpected
                }
              } catch (error) {
                for (const step of [...applied].reverse()) {
                  const actual = await options.adapter.inspectSlideMasters()
                  if (sameMasterOperationValue(actual, step.before, step.operation)) continue
                  if (!sameMasterOperationValue(actual, step.after, step.operation))
                    throw new Error('office_concurrent_change', { cause: error })
                  try {
                    await options.adapter.executeMasterOperations([step.inverse])
                  } catch (recoveryError) {
                    throw new Error('office_recovery_failed', { cause: recoveryError })
                  }
                  const restoredStep = await options.adapter.inspectSlideMasters()
                  if (!sameMasterOperationValue(restoredStep, step.before, step.operation))
                    throw new Error('office_recovery_failed', { cause: error })
                }
                const restored = await options.adapter.inspectSlideMasters()
                if (
                  affectedMasterFingerprint(restored, operations) !==
                  affectedMasterFingerprint(before, operations)
                )
                  throw new Error('office_recovery_failed', { cause: error })
                throw error
              }
            },
            verify: async (s) => {
              if (
                affectedMasterFingerprint(
                  await options.adapter.inspectSlideMasters(s),
                  operations,
                ) !== affectedMasterFingerprint(after, operations)
              )
                throw new Error('office_verify_failed')
            },
          })
          return {
            output: boundedJson(proposal),
            mutated: false,
            summary: 'Proposed native PowerPoint master edit',
          }
        }
        if (call.name === 'edit_slide_master_xml') {
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
