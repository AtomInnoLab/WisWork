import {
  encodeOfficeScreenshotResult,
  formatPresentationDesignReadinessFailure,
  parsePresentationDesignContract,
  PRESENTATION_DESIGN_CONTRACT_SCHEMA,
  PRESENTATION_DESIGN_WORKFLOW_PROMPT,
  renderPresentationDesignContract,
  transitionPresentationDesignContract,
  validatePresentationDesignReadiness,
  type AgentSkill,
  type AgentImage,
  type PresentationDesignContract,
  type ToolExecution,
} from '@wiswork/agent-core'
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
const visibleDesignDocument = (contract: PresentationDesignContract): string =>
  renderPresentationDesignContract(contract).replace(
    /\n?<!-- WISWORK_PRESENTATION_DESIGN_CONTRACT:[^\n]* -->/g,
    '',
  )

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
const slideBackgroundInput = exactObject({
  slide_index: integerField({ min: 0, max: MAX_SLIDE_INDEX }),
  color: (value: unknown) => {
    if (typeof value !== 'string' || !/^#[0-9A-Fa-f]{6}$/.test(value))
      throw new Error('invalid_tool_input')
    return value
  },
  transparency: optionalField((value: unknown) => {
    if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > 1)
      throw new Error('invalid_tool_input')
    return value
  }),
  explanation: optionalField(stringField({ maxLength: 100 })),
})
const textEditInput = exactObject({
  slide_index: integerField({ min: 0, max: MAX_SLIDE_INDEX }),
  shape_id: stringField({ minLength: 1, maxLength: 256 }),
  text: stringField({ maxLength: 12_000 }),
  explanation: optionalField(stringField({ maxLength: 50 })),
})
const slideProperties = {
  slide_index: {
    type: 'integer',
    minimum: 0,
    maximum: MAX_SLIDE_INDEX,
    description: "Zero-based slide index: the user's slide 1 is index 0.",
  },
  explanation: { type: 'string', maxLength: 50 },
} as const
const operationSlideIndex = {
  type: 'integer',
  minimum: 0,
  maximum: MAX_SLIDE_INDEX,
  description: "Zero-based slide index: the user's slide 1 is index 0.",
} as const
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
    name: 'get_presentation_state',
    description:
      'Read the bounded PowerPoint document state before planning or editing. Returns slide count, selected zero-based slide indices, supported PowerPoint API versions, and measured slideWidth/slideHeight in points when supported. Use these dimensions for layouts; screenshot pixels are not Office coordinates.',
    inputSchema: {
      type: 'object',
      properties: { explanation: { type: 'string', maxLength: 50 } },
      required: [],
      additionalProperties: false,
    },
  },
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
    name: 'review_slide_screenshot',
    description:
      'After visually inspecting the latest screenshot, record the result against every DESIGN.md acceptance ID for that slide. A failed review keeps production blocked for repair.',
    inputSchema: {
      type: 'object',
      properties: {
        ...slideProperties,
        acceptance_ids: {
          type: 'array',
          minItems: 1,
          maxItems: 20,
          items: { type: 'string', minLength: 1, maxLength: 80 },
        },
        passed: { type: 'boolean' },
        issues: {
          type: 'array',
          maxItems: 20,
          items: { type: 'string', minLength: 1, maxLength: 500 },
        },
      },
      required: ['slide_index', 'acceptance_ids', 'passed'],
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
    description:
      'Read bounded text and available current font/style readback from a shape selected by stable ID. Use the actual style to repair a failed font change without replaying the batch.',
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
    name: 'set_slide_background',
    description:
      'Propose setting one slide background to a solid color using the native PowerPoint background API.',
    inputSchema: {
      type: 'object',
      properties: {
        ...slideProperties,
        color: { type: 'string', pattern: '^#[0-9A-Fa-f]{6}$' },
        transparency: { type: 'number', minimum: 0, maximum: 1 },
      },
      required: ['slide_index', 'color'],
      additionalProperties: false,
    },
  },
  {
    name: 'ask_clarification',
    description:
      'For a whole new deck, ask one concise multiple-choice question at a time about audience, focus, style, or page count. Wait for each answer before asking the next question or continuing with plan_deck. Skip only when the user already supplied or delegated these choices.',
    inputSchema: {
      type: 'object',
      properties: {
        questions: {
          type: 'array',
          minItems: 1,
          maxItems: 1,
          items: {
            type: 'object',
            properties: {
              id: { type: 'string', minLength: 1, maxLength: 40 },
              label: { type: 'string', minLength: 1, maxLength: 300 },
              description: { type: 'string', maxLength: 300 },
              options: {
                type: 'array',
                minItems: 2,
                maxItems: 5,
                items: {
                  oneOf: [
                    { type: 'string', minLength: 1, maxLength: 120 },
                    {
                      type: 'object',
                      properties: {
                        label: { type: 'string', minLength: 1, maxLength: 120 },
                        description: { type: 'string', maxLength: 300 },
                      },
                      required: ['label'],
                      additionalProperties: false,
                    },
                  ],
                },
              },
            },
            required: ['id', 'label', 'options'],
            additionalProperties: false,
          },
        },
      },
      required: ['questions'],
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
        style: {
          type: 'string',
          minLength: 1,
          maxLength: 6_000,
          description:
            'Complete DESIGN.md body: concrete color tokens, type hierarchy, margins/grid, image treatment, density limits, layout families, and composition rules',
        },
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
              purpose: { type: 'string', minLength: 1, maxLength: 500 },
              visual: { type: 'string', minLength: 1, maxLength: 1_000 },
              evidence: {
                type: 'array',
                maxItems: 8,
                items: { type: 'string', minLength: 1, maxLength: 500 },
              },
              acceptance: {
                type: 'array',
                maxItems: 8,
                items: { type: 'string', minLength: 1, maxLength: 300 },
              },
              density: { type: 'string', enum: ['low', 'medium', 'high'] },
            },
            required: ['title', 'brief', 'layout', 'purpose', 'visual', 'acceptance', 'density'],
            additionalProperties: false,
          },
        },
        prototype_pages: {
          type: 'array',
          minItems: 1,
          maxItems: 3,
          items: { type: 'integer', minimum: 0 },
          description:
            'Zero-based indexes of the cover, representative content page, and most complex visual page; use every page when fewer than three',
        },
        contract: {
          ...PRESENTATION_DESIGN_CONTRACT_SCHEMA,
          properties: {
            ...(PRESENTATION_DESIGN_CONTRACT_SCHEMA.properties as Record<string, unknown>),
            status: { type: 'string', enum: ['draft', 'ready'] },
          },
          description:
            'Submit draft or ready only. The host owns producing/verified. Use review_slide_screenshot to record visual reviews, not plan_deck. Unknown future fields are ignored for mixed-version compatibility.',
        },
      },
      anyOf: [
        { required: ['contract'] },
        { required: ['core_hook', 'style', 'pages', 'prototype_pages'] },
      ],
      additionalProperties: false,
    },
  },
  {
    name: 'execute_office_js',
    description:
      'Execute a transaction-protected bounded declarative PowerPoint program under the PC-managed session policy. The input shape is exactly { program: { version: 1, operations: [...] }, explanation?: string }; do not place version or operations at the top level; do not stringify it and do not send JavaScript. Use snake_case fields except the bounded text-style properties. Supported operations are set_shape_text, set_shape_text_style (color/fontFamily/fontSize/bold/italic), set_shape_geometry, add_text_box, and delete_shape. Use the dedicated duplicate_slide tool for slide duplication.',
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
function equivalentPowerPointText(actual: string, expected: string): boolean {
  const normalize = (value: string) => value.replace(/\r\n|\r|\v/g, '\n')
  return normalize(actual) === normalize(expected)
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
  if (code === 'office_screenshot_unavailable') return code
  if (code === 'tool_cancelled') return 'cancelled'
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
  prepareScreenshot?: (image: AgentImage, signal?: AbortSignal) => Promise<AgentImage>
  presentationFlags?: PresentationVerificationFlags
  presentationTelemetry?: (event: PresentationTelemetryEvent) => void
}): AgentSkill & { validateImageMutation: (slideIndex: number) => string | undefined } {
  const mutationTools = new Set([
    'set_slide_background',
    'execute_office_js',
    'edit_slide_text',
    'edit_slide_xml',
    'edit_slide_chart',
    'edit_slide_master',
    'edit_slide_master_xml',
    'duplicate_slide',
  ])
  let mutationRevision = 0
  let screenshotRevision = 0
  let verificationRevision = 0
  let knownSlideCount = 0
  let canvas: { slideWidth: number; slideHeight: number; coordinateUnit: 'pt' } | undefined
  let unknownMutationPages = false
  let activeDesignContract: PresentationDesignContract | undefined
  let activeDesignContractIsModern = false
  const dirtySlideIndexes = new Set<number>()
  const builtDesignSlides = new Set<number>()
  const appliedUnverifiedDesignSlides = new Set<number>()
  const pendingDesignReviews = new Set<number>()
  const repairRequiredDesignReviews = new Set<number>()
  const failedScreenshotSlides = new Set<number>()
  const reviewRecovery = () => ({
    nextTool: failedScreenshotSlides.size
      ? 'list_slide_shapes'
      : [...pendingDesignReviews].some((index) => dirtySlideIndexes.has(index))
        ? 'screenshot_slide'
        : repairRequiredDesignReviews.size
          ? 'list_slide_shapes'
          : pendingDesignReviews.size
            ? 'review_slide_screenshot'
            : !activeDesignContract || activeDesignContract.status === 'draft'
              ? 'plan_deck'
              : builtDesignSlides.size < activeDesignContract.slides.length
                ? 'get_presentation_state'
                : 'verify_slides',
    pendingReviews: [...pendingDesignReviews]
      .sort((a, b) => a - b)
      .map((index) => ({
        slide_index: index,
        acceptance_ids:
          activeDesignContract?.slides[index]?.acceptance.map((rule) => rule.id) ?? [],
        needsScreenshot: dirtySlideIndexes.has(index),
        needsRepair: repairRequiredDesignReviews.has(index),
        ...(failedScreenshotSlides.has(index) ? { screenshotUnavailable: true } : {}),
      })),
    ...(failedScreenshotSlides.size
      ? { failedScreenshotSlideIndexes: [...failedScreenshotSlides].sort((a, b) => a - b) }
      : {}),
    instruction: failedScreenshotSlides.size
      ? 'The screenshot is unavailable for visual inspection. Inspect list_slide_shapes and read_slide_text on the affected page. Existing built or pending pages remain writable: repair the image or layout on that same page, then retry screenshot_slide and review its real image. Do not repeat an unchanged failing screenshot or claim that all writes are blocked. If the native host still cannot render, report that page as blocked; final verification remains required.'
      : pendingDesignReviews.size
        ? 'Inspect each current screenshot, then call review_slide_screenshot with its acceptance_ids and the actual result. Repair failed pages and screenshot again. Do not resubmit plan_deck to record a review.'
        : !activeDesignContract || activeDesignContract.status === 'draft'
          ? 'Submit the complete validated plan with draft or ready. Producing and verified are host-owned states.'
          : builtDesignSlides.size < activeDesignContract.slides.length
            ? 'Read the current presentation and continue the remaining production under this contract. Do not resubmit plan_deck to record progress.'
            : 'Call verify_slides and resolve remaining issues before reporting completion.',
  })
  const proposalDesignSlides = new Map<string, { indexes: number[]; scaffold: boolean }>()
  const mutationSlideIndexes = (call: { name: string; input: Record<string, unknown> }) => {
    if (['edit_slide_master', 'edit_slide_master_xml'].includes(call.name))
      return Array.from({ length: knownSlideCount }, (_, index) => index)
    if (call.name === 'execute_office_js') {
      const program = call.input.program
      if (!program || typeof program !== 'object' || Array.isArray(program)) return []
      const operations = (program as Record<string, unknown>).operations
      if (!Array.isArray(operations)) return []
      return operations.flatMap((operation) => {
        if (!operation || typeof operation !== 'object' || Array.isArray(operation)) return []
        const index = (operation as Record<string, unknown>).slide_index
        return Number.isSafeInteger(index) ? [index as number] : []
      })
    }
    const index = call.input.slide_index
    if (!Number.isSafeInteger(index)) return []
    // Duplicating reads the source but only materializes the newly inserted slide.
    // Counting the source as produced deadlocks the next batch and can report a
    // contract complete before all planned pages exist.
    return call.name === 'duplicate_slide' ? [(index as number) + 1] : [index as number]
  }
  const isMac = options.platform?.toLowerCase() === 'mac'
  const masterXmlEditingSupported = !isMac
  const nativeMasterEditingSupported = !isMac && options.nativeMasterEditingSupported !== false
  const presentation =
    !isMac &&
    options.verificationAuthority &&
    options.presentationFlags?.verifiedCompletion !== false
      ? {
          ...createOfficePowerPointVerification({
            authority: options.verificationAuthority,
            platform: options.platform,
            reviewer: options.visualReviewer,
            flags: options.presentationFlags,
            telemetry: options.presentationTelemetry,
          }),
          batchScoped: true,
        }
      : undefined
  const recordDesignMutation = (
    indexes: readonly number[],
    scaffold: boolean,
    confirmed = true,
  ) => {
    if (!activeDesignContractIsModern || !activeDesignContract) return
    if (confirmed && activeDesignContract.status === 'ready')
      activeDesignContract = transitionPresentationDesignContract(activeDesignContract, 'producing')
    else if (activeDesignContract.status === 'verified')
      activeDesignContract = { ...activeDesignContract, status: 'producing' }
    for (const index of indexes) {
      if (!scaffold) {
        appliedUnverifiedDesignSlides.delete(index)
        if (confirmed) {
          builtDesignSlides.add(index)
          repairRequiredDesignReviews.delete(index)
          failedScreenshotSlides.delete(index)
        }
        pendingDesignReviews.add(index)
      }
      dirtySlideIndexes.add(index)
    }
    mutationRevision++
  }
  const designProductionError = (
    indexes: readonly number[],
    inserting: boolean,
  ): string | undefined => {
    if (!activeDesignContractIsModern || !activeDesignContract) return
    const targets = [...new Set(indexes)]
    if (
      targets.length === 0 ||
      targets.some((index) => index < 0 || index >= activeDesignContract!.slides.length)
    )
      return 'design_contract_scope_mismatch'
    // Repair is not expansion. Previously reviewed pages must be writable again;
    // confirmation invalidates their screenshot/review just like any other write.
    if (!inserting && targets.every((index) => builtDesignSlides.has(index))) return
    if (inserting && targets.some((index) => builtDesignSlides.has(index)))
      return 'design_contract_batch_mismatch'
    if (targets.every((index) => pendingDesignReviews.has(index))) return
    const newTargets = targets.filter((index) => !builtDesignSlides.has(index))
    const prototypes = new Set(activeDesignContract.prototypePages.map((number) => number - 1))
    if (![...prototypes].every((index) => builtDesignSlides.has(index)))
      return newTargets.every((index) => prototypes.has(index))
        ? undefined
        : 'design_contract_prototype_required'
    if (new Set([...pendingDesignReviews, ...targets]).size > 3)
      return 'design_contract_review_required'
    // Office exposes slide creation as one confirmed duplicate proposal at a time.
    // Permit those sequential writes; pendingDesignReviews still caps expansion and
    // forces screenshot review before the agent can move beyond the current batch.
    if (targets.length > 3) return 'design_contract_batch_size'
  }
  options.proposals.subscribeAudit?.((event) => {
    if (event.kind === 'proposed' && event.powerPointMutation)
      proposalDesignSlides.set(event.id, {
        ...event.powerPointMutation,
        indexes: [...event.powerPointMutation.indexes],
      })
    if (event.kind === 'settled') {
      const mutation = proposalDesignSlides.get(event.id)
      proposalDesignSlides.delete(event.id)
      if (event.status === 'confirmed' && mutation)
        recordDesignMutation(mutation.indexes, mutation.scaffold)
      else if (mutation && ['failed', 'applied_unverified'].includes(event.status)) {
        // Verification can fail after Office has applied some or all operations. Keep the
        // page unbuilt until a fresh screenshot review confirms an applied-unverified write.
        recordDesignMutation(mutation.indexes, mutation.scaffold, false)
        if (
          event.status === 'applied_unverified' &&
          event.safeCode !== 'office_write_pending' &&
          !mutation.scaffold
        )
          for (const index of mutation.indexes) {
            appliedUnverifiedDesignSlides.add(index)
            // A known applied repair still needs its new screenshot. A pending or
            // failed write cannot stand in for a repair of the rejected image.
            repairRequiredDesignReviews.delete(index)
            failedScreenshotSlides.delete(index)
          }
      }
    }
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
      powerPointMutation: {
        indexes: mutationSlideIndexes({ name: toolName, input: { slide_index: slideIndex } }),
        scaffold: false,
      },
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
    validateImageMutation: (slideIndex) => {
      if (!Number.isSafeInteger(slideIndex) || slideIndex < 0) return 'invalid_tool_input'
      if (knownSlideCount > 0 && slideIndex >= knownSlideCount) return 'invalid_tool_input'
      const error = designProductionError([slideIndex], false)
      return error === 'design_contract_review_required'
        ? boundedJson({ error, ...reviewRecovery() })
        : error
    },
    repeatFinalResponseCorrection: true,
    systemPrompt:
      `${PRESENTATION_DESIGN_WORKFLOW_PROMPT}\n` +
      'Follow the same complete workflow as WisWork Slides in this agent run: understand the document with get_presentation_state and bounded reads, and inspect the presentation before planning; for a new deck, you must call ask_clarification for missing audience, focus, style, and page-count choices unless the user already supplied or delegated them; use plan_deck before the first mutation to record the narrative and visual plan; run web_search and image_search for needed facts and visuals; implement the complete plan with bounded slide edits; screenshot every created or changed slide and inspect the native images; repair concrete clipping, overlap, hierarchy, spacing, contrast, and balance defects; screenshot every repaired slide again; then call verify_slides after the approved build before reporting completion. A screenshot call alone is not a visual pass: inspect its image and keep the screenshot-repair-screenshot loop in this same run until the checked pages are satisfactory or a concrete blocker remains. Never end the run expecting another user message or host post-processing to finish the deck. All slide_index values are zero-based, so the user’s first slide is index 0. Never replace that tool call with prose questions; the host renders its model-authored questions as interactive feedback and returns the answers so you can continue the same task. Emit a concise user-visible progress note before every tool batch, explaining the current design decision and next action without revealing private chain-of-thought. ' +
      'PowerPoint reads are bounded. Every write creates an explicit proposal and is semantically verified after confirmation. execute_office_js accepts only a versioned declarative JSON program; JavaScript and ambient browser authority are rejected. XML tools accept only allowlisted bounded package parts.' +
      ' Use the measured slideWidth and slideHeight from get_presentation_state for every layout; coordinates are points, not screenshot pixels. Never assume a 720x405 or 960x540 canvas. If dimensions are unavailable, obtain the actual size before positioning content. A full-bleed image must cover the actual canvas, not only its upper-left area. Plan text and imagery together: preserve image proportions, leave deliberate clear space for titles, and use a contrasting text panel when the photo is too busy or bright. Inspect the entire screenshot including right and bottom edges; accidental white bands, blank planned pages, distorted images, and unreadable text over photos fail visual review. Do not change font families across a batch merely for styling; use the current family when font readback fails and repair size, color, spacing, and background separately.' +
      ' Submit plan_deck with draft or ready only; producing and verified are host-owned states. After inspecting each screenshot call review_slide_screenshot with its acceptance_ids; screenshot_slide and verify_slides do not register a visual review. Do not resubmit plan_deck to record a review.' +
      ' ' +
      (isMac
        ? 'On PowerPoint for Mac, build new decks with slide-level tools and never call slide-master tools; finish with screenshot_slide and verify_slides.'
        : 'Prefer inspect_slide_masters and native edit_slide_master for backgrounds, theme colors, and layout inheritance.'),
    tools: tools.filter(
      (tool) =>
        (masterXmlEditingSupported || tool.name !== 'edit_slide_master_xml') &&
        (nativeMasterEditingSupported ||
          !['inspect_slide_masters', 'edit_slide_master'].includes(tool.name)),
    ),
    buildContext: () =>
      (canvas ? `<presentation canvas>\n${boundedJson(canvas)}\n</presentation canvas>\n` : '') +
      (activeDesignContract
        ? `<active presentation design contract>\n${boundedJson(activeDesignContract)}\n</active presentation design contract>\n<presentation review progress>\n${boundedJson(reviewRecovery())}\n</presentation review progress>`
        : ''),
    reviewFinalResponse(context) {
      if (!context.mutated) return undefined
      if (unknownMutationPages)
        return '[System correction] Read the presentation state, then screenshot every slide affected by the master change.'
      if (failedScreenshotSlides.size) return `[System correction] ${boundedJson(reviewRecovery())}`
      if (dirtySlideIndexes.size)
        return `[System correction] Continue the WisWork Slides quality loop now: call screenshot_slide for every created or changed slide (${[...dirtySlideIndexes].map((index) => index + 1).join(', ')}), inspect each native image, repair concrete defects, and screenshot each repaired slide again before finishing.`
      if (pendingDesignReviews.size) return `[System correction] ${boundedJson(reviewRecovery())}`
      if (verificationRevision < mutationRevision)
        return '[System correction] The changed slides have been visually inspected. Call verify_slides now and resolve any remaining failure before reporting completion.'
      return undefined
    },
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
        if (call.name === 'ask_clarification')
          return failure(call.name, 'questionnaire_unavailable')
        if (call.name === 'get_presentation_state') {
          verifyInput(call.input)
          const state = await options.adapter.getPresentationState(signal)
          knownSlideCount = state.slideCount
          canvas =
            state.coordinateUnit === 'pt' &&
            typeof state.slideWidth === 'number' &&
            typeof state.slideHeight === 'number' &&
            Number.isFinite(state.slideWidth) &&
            Number.isFinite(state.slideHeight) &&
            state.slideWidth > 0 &&
            state.slideHeight > 0
              ? {
                  slideWidth: state.slideWidth,
                  slideHeight: state.slideHeight,
                  coordinateUnit: 'pt',
                }
              : undefined
          if (unknownMutationPages) {
            for (let index = 0; index < knownSlideCount; index++) dirtySlideIndexes.add(index)
            unknownMutationPages = false
          }
          return {
            output: boundedJson(state),
            mutated: false,
            summary: 'Read PowerPoint presentation state',
          }
        }
        if (call.name === 'plan_deck') {
          let contract: PresentationDesignContract
          try {
            contract = parsePresentationDesignContract(call.input.contract ?? call.input)
          } catch {
            return failure(call.name, 'invalid_tool_input')
          }
          const readiness = validatePresentationDesignReadiness(contract)
          if ('contract' in call.input) {
            if (!['draft', 'ready'].includes(contract.status))
              return failure(
                call.name,
                boundedJson({
                  error: 'design_contract_invalid_status',
                  allowedStatuses: ['draft', 'ready'],
                  ...reviewRecovery(),
                }),
              )
            if (contract.status === 'ready' && !readiness.ready)
              return failure(
                call.name,
                `design_contract_not_ready: ${formatPresentationDesignReadinessFailure(contract, readiness.issues)}`,
              )
          }
          const unchanged =
            'contract' in call.input &&
            activeDesignContractIsModern &&
            activeDesignContract &&
            activeDesignContract.status !== 'draft' &&
            contract.status === 'ready' &&
            JSON.stringify({ ...contract, status: activeDesignContract.status }) ===
              JSON.stringify(activeDesignContract)
          if (unchanged) {
            contract = activeDesignContract!
          } else {
            activeDesignContract = contract
            activeDesignContractIsModern = 'contract' in call.input
            builtDesignSlides.clear()
            appliedUnverifiedDesignSlides.clear()
            pendingDesignReviews.clear()
            repairRequiredDesignReviews.clear()
            failedScreenshotSlides.clear()
            proposalDesignSlides.clear()
          }
          return {
            output: boundedJson({
              status: contract.status,
              revision: contract.revision,
              designMd: activeDesignContractIsModern
                ? visibleDesignDocument(contract)
                : renderPresentationDesignContract(contract),
              coreHook: contract.narrative.coreHook,
              prototypePages: contract.prototypePages.map((number) => number - 1),
              ...(activeDesignContractIsModern
                ? {}
                : {
                    pages: contract.slides.map((slide) => ({
                      page: slide.number,
                      title: slide.title,
                      brief: slide.claim,
                      layout: slide.layoutFamily,
                      purpose: slide.role,
                      visual: slide.visualRoute,
                      evidence: slide.evidence,
                      acceptance: slide.acceptance.map((rule) => rule.criterion),
                      density: slide.density,
                    })),
                  }),
            }),
            mutated: false,
            summary: `Planned ${contract.slides.length} slides`,
          }
        }
        if (presentation?.shouldSkip(call))
          return {
            output: JSON.stringify({ status: 'unchanged' }),
            mutated: false,
            summary: 'PowerPoint state already matched',
          }
        if (activeDesignContractIsModern && call.name === 'execute_office_js') {
          const input = declarativeInput(call.input, { slide: false, explanationMax: 100 })
          const program = parseDeclarativeProgram(input.code, parsePowerPointOperation)
          if (program.operations.some((operation) => operation.op === 'duplicate_slide'))
            return failure(
              call.name,
              JSON.stringify({
                error: 'invalid_tool_input',
                instruction:
                  'Use the dedicated duplicate_slide tool so inserted pages follow the design contract.',
              }),
            )
        }
        let scaffolding = false
        if (mutationTools.has(call.name)) {
          const indexes = mutationSlideIndexes(call)
          const inputIndexes =
            call.name === 'duplicate_slide' ? [Number(call.input.slide_index)] : indexes
          if (
            knownSlideCount > 0 &&
            inputIndexes.some((index) => index < 0 || index >= knownSlideCount)
          )
            return failure(call.name, 'invalid_tool_input')
          const productionError = designProductionError(indexes, call.name === 'duplicate_slide')
          const contract = activeDesignContract
          // Append only the unbuilt placeholders needed to reach a later prototype.
          scaffolding =
            productionError === 'design_contract_prototype_required' &&
            call.name === 'duplicate_slide' &&
            contract !== undefined &&
            ['ready', 'producing'].includes(contract.status) &&
            knownSlideCount > 0 &&
            call.input.slide_index === knownSlideCount - 1 &&
            indexes[0]! < Math.max(...contract.prototypePages) - 1
          if (productionError && !scaffolding)
            return failure(
              call.name,
              productionError === 'design_contract_review_required'
                ? boundedJson({ error: productionError, ...reviewRecovery() })
                : productionError,
            )
          if (!activeDesignContractIsModern) {
            mutationRevision++
            for (const index of indexes) dirtySlideIndexes.add(index)
          }
          if (!indexes.length && ['edit_slide_master', 'edit_slide_master_xml'].includes(call.name))
            unknownMutationPages = true
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
          const capturedMutation = mutationRevision
          const capturedContract = activeDesignContract
          const result = await options.adapter.screenshotSlide(input.slide_index, signal)
          assertNotCancelled(signal)
          if (result.mime !== 'image/png' || !validPng(result.base64))
            throw new Error('office_read_failed')
          const modelImage = options.prepareScreenshot
            ? await options.prepareScreenshot(result, signal)
            : result
          assertNotCancelled(signal)
          if (capturedMutation !== mutationRevision || capturedContract !== activeDesignContract)
            throw new Error('office_screenshot_unavailable')
          const output = boundedJson({
            mime: modelImage.mime,
            bytes: base64Bytes(modelImage.base64),
            fingerprint: fingerprint(modelImage.base64),
            visualAvailableToModel: true,
            ...(activeDesignContractIsModern && activeDesignContract
              ? {
                  designRevision: activeDesignContract.revision,
                  designStatus: activeDesignContract.status,
                  acceptanceIds:
                    activeDesignContract.slides[input.slide_index]?.acceptance.map(
                      (rule) => rule.id,
                    ) ?? [],
                  designMd: visibleDesignDocument(activeDesignContract),
                  nextTool: pendingDesignReviews.has(input.slide_index)
                    ? 'review_slide_screenshot'
                    : 'verify_slides',
                }
              : {}),
          })
          const modelContent = [{ type: 'image' as const, image: modelImage }]
          // Production preparation must fit the actual Relay envelope before clearing the gate.
          if (options.prepareScreenshot) encodeOfficeScreenshotResult(output, modelContent)
          screenshotRevision = mutationRevision
          dirtySlideIndexes.delete(input.slide_index)
          failedScreenshotSlides.delete(input.slide_index)
          return {
            output,
            modelContent,
            display: {
              kind: 'images',
              items: [{ url: `data:${result.mime};base64,${result.base64}` }],
            },
            mutated: false,
            summary:
              activeDesignContractIsModern && activeDesignContract
                ? `Rendered PowerPoint slide · DESIGN r${activeDesignContract.revision} ${activeDesignContract.status}`
                : 'Rendered PowerPoint slide',
          }
        }
        if (call.name === 'review_slide_screenshot') {
          const slideIndex = Number(call.input.slide_index)
          if (!Number.isSafeInteger(slideIndex) || slideIndex < 0)
            return failure(call.name, 'invalid_tool_input')
          const input = { slide_index: slideIndex }
          if (!activeDesignContractIsModern || !activeDesignContract)
            return failure(call.name, 'design_contract_required')
          const expected =
            activeDesignContract.slides[input.slide_index]?.acceptance.map((rule) => rule.id) ?? []
          const supplied = Array.isArray(call.input.acceptance_ids)
            ? call.input.acceptance_ids.map(String)
            : []
          if (
            supplied.length !== expected.length ||
            supplied.some((id, index) => id !== expected[index])
          )
            return failure(call.name, 'design_contract_acceptance_mismatch')
          const alreadyReviewed =
            !pendingDesignReviews.has(input.slide_index) &&
            builtDesignSlides.has(input.slide_index) &&
            !dirtySlideIndexes.has(input.slide_index)
          if (alreadyReviewed && call.input.passed === true)
            return {
              output: boundedJson({
                status: 'already_reviewed',
                slide: input.slide_index + 1,
                revision: activeDesignContract.revision,
              }),
              mutated: false,
              summary: `PowerPoint slide already reviewed · DESIGN r${activeDesignContract.revision}`,
            }
          if (!pendingDesignReviews.has(input.slide_index) && !alreadyReviewed)
            return failure(call.name, 'design_contract_review_not_pending')
          if (dirtySlideIndexes.has(input.slide_index))
            return failure(call.name, 'design_contract_screenshot_required')
          if (call.input.passed !== true) {
            const issues = Array.isArray(call.input.issues)
              ? call.input.issues.map(String)
              : ['Repair and re-screenshot this slide']
            pendingDesignReviews.add(input.slide_index)
            repairRequiredDesignReviews.add(input.slide_index)
            appliedUnverifiedDesignSlides.delete(input.slide_index)
            if (activeDesignContract.status === 'verified')
              activeDesignContract = { ...activeDesignContract, status: 'producing' }
            return {
              output: boundedJson({
                status: 'needs_repair',
                slide: input.slide_index + 1,
                revision: activeDesignContract.revision,
                acceptanceIds: expected,
                issues,
                nextTool: 'list_slide_shapes',
                instruction:
                  'Repair the reported visual issues, then take a new screenshot before reviewing this slide again.',
              }),
              mutated: false,
              summary: `PowerPoint slide needs repair · DESIGN r${activeDesignContract.revision}`,
            }
          }
          if (repairRequiredDesignReviews.has(input.slide_index))
            return {
              output: boundedJson({
                status: 'repair_required',
                slide: input.slide_index + 1,
                revision: activeDesignContract.revision,
                acceptanceIds: expected,
                nextTool: 'list_slide_shapes',
                instruction:
                  'Apply a repair for the previously reported issues, then take a new screenshot before reviewing this slide again.',
              }),
              mutated: false,
              summary: `PowerPoint slide still needs repair · DESIGN r${activeDesignContract.revision}`,
            }
          pendingDesignReviews.delete(input.slide_index)
          if (appliedUnverifiedDesignSlides.delete(input.slide_index)) {
            builtDesignSlides.add(input.slide_index)
            if (activeDesignContract.status === 'ready')
              activeDesignContract = transitionPresentationDesignContract(
                activeDesignContract,
                'producing',
              )
          }
          return {
            output: boundedJson({
              status: 'passed',
              slide: input.slide_index + 1,
              revision: activeDesignContract.revision,
              acceptanceIds: expected,
            }),
            mutated: false,
            summary: `Reviewed PowerPoint slide · DESIGN r${activeDesignContract.revision}`,
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
          const text = await options.adapter.readSlideText(
            input.slide_index,
            input.shape_id,
            signal,
          )
          let textStyle
          try {
            textStyle = await options.adapter.readShapeTextStyle?.(
              input.slide_index,
              input.shape_id,
              signal,
            )
          } catch (error) {
            if (signal?.aborted) throw error
            // Mixed/unsupported font properties must not make readable text unavailable.
          }
          return {
            output: boundedJson({ ...text, ...(textStyle ? { textStyle } : {}) }),
            mutated: false,
            summary: 'Read PowerPoint text',
          }
        }
        if (call.name === 'verify_slides') {
          verifyInput(call.input)
          const verified = await options.adapter.verifySlides(signal)
          const designClean =
            !verified.truncated &&
            verified.slides.every(
              (slide) =>
                !slide.shapesTruncated &&
                !slide.overlapsTruncated &&
                slide.overflows.length === 0 &&
                slide.overlaps.length === 0,
            )
          const designComplete =
            !activeDesignContractIsModern ||
            (activeDesignContract !== undefined &&
              builtDesignSlides.size === activeDesignContract.slides.length &&
              pendingDesignReviews.size === 0)
          if (activeDesignContractIsModern && (!designClean || !designComplete))
            return failure(
              call.name,
              boundedJson({
                error: !designComplete
                  ? 'design_contract_production_incomplete'
                  : 'design_contract_verification_failed',
                ...reviewRecovery(),
                ...(!designClean
                  ? {
                      nextTool: 'list_slide_shapes',
                      instruction:
                        'Repair the reported slide/shape geometry, then screenshot and review each changed page again. Use real slide backgrounds instead of overlapping background rectangles. Do not repeat verification without fixing the reported defects.',
                    }
                  : {}),
                // Omit unrelated shape inventories; preserve bounded, actionable failures.
                verification: {
                  slideWidth: verified.slideWidth,
                  slideHeight: verified.slideHeight,
                  truncated: verified.truncated,
                  slides: verified.slides.map(({ shapes: _shapes, ...slide }) => slide),
                },
                unbuiltSlideIndexes: activeDesignContract?.slides.flatMap((_, index) =>
                  builtDesignSlides.has(index) ? [] : [index],
                ),
              }),
            )
          if (
            dirtySlideIndexes.size === 0 &&
            screenshotRevision === mutationRevision &&
            designClean &&
            designComplete
          ) {
            verificationRevision = mutationRevision
            if (activeDesignContractIsModern && activeDesignContract?.status === 'producing')
              activeDesignContract = transitionPresentationDesignContract(
                activeDesignContract,
                'verified',
              )
          }
          return {
            output: boundedJson(
              activeDesignContractIsModern && activeDesignContract?.status === 'verified'
                ? {
                    verification: verified,
                    status: activeDesignContract.status,
                    revision: activeDesignContract.revision,
                    designMd: visibleDesignDocument(activeDesignContract),
                  }
                : verified,
            ),
            mutated: false,
            summary: 'Verified PowerPoint slides',
          }
        }
        if (call.name === 'set_slide_background') {
          const input = slideBackgroundInput(call.input)
          if (!options.adapter.readSlideBackground || !options.adapter.setSlideBackground)
            return failure(call.name, 'office_api_unsupported')
          const before = await options.adapter.readSlideBackground(input.slide_index, signal)
          if (
            before.type.toLowerCase() !== 'solid' ||
            !before.backgroundColor ||
            before.transparency === undefined
          )
            return failure(call.name, 'office_api_unsupported')
          const color = input.color.toUpperCase()
          const transparency = input.transparency ?? 0
          const proposal = options.proposals.propose({
            powerPointMutation: { indexes: mutationSlideIndexes(call), scaffold: scaffolding },
            operation: call.name,
            toolName: call.name,
            title: input.explanation || 'Set slide background',
            preview: { slideIndex: input.slide_index, color, transparency },
            impact: {
              host: 'powerpoint',
              targets: [`${before.slideId}/background`],
              count: 1,
            },
            verificationBinding: canonicalPowerPointVerificationBinding(call, [
              `${before.slideId}/background`,
            ]),
            fingerprint: fingerprint(JSON.stringify(before)),
            before,
            after: { slideId: before.slideId, type: 'Solid', backgroundColor: color, transparency },
            validate: async (s) =>
              fingerprint(
                JSON.stringify(await options.adapter.readSlideBackground!(input.slide_index, s)),
              ) === fingerprint(JSON.stringify(before)),
            execute: async (s) =>
              options.adapter.setSlideBackground!(input.slide_index, color, transparency, s),
            verify: async (s) => {
              const current = await options.adapter.readSlideBackground!(input.slide_index, s)
              if (
                current.slideId !== before.slideId ||
                current.backgroundColor?.toUpperCase() !== color ||
                current.transparency !== transparency
              ) {
                await options.adapter.setSlideBackground!(
                  input.slide_index,
                  before.backgroundColor!,
                  before.transparency!,
                )
                const restored = await options.adapter.readSlideBackground!(input.slide_index)
                if (
                  restored.backgroundColor?.toUpperCase() !==
                    before.backgroundColor!.toUpperCase() ||
                  restored.transparency !== before.transparency
                )
                  throw new Error('office_recovery_failed')
                throw new Error('office_verify_failed')
              }
            },
          })
          return {
            output: boundedJson(proposal),
            mutated: false,
            summary: 'Proposed PowerPoint slide background',
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
            powerPointMutation: { indexes: mutationSlideIndexes(call), scaffold: scaffolding },
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
          const scaffoldContract = scaffolding ? activeDesignContract : undefined
          await options.adapter.verifySlides(signal)
          const snapshot = await options.adapter.snapshotSlide(input.slide_index, signal)
          let insertedSlideId: string | undefined
          const proposal = options.proposals.propose({
            powerPointMutation: { indexes: mutationSlideIndexes(call), scaffold: scaffolding },
            operation: 'duplicate_slide',
            toolName: call.name,
            title: input.explanation || 'Duplicate slide',
            preview: {
              slideIndex: input.slide_index,
              slideId: snapshot.slideId,
              ...(scaffolding
                ? {
                    scaffold: true,
                    instruction:
                      'This placeholder is not produced. Fill and review it after the prototype pages.',
                  }
                : {}),
            },
            impact: { host: 'powerpoint', targets: [snapshot.slideId], count: 1 },
            fingerprint: snapshot.fingerprint,
            before: snapshot,
            validate: async (confirmSignal) =>
              (await options.adapter.snapshotSlide(input.slide_index, confirmSignal))
                .fingerprint === snapshot.fingerprint &&
              (!scaffolding ||
                (activeDesignContract === scaffoldContract &&
                  (await options.adapter.getPresentationState(confirmSignal)).slideCount ===
                    input.slide_index + 1)),
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
              // Keep the cached bounds in step with this verified insertion. Zero means unknown.
              if (knownSlideCount > 0) knownSlideCount++
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
            throw invalidToolInput('program.operations')
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
            throw invalidToolInput('program.operations')
          await options.adapter.verifySlides(signal)
          const slideIndexes = [
            ...new Set(program.operations.map((operation) => operation.slide_index)),
          ]
          if (slideIndexes.length > 8) throw invalidToolInput('program.operations')
          const snapshots = await Promise.all(
            slideIndexes.map((index) => options.adapter.snapshotSlide(index, signal)),
          )
          const slideIds = new Map(
            slideIndexes.map((slideIndex, index) => [slideIndex, snapshots[index]!.slideId]),
          )
          const shapeOperations = program.operations.filter(
            (operation) => 'shape_id' in operation && operation.op !== 'set_shape_text',
          )
          const shapeSlides = [
            ...new Set(shapeOperations.map((operation) => operation.slide_index)),
          ]
          const shapesBySlide = new Map(
            await Promise.all(
              shapeSlides.map(async (slideIndex) => {
                const current = await options.adapter.listSlideShapes(slideIndex, signal)
                return [slideIndex, new Set(current.shapes.map((shape) => shape.id))] as const
              }),
            ),
          )
          if (
            shapeOperations.some(
              (operation) => !shapesBySlide.get(operation.slide_index)?.has(operation.shape_id),
            )
          )
            throw invalidToolInput('program.operations')
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
            powerPointMutation: { indexes: mutationSlideIndexes(call), scaffold: scaffolding },
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
              const verifyOperationReadback = async (
                index: number,
                operation: PowerPointDeclarativeOperation,
                read: () => Promise<string | undefined>,
              ) => {
                let mismatch: string | undefined
                try {
                  await verifyPowerPointReadback(async () => {
                    mismatch = await read()
                    return mismatch === undefined
                  }, confirmSignal)
                } catch (error) {
                  if (error instanceof Error && error.message === 'office_verify_failed')
                    Object.assign(error, {
                      debugInfo: {
                        errorLocation: `PowerPoint.operations.${index}.${operation.op}.${mismatch ?? 'readback'}`,
                      },
                    })
                  throw error
                }
              }
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
                  await verifyOperationReadback(operationIndex, operation, async () => {
                    const current = await options.adapter.readSlideText(
                      operation.slide_index,
                      operation.shape_id,
                      confirmSignal,
                    )
                    return equivalentPowerPointText(current.text, operation.text)
                      ? undefined
                      : 'text'
                  })
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
                  await verifyOperationReadback(operationIndex, operation, async () => {
                    const current = await options.adapter.readShapeTextStyle!(
                      operation.slide_index,
                      operation.shape_id,
                      confirmSignal,
                    )
                    const normalizeFontFamily = (value: string | undefined) =>
                      value?.trim().replace(/\s+/g, ' ').toLocaleLowerCase()
                    if (
                      expectedStyle.color !== undefined &&
                      current.color?.toUpperCase() !== expectedStyle.color.toUpperCase()
                    )
                      return 'color'
                    if (
                      expectedStyle.fontFamily !== undefined &&
                      normalizeFontFamily(current.fontFamily) !==
                        normalizeFontFamily(expectedStyle.fontFamily)
                    )
                      return 'fontFamily'
                    if (
                      expectedStyle.fontSize !== undefined &&
                      !(
                        current.fontSize !== undefined &&
                        Math.abs(current.fontSize - expectedStyle.fontSize) <= 0.1
                      )
                    )
                      return 'fontSize'
                    if (expectedStyle.bold !== undefined && current.bold !== expectedStyle.bold)
                      return 'bold'
                    if (
                      expectedStyle.italic !== undefined &&
                      current.italic !== expectedStyle.italic
                    )
                      return 'italic'
                    return undefined
                  })
                } else if (operation.op !== 'duplicate_slide') {
                  if (operation.op === 'add_text_box') {
                    const createdShapeId = declarativeResult?.createdShapeIds[createdShapeIndex++]
                    await verifyOperationReadback(operationIndex, operation, async () => {
                      const current = await options.adapter.listSlideShapes(
                        operation.slide_index,
                        confirmSignal,
                      )
                      const shape = current.shapes.find((item) => item.id === createdShapeId)
                      if (!shape) return 'shape_id'
                      const geometryMismatch = (['left', 'top', 'width', 'height'] as const).find(
                        (property) => !sameGeometry(shape[property], operation[property]),
                      )
                      if (geometryMismatch) return geometryMismatch
                      const text = await options.adapter.readSlideText(
                        operation.slide_index,
                        shape.id,
                        confirmSignal,
                      )
                      return equivalentPowerPointText(text.text, operation.text)
                        ? undefined
                        : 'text'
                    })
                    continue
                  }
                  await verifyOperationReadback(operationIndex, operation, async () => {
                    const current = await options.adapter.listSlideShapes(
                      operation.slide_index,
                      confirmSignal,
                    )
                    const shape = current.shapes.find((item) => item.id === operation.shape_id)
                    if (operation.op === 'delete_shape') return shape ? 'exists' : undefined
                    if (!shape) return 'shape_id'
                    return (['left', 'top', 'width', 'height'] as const).find(
                      (property) => !sameGeometry(shape[property], operation[property]),
                    )
                  })
                }
              }
              if (program.operations[0]?.op === 'duplicate_slide') {
                const operation = program.operations[0]
                const insertedSlideId = declarativeResult?.insertedSlideId
                if (!insertedSlideId) throw new Error('office_verify_failed')
                await verifyOperationReadback(0, operation, async () => {
                  const inserted = await options.adapter.listSlideShapes(
                    operation.slide_index + 1,
                    confirmSignal,
                  )
                  return inserted.slideId === insertedSlideId ? undefined : 'slide_id'
                })
                if (knownSlideCount > 0) knownSlideCount++
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
            powerPointMutation: { indexes: mutationSlideIndexes(call), scaffold: scaffolding },
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
        if (
          call.name === 'screenshot_slide' &&
          ['office_read_failed', 'office_screenshot_unavailable'].includes(code)
        ) {
          const index = Number(call.input.slide_index)
          failedScreenshotSlides.add(index)
          dirtySlideIndexes.add(index)
          if (builtDesignSlides.has(index) || pendingDesignReviews.has(index)) {
            pendingDesignReviews.add(index)
            if (activeDesignContract?.status === 'verified')
              activeDesignContract = { ...activeDesignContract, status: 'producing' }
          }
          return failure(
            call.name,
            boundedJson({
              error: 'office_read_failed',
              reason: 'office_screenshot_unavailable',
              ...reviewRecovery(),
              slide_index: index,
              repairAllowed: designProductionError([index], false) === undefined,
              visualAvailableToModel: false,
            }),
            error,
          )
        }
        return failure(call.name, code, error)
      }
    },
  }
}
