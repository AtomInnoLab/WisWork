import { MASTER_PATTERN_TYPES, masterOperationKey } from './presentation-master-program.js'
import type { AgentSkill, AgentToolDef, ToolExecution } from '@wiswork/agent-core'
import { parsePresentationDeck, PRESENTATION_DECK_SCHEMA } from '@wiswork/pptx-engine/presentation'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { exactObject, integerField, optionalField, stringField } from '../../agent/tool-schema.js'
import { parseDeclarativeProgram } from '../shared/declarative-program.js'
import { readBoundedImage } from '../shared/import-media.js'
import type { InMemoryVfs } from '../shared/vfs.js'
import type { PowerPointAdapter } from './browser-powerpoint-adapter.js'
import {
  MAX_POWERPOINT_RESULT_BYTES,
  type PowerPointDeclarativeOperation,
  type PowerPointMasterOperation,
} from './browser-powerpoint-adapter.js'
import {
  captureChartValuePackageEdit,
  presentationPackageDigest,
  verifyImportedPowerPointPackage,
  verifyPowerPointPackageInputs,
  type PackageEditKind,
  type XmlReplacement,
} from './powerpoint-package.js'
import { updatePowerPointChartDataPackage } from './presentation-chart-source-package.js'
import {
  saveChartPackageBackup,
  readChartPackageBackup,
  describePagePackageBackup,
  cleanupUncommittedChartPackageBackup,
} from './presentation-chart-backup.js'
import {
  validatePresentationExistingChartChange,
  type PresentationExistingChartChange,
} from './presentation-existing-chart.js'
import { createPresentationNativeAddRestoration } from './presentation-native-add-restoration.js'
import { createPresentationNativeAddRelease } from './presentation-native-add-release.js'
import type { PresentationExistingPageChange } from './presentation-existing-page.js'
import { createPresentationNativeAddExecution } from './presentation-native-add-execution.js'
import { createPresentationNativeAddProposal } from './presentation-native-add-proposal.js'
import type {
  PresentationExistingBatch,
  NativeAddOperation,
} from './presentation-existing-batch.js'
import { officeOperationsForSlideIR } from './presentation-office-ir.js'

const MAX_SLIDE_INDEX = 100_000
const MAX_CODE = 32 * 1024
const MAX_SCREENSHOT_BYTES = 4 * 1024 * 1024
const PROGRAM_TOOLS = new Set([
  'execute_office_js',
  'add_slide_ir_objects',
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
const verifyInput = exactObject({
  slide_index: optionalField(integerField({ min: 0, max: MAX_SLIDE_INDEX })),
  explanation: optionalField(stringField({ maxLength: 50 })),
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
              fontFace: { type: 'string', minLength: 1, maxLength: 128 },
              fontSize: { type: 'number', minimum: 6, maximum: 96 },
              color: { type: 'string', pattern: '^[0-9A-Fa-f]{6}$' },
              bold: { type: 'boolean' },
              align: { type: 'string', enum: ['left', 'center', 'right'] },
              margin: { type: 'number', minimum: 0, maximum: 72 },
              verticalAlignment: { type: 'string', enum: ['top', 'middle', 'bottom'] },
              ...geometryProperties,
            },
            ['op', 'slide_index', 'name', 'text', 'left', 'top', 'width', 'height'],
          ),
          exactOperation(
            {
              op: { type: 'string', enum: ['add_geometric_shape'] },
              slide_index: operationSlideIndex,
              name: { type: 'string', minLength: 1, maxLength: 256 },
              shape: { type: 'string', enum: ['rect', 'ellipse', 'roundRect'] },
              fill: { type: 'string', pattern: '^[0-9A-Fa-f]{6}$' },
              lineColor: { type: 'string', pattern: '^[0-9A-Fa-f]{6}$' },
              ...geometryProperties,
            },
            [
              'op',
              'slide_index',
              'name',
              'shape',
              'fill',
              'lineColor',
              'left',
              'top',
              'width',
              'height',
            ],
          ),
          exactOperation(
            {
              op: { type: 'string', enum: ['add_native_table'] },
              slide_index: operationSlideIndex,
              name: { type: 'string', minLength: 1, maxLength: 256 },
              rows: {
                type: 'array',
                minItems: 1,
                maxItems: 20,
                items: {
                  type: 'array',
                  minItems: 1,
                  maxItems: 12,
                  items: { type: 'string', maxLength: 256 },
                },
              },
              fontFace: { type: 'string', minLength: 1, maxLength: 128 },
              fontSize: { type: 'number', minimum: 6, maximum: 48 },
              color: { type: 'string', pattern: '^[0-9A-Fa-f]{6}$' },
              borderColor: { type: 'string', pattern: '^[0-9A-Fa-f]{6}$' },
              cellMargin: { type: 'number', minimum: 0, maximum: 36 },
              ...geometryProperties,
            },
            [
              'op',
              'slide_index',
              'name',
              'rows',
              'fontFace',
              'fontSize',
              'color',
              'left',
              'top',
              'width',
              'height',
            ],
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
    description:
      'Check bounded slides for negative, out-of-bounds, and overlapping geometry. Pass slide_index to return one full page when the whole-deck result exceeds the output budget.',
    inputSchema: {
      type: 'object',
      properties: slideProperties,
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: 'execute_office_js',
    description:
      'Execute a confirmation-gated bounded declarative PowerPoint program. Pass program directly as an object with version 1 and an operations array; do not stringify it and do not send JavaScript. Use snake_case fields. Native additions, modifications, and duplication require a durable paired-PC savepoint; use one page of pure additions with explicit styling, or validated SlideIR. Mixed modification/addition and multi-page addition programs are rejected before writing. Supported operations are set_shape_text, set_shape_geometry, add_text_box, add_geometric_shape, add_native_table (bounded string cells), delete_shape, and duplicate_slide (it must be the only operation).',
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
    name: 'add_slide_ir_objects',
    description:
      'Propose adding all text, shape and table objects plus a visible source footer from one validated SlideIR and Claim Ledger to an existing PowerPoint slide using native Office.js objects. If style.fontFallbacks is set, pass resolved_font_face from a checked PC compilation report or another explicitly reviewed choice; it must match style.fontFace or one of its declared candidates. The host still needs visual font review. Source truth is not verified here. Unsupported image/chart pages are rejected before writing. Requires a paired PC original-page savepoint and durable per-object receipts. Interrupted writes require explicit inspection and recovery; no automatic replay. Original-page restoration uses the separately confirmed package restore flow. This does not create a slide; visual and professional QA remain required.',
    inputSchema: {
      type: 'object',
      properties: {
        slide_index: { type: 'integer', minimum: 0, maximum: 31 },
        slide: PRESENTATION_DECK_SCHEMA.properties!.slides.items!,
        style: PRESENTATION_DECK_SCHEMA.properties!.style,
        claims: PRESENTATION_DECK_SCHEMA.properties!.claims,
        resolved_font_face: { type: 'string', minLength: 1, maxLength: 80 },
        explanation: { type: 'string', maxLength: 100 },
      },
      required: ['slide_index', 'slide', 'style', 'claims'],
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
    name: 'update_slide_chart_values',
    description:
      'Propose a confirmed update of numeric values for one native chart whose embedded Sheet1 XLSX and cache already match. Preserves categories and chart structure; rechecks and reads back both package parts.',
    inputSchema: {
      type: 'object',
      properties: {
        ...slideProperties,
        shape_id: { type: 'string', pattern: '^[1-9][0-9]{0,9}$' },
        values: {
          type: 'array',
          minItems: 1,
          maxItems: 8,
          items: {
            type: 'array',
            minItems: 1,
            maxItems: 32,
            items: { type: 'string', maxLength: 32 },
          },
        },
        explanation: { type: 'string', maxLength: 100 },
      },
      required: ['slide_index', 'shape_id', 'values'],
      additionalProperties: false,
    },
  },
  ...(['inspect', 'resume', 'undo', 'release', 'reapply'] as const).map((action) => ({
    name: `${action}_slide_chart_values_change`,
    description:
      action === 'inspect'
        ? 'Inspect a durable chart value change against the current host package without writing.'
        : action === 'resume'
          ? 'Finalize a known interrupted chart value write after classifying the host package; never replay an unknown host write.'
          : action === 'reapply'
            ? 'Propose reapplication of an undone chart value change using retained numeric values and a readable original backup. Creates an independent savepoint linked to the old change and rejects modified restored pages.'
            : action === 'release'
              ? 'Propose release of the PC backup only for a cancelled or undone chart change.'
              : 'Propose confirmed restoration of the original backed-up chart package when the exact applied package is still current.',
    inputSchema: {
      type: 'object',
      properties: { change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' } },
      required: ['change_id'],
      additionalProperties: false,
    },
  })),
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
    description:
      'Propose duplicating a slide immediately after its source with a durable paired-PC savepoint and recoverable receipt.',
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

function boundedJson(value: unknown): string {
  const result = JSON.stringify(value)
  if (new TextEncoder().encode(result).byteLength > MAX_POWERPOINT_RESULT_BYTES)
    throw new Error('office_read_failed')
  return result
}
function errorCode(error: unknown, write = false): string {
  const code = error instanceof Error ? error.message : ''
  if (
    [
      'invalid_tool_input',
      'office_api_unsupported',
      'office_concurrent_change',
      'presentation_existing_persistence_unavailable',
      'presentation_package_backup_invalid',
      'presentation_package_backup_capacity',
      'presentation_package_conflict',
      'presentation_package_invalid',
      'presentation_package_missing',
      'presentation_package_persistence_unavailable',
      'presentation_master_xml_persistence_unavailable',
      'presentation_package_qa_stale',
      'presentation_package_stale',
      'presentation_package_state_invalid',
      'presentation_package_unknown',
      'presentation_master_backup_invalid',
      'presentation_master_backup_capacity',
      'presentation_native_master_inverse_unproven',
      'presentation_native_master_package_unproven',
      'presentation_native_master_conflict',
      'presentation_native_master_state_invalid',
      'presentation_native_master_stale',
      'presentation_native_master_write_uncertain',
      'presentation_page_backup_cleanup_failed',
      'presentation_native_add_conflict',
      'presentation_native_add_pending',
      'presentation_native_add_state_invalid',
      'presentation_native_add_backup_release_failed',
      'presentation_existing_batch_missing',
      'presentation_existing_batch_stale',
      'presentation_existing_batch_state_invalid',
      'presentation_chart_backup_invalid',
      'presentation_existing_backup_capacity',
      'presentation_document_changed',
      'presentation_baseline_changed',
      'cancelled',
    ].includes(code)
  )
    return code
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
    'shape',
    'fill',
    'lineColor',
    'rows',
    'fontFace',
    'fontSize',
    'color',
    'borderColor',
    'cellMargin',
    'bold',
    'align',
    'margin',
    'verticalAlignment',
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
          ![
            'op',
            'slide_index',
            'name',
            'text',
            'left',
            'top',
            'width',
            'height',
            'fontFace',
            'fontSize',
            'color',
            'bold',
            'align',
            'margin',
            'verticalAlignment',
          ].includes(key),
      ) ||
      typeof operation.name !== 'string' ||
      !operation.name ||
      operation.name.length > 256 ||
      typeof operation.text !== 'string' ||
      operation.text.length > 12_000 ||
      (operation.fontFace !== undefined &&
        (typeof operation.fontFace !== 'string' ||
          !operation.fontFace ||
          operation.fontFace.length > 128)) ||
      (operation.fontSize !== undefined &&
        (typeof operation.fontSize !== 'number' ||
          !Number.isFinite(operation.fontSize) ||
          operation.fontSize < 6 ||
          operation.fontSize > 96)) ||
      (operation.color !== undefined &&
        (typeof operation.color !== 'string' || !/^[0-9A-Fa-f]{6}$/.test(operation.color))) ||
      (operation.bold !== undefined && typeof operation.bold !== 'boolean') ||
      (operation.align !== undefined &&
        !['left', 'center', 'right'].includes(String(operation.align))) ||
      (operation.margin !== undefined &&
        (typeof operation.margin !== 'number' ||
          !Number.isFinite(operation.margin) ||
          operation.margin < 0 ||
          operation.margin > 72)) ||
      (operation.verticalAlignment !== undefined &&
        !['top', 'middle', 'bottom'].includes(String(operation.verticalAlignment)))
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
      ...(operation.fontFace !== undefined ? { fontFace: operation.fontFace as string } : {}),
      ...(operation.fontSize !== undefined ? { fontSize: operation.fontSize as number } : {}),
      ...(operation.color !== undefined ? { color: operation.color as string } : {}),
      ...(operation.bold !== undefined ? { bold: operation.bold as boolean } : {}),
      ...(operation.align !== undefined
        ? { align: operation.align as 'left' | 'center' | 'right' }
        : {}),
      ...(operation.margin !== undefined ? { margin: operation.margin as number } : {}),
      ...(operation.verticalAlignment !== undefined
        ? { verticalAlignment: operation.verticalAlignment as 'top' | 'middle' | 'bottom' }
        : {}),
    }
  }
  if (operation.op === 'add_geometric_shape') {
    if (
      Object.keys(operation).some(
        (key) =>
          ![
            'op',
            'slide_index',
            'name',
            'shape',
            'fill',
            'lineColor',
            'left',
            'top',
            'width',
            'height',
          ].includes(key),
      ) ||
      typeof operation.name !== 'string' ||
      !operation.name ||
      operation.name.length > 256 ||
      !['rect', 'ellipse', 'roundRect'].includes(String(operation.shape)) ||
      typeof operation.fill !== 'string' ||
      !/^[0-9A-Fa-f]{6}$/.test(operation.fill) ||
      typeof operation.lineColor !== 'string' ||
      !/^[0-9A-Fa-f]{6}$/.test(operation.lineColor)
    )
      throw new Error('invalid_tool_input')
    finiteGeometry()
    return {
      op: 'add_geometric_shape',
      slide_index: operation.slide_index as number,
      name: operation.name,
      shape: operation.shape as 'rect' | 'ellipse' | 'roundRect',
      fill: operation.fill,
      lineColor: operation.lineColor,
      left: operation.left as number,
      top: operation.top as number,
      width: operation.width as number,
      height: operation.height as number,
    }
  }
  if (operation.op === 'add_native_table') {
    const rows = operation.rows
    if (
      Object.keys(operation).some(
        (key) =>
          ![
            'op',
            'slide_index',
            'name',
            'rows',
            'fontFace',
            'fontSize',
            'color',
            'borderColor',
            'cellMargin',
            'left',
            'top',
            'width',
            'height',
          ].includes(key),
      ) ||
      typeof operation.name !== 'string' ||
      !operation.name ||
      operation.name.length > 256 ||
      !Array.isArray(rows) ||
      rows.length < 1 ||
      rows.length > 20 ||
      !Array.isArray(rows[0]) ||
      rows[0].length < 1 ||
      rows[0].length > 12 ||
      rows.length * rows[0].length > 128 ||
      rows.some(
        (row) =>
          !Array.isArray(row) ||
          row.length !== rows[0].length ||
          row.some((cell) => typeof cell !== 'string' || cell.length > 256),
      ) ||
      JSON.stringify(rows).length > 12_000 ||
      typeof operation.fontFace !== 'string' ||
      !operation.fontFace ||
      operation.fontFace.length > 128 ||
      typeof operation.fontSize !== 'number' ||
      !Number.isFinite(operation.fontSize) ||
      operation.fontSize < 6 ||
      operation.fontSize > 48 ||
      typeof operation.color !== 'string' ||
      !/^[0-9A-Fa-f]{6}$/.test(operation.color) ||
      (operation.borderColor !== undefined &&
        (typeof operation.borderColor !== 'string' ||
          !/^[0-9A-Fa-f]{6}$/.test(operation.borderColor))) ||
      (operation.cellMargin !== undefined &&
        (typeof operation.cellMargin !== 'number' ||
          !Number.isFinite(operation.cellMargin) ||
          operation.cellMargin < 0 ||
          operation.cellMargin > 36))
    )
      throw new Error('invalid_tool_input')
    finiteGeometry()
    return {
      op: 'add_native_table',
      slide_index: operation.slide_index as number,
      name: operation.name,
      rows: rows as string[][],
      fontFace: operation.fontFace,
      fontSize: operation.fontSize,
      color: operation.color,
      ...(operation.borderColor !== undefined
        ? { borderColor: operation.borderColor as string }
        : {}),
      ...(operation.cellMargin !== undefined ? { cellMargin: operation.cellMargin as number } : {}),
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
  screenshotFallback?(
    slideIndex: number,
    signal?: AbortSignal,
    expectedSlideId?: string,
  ): Promise<{ slideId: string; base64: string; mime: 'image/png'; renderer?: 'libreoffice' }>
  nativeAddSavepoint?: {
    documentId(): Promise<string>
    request(body: unknown, signal?: AbortSignal): Promise<Response>
    readExistingBatch(id: string): PresentationExistingBatch | undefined
    readExistingPageChange?(id: string): PresentationExistingPageChange | undefined
    writeExistingBatch(
      record: PresentationExistingBatch,
      expected: PresentationExistingBatch | undefined,
    ): Promise<void>
  }
  durablePackage?(
    kind: 'slide' | 'chart',
    slideIndex: number,
    replacements: XmlReplacement[],
    explanation?: string,
    signal?: AbortSignal,
  ): Promise<unknown>
  durableMasterXml?(
    replacements: XmlReplacement[],
    explanation?: string,
    signal?: AbortSignal,
  ): Promise<unknown>
  durableMaster?(
    operations: PowerPointMasterOperation[],
    explanation?: string,
    signal?: AbortSignal,
  ): Promise<ReturnType<StructuredProposalController['propose']>>
  durableDuplicate?(
    slideIndex: number,
    explanation?: string,
    signal?: AbortSignal,
    toolName?: 'duplicate_slide' | 'execute_office_js',
  ): Promise<{ proposalId: string; changeId: string; status: string }>
  durableModify?(
    operations: import('./presentation-existing-batch.js').NativeModifyOperation[],
    explanation?: string,
    signal?: AbortSignal,
  ): Promise<{ proposalId: string; changeId: string; status: string }>
  durableTextEditAvailable?(): boolean
  durableTextEdit?(
    input: { slide_index: number; shape_id: string; text: string; explanation?: string },
    signal?: AbortSignal,
  ): Promise<ToolExecution>
  chartSavepoint?: {
    documentId(): Promise<string>
    request(body: unknown, signal?: AbortSignal): Promise<Response>
    readExistingChartChange(id: string): PresentationExistingChartChange | undefined
    writeExistingChartChange(
      record: PresentationExistingChartChange,
      expected: PresentationExistingChartChange | undefined,
    ): Promise<void>
  }
}): AgentSkill {
  const nativeExecution = options.nativeAddSavepoint
    ? createPresentationNativeAddExecution({
        ...options.nativeAddSavepoint,
        adapter: options.adapter,
      })
    : undefined
  const nativeRelease = options.nativeAddSavepoint
    ? createPresentationNativeAddRelease({
        ...options.nativeAddSavepoint,
        proposals: options.proposals,
      })
    : undefined
  const nativeReleaseTool: AgentToolDef = {
    name: 'release_slide_ir_addition',
    description:
      'Propose irreversible release of the original PC page backup only after a native addition is durably undone. Requires separate confirmation. Does not modify PowerPoint or certify QA.',
    inputSchema: {
      type: 'object',
      properties: { change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' } },
      required: ['change_id'],
      additionalProperties: false,
    },
  }
  const nativeRestoration =
    options.nativeAddSavepoint?.readExistingPageChange &&
    options.adapter.exportPresentationPagePackage
      ? createPresentationNativeAddRestoration({
          ...options.nativeAddSavepoint,
          readExistingPageChange: options.nativeAddSavepoint.readExistingPageChange,
          exportPresentationPagePackage: options.adapter.exportPresentationPagePackage.bind(
            options.adapter,
          ),
        })
      : undefined
  const nativeRestorationTool: AgentToolDef = {
    name: 'finalize_slide_ir_addition_restore',
    description:
      'Close a native-addition journal only after the separately confirmed original-page restore is durably applied. Verify the exact restoration source, both original backups, complete page order and live original package before recording the actual restored page ID. Never writes the host or certifies visual QA.',
    inputSchema: {
      type: 'object',
      properties: {
        change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
        restoration_change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
      },
      required: ['change_id', 'restoration_change_id'],
      additionalProperties: false,
    },
  }
  const nativeProposal =
    options.nativeAddSavepoint && nativeExecution
      ? createPresentationNativeAddProposal({
          ...options.nativeAddSavepoint,
          adapter: options.adapter,
          proposals: options.proposals,
          execution: nativeExecution,
        })
      : undefined
  const nativeTools: AgentToolDef[] = ['inspect_slide_ir_addition', 'resume_slide_ir_addition'].map(
    (name) => ({
      name,
      description: name.startsWith('inspect')
        ? 'Read-only inspect a durable native addition against the original page backup and actual SDK identities. Does not replay writes or certify visual QA.'
        : 'Propose explicit continuation of a verified native addition. Claim a proven lost receipt before further stepwise writes. Unproven in-flight writes remain pending; restore the original page through the controlled package workflow.',
      inputSchema: {
        type: 'object',
        properties: { change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' } },
        required: ['change_id'],
        additionalProperties: false,
      },
    }),
  )
  const masterXmlEditingSupported = options.platform?.toLowerCase() !== 'mac'
  const nativeMasterEditingSupported = options.nativeMasterEditingSupported !== false
  async function proposePackageEdit(
    _toolName: string,
    kind: PackageEditKind,
    slideIndex: number,
    replacements: XmlReplacement[],
    explanation: string | undefined,
    signal?: AbortSignal,
  ): Promise<ToolExecution> {
    let proposal: unknown
    if (kind === 'master') {
      if (!options.durableMasterXml)
        throw new Error('presentation_master_xml_persistence_unavailable')
      proposal = await options.durableMasterXml(structuredClone(replacements), explanation, signal)
    } else {
      if (!options.durablePackage) throw new Error('presentation_package_persistence_unavailable')
      proposal = await options.durablePackage(
        kind,
        slideIndex,
        structuredClone(replacements),
        explanation,
        signal,
      )
    }
    return {
      output: boundedJson(proposal),
      mutated: false,
      summary: `Proposed PowerPoint ${kind} XML edit`,
    }
  }

  async function proposeChartValues(
    slideIndex: number,
    shapeId: string,
    values: string[][],
    explanation: string | undefined,
    signal?: AbortSignal,
    reapply?: { source: PresentationExistingChartChange; check(s?: AbortSignal): Promise<void> },
  ): Promise<ToolExecution> {
    const durable = options.chartSavepoint
    if (!durable) throw new Error('office_api_unsupported')
    const documentId = await durable.documentId()
    await reapply?.check(signal)
    const deck = await options.adapter.verifySlides(signal)
    const before = await options.adapter.exportSlidePackage(slideIndex, signal)
    const beforeSlideIds = deck.slides.map((slide) => slide.slideId)
    if (!documentId || beforeSlideIds[slideIndex] !== before.slideId)
      throw new Error('office_state_uncertain')
    const beforeDigest = await presentationPackageDigest(before.base64, signal)
    const prepared = await updatePowerPointChartDataPackage(before.base64, shapeId, values, signal)
    const afterDigest = await presentationPackageDigest(prepared.base64, signal)
    if (
      reapply &&
      (before.slideId !== reapply.source.restoredSlideId ||
        beforeDigest !== reapply.source.beforePackageDigest ||
        afterDigest !== reapply.source.afterPackageDigest)
    )
      throw new Error('presentation_existing_chart_manual_review')
    await reapply?.check(signal)
    const edit = await captureChartValuePackageEdit(before.base64, prepared.base64, signal)
    let applied: typeof edit | undefined
    const changeId = crypto.randomUUID(),
      backupId = crypto.randomUUID()
    let record: PresentationExistingChartChange | undefined
    const store = async (next: PresentationExistingChartChange) => {
      if (!validatePresentationExistingChartChange(next))
        throw new Error('presentation_existing_chart_state_invalid')
      if ((await durable.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      await durable.writeExistingChartChange(next, record)
      if (JSON.stringify(durable.readExistingChartChange(changeId)) !== JSON.stringify(next))
        throw new Error('office_state_uncertain')
      record = structuredClone(next)
    }
    const unchanged = async (s?: AbortSignal) => {
      await reapply?.check(s)
      if ((await durable.documentId()) !== documentId) return false
      const host = await options.adapter.verifySlides(s)
      const current = await options.adapter.exportSlidePackage(slideIndex, s)
      return (
        JSON.stringify(host.slides.map((slide) => slide.slideId)) ===
          JSON.stringify(beforeSlideIds) &&
        current.slideId === before.slideId &&
        (await presentationPackageDigest(current.base64, s)) === beforeDigest &&
        (await verifyPowerPointPackageInputs(current.base64, edit.beforeHashes, s))
      )
    }
    const proposal = options.proposals.propose({
      operation: reapply ? 'reapply_slide_chart_values_change' : 'update_slide_chart_values',
      toolName: reapply ? 'reapply_slide_chart_values_change' : 'update_slide_chart_values',
      title: explanation || 'Update native chart values',
      preview: {
        changeId,
        ...(reapply ? { reapplies: reapply.source.changeId } : {}),
        slideIndex,
        shapeId,
        values,
        changedPaths: edit.changedPaths,
        beforeHashes: edit.beforeHashes,
        afterHashes: edit.afterHashes,
      },
      impact: {
        host: 'powerpoint',
        targets: [
          `slide:${deck.slides[slideIndex]?.slideId ?? before.slideId}`,
          `shape:${shapeId}`,
        ],
        count: 1,
      },
      fingerprint: before.fingerprint,
      before: { slideId: before.slideId, hashes: edit.beforeHashes },
      after: { hashes: edit.afterHashes },
      code: JSON.stringify({ version: 1, operation: 'update_chart_values', shapeId, values }),
      validate: async (confirmSignal) =>
        !durable.readExistingChartChange(changeId) && (await unchanged(confirmSignal)),
      execute: async (confirmSignal) => {
        if (!(await unchanged(confirmSignal)) || durable.readExistingChartChange(changeId))
          throw new Error('proposal_stale')
        const current = await options.adapter.exportSlidePackage(slideIndex, confirmSignal)
        const updated = await updatePowerPointChartDataPackage(
          current.base64,
          shapeId,
          values,
          confirmSignal,
        )
        applied = await captureChartValuePackageEdit(current.base64, updated.base64, confirmSignal)
        if ((await presentationPackageDigest(applied.base64, confirmSignal)) !== afterDigest)
          throw new Error('proposal_stale')
        const { sha256, sizeBytes } = await describePagePackageBackup(current.base64, confirmSignal)
        const backup = { backupId, sha256, sizeBytes }
        try {
          const saved = await saveChartPackageBackup(
            {
              request: durable.request,
              documentId,
              hostSlideId: before.slideId,
              slideIds: beforeSlideIds,
              base64: current.base64,
              backupId,
            },
            confirmSignal,
          )
          if (saved.sha256 !== backup.sha256 || saved.sizeBytes !== backup.sizeBytes)
            throw new Error('presentation_chart_backup_invalid')
          if (!(await unchanged(confirmSignal))) throw new Error('proposal_stale')
          await store({
            version: 1,
            changeId,
            documentId,
            oldSlideId: before.slideId,
            shapeId,
            slideIndex,
            beforeSlideIds,
            beforePackageDigest: beforeDigest,
            afterPackageDigest: afterDigest,
            backup,
            state: 'pending',
            values: structuredClone(values),
            ...(reapply ? { reapplies: reapply.source.changeId } : {}),
          })
        } catch (error) {
          // Only a confirmed absence of the first journal entry proves this backup is orphaned.
          try {
            if (
              (await durable.documentId()) === documentId &&
              !durable.readExistingChartChange(changeId)
            ) {
              await cleanupUncommittedChartPackageBackup({
                request: durable.request,
                documentId,
                hostSlideId: before.slideId,
                slideIds: beforeSlideIds,
                backup,
              })
            }
          } catch {
            /* Keep the original failure; an uncertain backup must remain recoverable. */
          }
          throw error
        }
        await store({ ...record!, state: 'write_pending' })
        if (!(await unchanged(confirmSignal))) throw new Error('proposal_stale')
        const inserted = await options.adapter.replaceSlidePackage(
          slideIndex,
          applied.base64,
          false,
          applied,
          confirmSignal,
          { slideId: before.slideId, packageDigest: beforeDigest },
        )
        await store({ ...record!, state: 'applied', newSlideId: inserted.slideId })
      },
      verify: async (confirmSignal) => {
        if (!applied || record?.state !== 'applied') throw new Error('office_verify_failed')
        const current = await options.adapter.exportSlidePackage(slideIndex, confirmSignal)
        if (
          current.slideId !== record.newSlideId ||
          !(await verifyImportedPowerPointPackage(current.base64, applied, confirmSignal)) ||
          (await presentationPackageDigest(current.base64, confirmSignal)) !== afterDigest
        )
          throw new Error('office_verify_failed')
      },
    })
    return {
      output: boundedJson({
        ...proposal,
        changeId,
        ...(reapply ? { reapplies: reapply.source.changeId } : {}),
      }),
      mutated: false,
      summary: 'Proposed native chart values update',
    }
  }

  async function chartChangeTool(
    name: string,
    changeId: string,
    signal?: AbortSignal,
  ): Promise<ToolExecution> {
    const durable = options.chartSavepoint
    if (!durable) throw new Error('office_api_unsupported')
    const saved = durable.readExistingChartChange(changeId)
    if (
      !saved ||
      !validatePresentationExistingChartChange(saved) ||
      (await durable.documentId()) !== saved.documentId
    )
      throw new Error('presentation_existing_chart_missing')
    let record = structuredClone(saved)
    const observe = async (s?: AbortSignal) => {
      if (
        (await durable.documentId()) !== record.documentId ||
        JSON.stringify(durable.readExistingChartChange(changeId)) !== JSON.stringify(record)
      )
        throw new Error('presentation_existing_chart_stale')
      const deck = await options.adapter.verifySlides(s)
      const ids = deck.slides.map((slide) => slide.slideId)
      if (
        ids.length !== record.beforeSlideIds.length ||
        ids.some((id, index) => index !== record.slideIndex && id !== record.beforeSlideIds[index])
      )
        return { status: 'conflict' as const, slideId: undefined, digest: undefined }
      const current = await options.adapter.exportSlidePackage(record.slideIndex, s)
      const digest = await presentationPackageDigest(current.base64, s)
      const status =
        digest === record.beforePackageDigest
          ? 'before'
          : digest === record.afterPackageDigest &&
              (record.newSlideId === undefined || current.slideId === record.newSlideId)
            ? 'after'
            : 'conflict'
      return { status, slideId: current.slideId, digest }
    }
    const store = async (next: PresentationExistingChartChange) => {
      if ((await durable.documentId()) !== record.documentId)
        throw new Error('presentation_document_changed')
      await durable.writeExistingChartChange(next, record)
      if (JSON.stringify(durable.readExistingChartChange(changeId)) !== JSON.stringify(next))
        throw new Error('office_state_uncertain')
      record = structuredClone(next)
    }
    if (name === 'reapply_slide_chart_values_change') {
      if (record.state !== 'undone' || !record.values || record.backupReleasedAt)
        throw new Error('presentation_existing_chart_state_invalid')
      const source = structuredClone(record)
      const check = async (s?: AbortSignal) => {
        const observed = await observe(s)
        if (observed.status !== 'before' || observed.slideId !== source.restoredSlideId)
          throw new Error('presentation_existing_chart_manual_review')
        await readChartPackageBackup(
          {
            request: durable.request,
            documentId: source.documentId,
            hostSlideId: source.oldSlideId,
            slideIds: source.beforeSlideIds,
            backup: source.backup,
            expectedPackageDigest: source.beforePackageDigest,
          },
          s,
        )
        if (
          (await durable.documentId()) !== source.documentId ||
          JSON.stringify(durable.readExistingChartChange(changeId)) !== JSON.stringify(source)
        )
          throw new Error('presentation_existing_chart_stale')
      }
      await check(signal)
      return proposeChartValues(
        source.slideIndex,
        source.shapeId,
        structuredClone(source.values!),
        'Reapply native chart values',
        signal,
        { source, check },
      )
    }
    if (name === 'release_slide_chart_values_change') {
      if (!['cancelled', 'undone'].includes(record.state) || record.backupReleasedAt)
        throw new Error('presentation_existing_chart_state_invalid')
      const initial = JSON.stringify(record)
      const currentRecord = async () =>
        (await durable.documentId()) === record.documentId &&
        JSON.stringify(durable.readExistingChartChange(changeId)) === initial
      const proposal = options.proposals.propose({
        operation: name,
        toolName: name,
        title: 'Release completed chart backup',
        preview: {
          changeId,
          state: record.state,
          backupId: record.backup.backupId,
          sizeBytes: record.backup.sizeBytes,
        },
        impact: { host: 'powerpoint', targets: [`backup:${record.backup.backupId}`], count: 1 },
        fingerprint: `${record.changeId}:${record.state}:${record.backup.sha256}`,
        before: { backupId: record.backup.backupId },
        after: { released: true },
        code: JSON.stringify({ changeId, operation: 'release_chart_backup' }),
        validate: async () => currentRecord(),
        execute: async (s) => {
          if (!(await currentRecord())) throw new Error('proposal_stale')
          const backup = record.backup
          const response = await durable.request(
            {
              operation: 'existing_page_backup_release',
              documentId: record.documentId,
              backupId: backup.backupId,
              hostSlideId: record.oldSlideId,
              slideIds: record.beforeSlideIds,
              sha256: backup.sha256,
              sizeBytes: backup.sizeBytes,
            },
            s,
          )
          const receipt = (await response.json()) as Record<string, unknown>
          if (
            !response.ok ||
            receipt.status !== 'released' ||
            receipt.documentId !== record.documentId ||
            receipt.backupId !== backup.backupId ||
            receipt.hostSlideId !== record.oldSlideId ||
            JSON.stringify(receipt.slideIds) !== JSON.stringify(record.beforeSlideIds) ||
            receipt.sha256 !== backup.sha256 ||
            receipt.sizeBytes !== backup.sizeBytes
          )
            throw new Error('presentation_chart_backup_release_failed')
          await store({ ...record, backupReleasedAt: new Date().toISOString() })
        },
        verify: async () => {
          if (!record.backupReleasedAt || !((await durable.documentId()) === record.documentId))
            throw new Error('office_verify_failed')
        },
      })
      return {
        output: boundedJson(proposal),
        mutated: false,
        summary: 'Proposed completed chart backup release',
      }
    }
    const observed = await observe(signal)
    if (name === 'inspect_slide_chart_values_change') {
      const verified =
        record.state === 'applied'
          ? observed.status === 'after'
          : record.state === 'undone' || record.state === 'cancelled'
            ? observed.status === 'before'
            : false
      return {
        output: boundedJson({
          changeId,
          state: record.state,
          hostStatus: observed.status,
          slideId: observed.slideId,
          currentHostVerified: verified,
          manualReview: !verified,
          qaPassed: false,
        }),
        mutated: false,
        summary: 'Inspected chart values change',
      }
    }
    if (name === 'resume_slide_chart_values_change') {
      const outcome =
        observed.status === 'before' && ['pending', 'write_pending'].includes(record.state)
          ? 'cancelled'
          : observed.status === 'after' && record.state === 'write_pending'
            ? 'applied'
            : observed.status === 'before' && record.state === 'undo_pending'
              ? 'undone'
              : undefined
      if (!outcome || !observed.slideId)
        throw new Error('presentation_existing_chart_manual_review')
      const proposal = options.proposals.propose({
        operation: name,
        toolName: name,
        title: 'Finalize interrupted chart change',
        preview: {
          changeId,
          previousState: record.state,
          observed: observed.status,
          nextState: outcome,
        },
        impact: { host: 'powerpoint', targets: [`slide:${observed.slideId}`], count: 1 },
        fingerprint: `${observed.slideId}:${observed.digest}`,
        before: { state: record.state },
        after: { state: outcome },
        code: JSON.stringify({ changeId, outcome }),
        validate: async (s) => {
          const next = await observe(s)
          return (
            next.status === observed.status &&
            next.slideId === observed.slideId &&
            next.digest === observed.digest
          )
        },
        execute: async (s) => {
          const next = await observe(s)
          if (
            next.status !== observed.status ||
            next.slideId !== observed.slideId ||
            next.digest !== observed.digest
          )
            throw new Error('proposal_stale')
          if (outcome !== 'cancelled')
            await readChartPackageBackup(
              {
                request: durable.request,
                documentId: record.documentId,
                hostSlideId: record.oldSlideId,
                slideIds: record.beforeSlideIds,
                backup: record.backup,
                expectedPackageDigest: record.beforePackageDigest,
              },
              s,
            )
          await store({
            ...record,
            state: outcome,
            ...(outcome === 'applied' ? { newSlideId: observed.slideId } : {}),
            ...(outcome === 'undone' ? { restoredSlideId: observed.slideId } : {}),
          })
        },
        verify: async (s) => {
          if ((await observe(s)).status !== observed.status || record.state !== outcome)
            throw new Error('office_verify_failed')
        },
      })
      return {
        output: boundedJson(proposal),
        mutated: false,
        summary: 'Proposed chart change recovery',
      }
    }
    if (
      record.state !== 'applied' ||
      observed.status !== 'after' ||
      observed.slideId !== record.newSlideId
    )
      throw new Error('presentation_existing_chart_manual_review')
    const original = await readChartPackageBackup(
      {
        request: durable.request,
        documentId: record.documentId,
        hostSlideId: record.oldSlideId,
        slideIds: record.beforeSlideIds,
        backup: record.backup,
        expectedPackageDigest: record.beforePackageDigest,
      },
      signal,
    )
    const proposal = options.proposals.propose({
      operation: name,
      toolName: name,
      title: 'Undo native chart values update',
      preview: {
        changeId,
        slideId: observed.slideId,
        shapeId: record.shapeId,
        from: record.afterPackageDigest,
        to: record.beforePackageDigest,
      },
      impact: { host: 'powerpoint', targets: [`slide:${observed.slideId}`], count: 1 },
      fingerprint: `${observed.slideId}:${observed.digest}`,
      before: { digest: record.afterPackageDigest },
      after: { digest: record.beforePackageDigest },
      code: JSON.stringify({ changeId, operation: 'undo_chart_values' }),
      validate: async (s) => {
        const next = await observe(s)
        return (
          next.status === 'after' &&
          next.slideId === observed.slideId &&
          next.digest === observed.digest
        )
      },
      execute: async (s) => {
        const next = await observe(s)
        if (
          next.status !== 'after' ||
          next.slideId !== observed.slideId ||
          next.digest !== observed.digest
        )
          throw new Error('proposal_stale')
        const backup = await readChartPackageBackup(
          {
            request: durable.request,
            documentId: record.documentId,
            hostSlideId: record.oldSlideId,
            slideIds: record.beforeSlideIds,
            backup: record.backup,
            expectedPackageDigest: record.beforePackageDigest,
          },
          s,
        )
        if (
          (await presentationPackageDigest(backup, s)) !==
          (await presentationPackageDigest(original, s))
        )
          throw new Error('presentation_chart_backup_invalid')
        const current = await options.adapter.exportSlidePackage(record.slideIndex, s)
        const reverse = await captureChartValuePackageEdit(current.base64, backup, s)
        await store({ ...record, state: 'undo_pending' })
        const restored = await options.adapter.replaceSlidePackage(
          record.slideIndex,
          backup,
          false,
          reverse,
          s,
          { slideId: record.newSlideId!, packageDigest: record.afterPackageDigest },
        )
        await store({ ...record, state: 'undone', restoredSlideId: restored.slideId })
      },
      verify: async (s) => {
        const next = await observe(s)
        if (
          record.state !== 'undone' ||
          next.status !== 'before' ||
          next.slideId !== record.restoredSlideId
        )
          throw new Error('office_verify_failed')
      },
    })
    return { output: boundedJson(proposal), mutated: false, summary: 'Proposed chart values undo' }
  }

  return {
    id: 'office-powerpoint',
    systemPrompt:
      'PowerPoint reads are bounded. Every write creates an explicit proposal and is semantically verified after confirmation. execute_office_js accepts only a versioned declarative JSON program; JavaScript and ambient browser authority are rejected. XML tools accept only allowlisted bounded package parts.' +
      ' Prefer inspect_slide_masters and native edit_slide_master for backgrounds, theme colors, and layout inheritance. PowerPoint for Mac must never use edit_slide_master_xml.',
    get tools() {
      return [
        ...tools.filter(
          (tool) =>
            (Boolean(options.chartSavepoint) ||
              ![
                'update_slide_chart_values',
                'inspect_slide_chart_values_change',
                'resume_slide_chart_values_change',
                'undo_slide_chart_values_change',
                'release_slide_chart_values_change',
                'reapply_slide_chart_values_change',
              ].includes(tool.name)) &&
            ((Boolean(options.durableTextEdit) && (options.durableTextEditAvailable?.() ?? true)) ||
              tool.name !== 'edit_slide_text') &&
            (masterXmlEditingSupported || tool.name !== 'edit_slide_master_xml') &&
            (nativeMasterEditingSupported ||
              !['inspect_slide_masters', 'edit_slide_master'].includes(tool.name)),
        ),
        ...(nativeExecution ? nativeTools : []),
        ...(nativeRestoration ? [nativeRestorationTool] : []),
        ...(nativeRelease ? [nativeReleaseTool] : []),
      ]
    },
    async executeTool(call, signal) {
      if (call.inputError || call.truncated)
        return failure(
          call.name,
          'invalid_tool_input',
          PROGRAM_TOOLS.has(call.name) ? invalidToolInput('program') : undefined,
        )
      try {
        assertNotCancelled(signal)
        if (call.name === 'release_slide_ir_addition') {
          if (!nativeRelease) throw new Error('office_api_unsupported')
          const input = exactRecord(call.input, ['change_id'])
          if (
            typeof input.change_id !== 'string' ||
            !/^[A-Za-z0-9_-]{1,128}$/.test(input.change_id)
          )
            throw new Error('invalid_tool_input')
          return {
            output: boundedJson(await nativeRelease.propose(input.change_id, signal)),
            mutated: false,
            summary: 'Proposed irreversible original-page backup release',
          }
        }
        if (call.name === 'finalize_slide_ir_addition_restore') {
          if (!nativeRestoration) throw new Error('office_api_unsupported')
          const input = exactRecord(call.input, ['change_id', 'restoration_change_id'])
          if (
            [input.change_id, input.restoration_change_id].some(
              (value) => typeof value !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(value),
            )
          )
            throw new Error('invalid_tool_input')
          const record = await nativeRestoration.finalize(
            input.change_id as string,
            input.restoration_change_id as string,
            signal,
          )
          return {
            output: boundedJson({
              changeId: record.changeId,
              state: record.state,
              restoredSlideId: record.restoredSlideId,
              createdShapeIds: record.createdShapeIds,
              historicalOnly: true,
              visualQaVerified: false,
            }),
            mutated: false,
            summary: 'Reconciled original-page restoration receipt',
          }
        }
        if (call.name === 'inspect_slide_ir_addition' || call.name === 'resume_slide_ir_addition') {
          if (!nativeExecution || !options.nativeAddSavepoint)
            throw new Error('office_api_unsupported')
          const input = exactRecord(call.input, ['change_id'])
          if (
            typeof input.change_id !== 'string' ||
            !/^[A-Za-z0-9_-]{1,128}$/.test(input.change_id)
          )
            throw new Error('invalid_tool_input')
          const changeId = input.change_id
          const saved = options.nativeAddSavepoint.readExistingBatch(changeId)
          if (!saved || saved.version !== 2 || saved.changeId !== changeId)
            throw new Error('presentation_existing_batch_missing')
          const record = structuredClone(saved)
          const observed = await nativeExecution.inspect(changeId, signal)
          if (call.name === 'inspect_slide_ir_addition')
            return {
              output: boundedJson({
                changeId,
                state: record.state,
                nextIndex: record.nextIndex,
                operationCount: record.operations.length,
                observation: observed.observation,
                createdShapeIds: observed.createdShapeIds,
                visualQaVerified: false,
              }),
              mutated: false,
              summary: 'Inspected durable native addition',
            }
          if (record.state !== 'applying') throw new Error('presentation_native_add_conflict')
          if (
            record.inFlightIndex !== undefined &&
            observed.observation.completedCount === record.nextIndex
          )
            throw new Error('presentation_native_add_pending')
          const stillCurrent = async (s?: AbortSignal) => {
            const proof = await nativeExecution.inspect(changeId, s)
            return (
              JSON.stringify(options.nativeAddSavepoint!.readExistingBatch(changeId)) ===
                JSON.stringify(record) &&
              proof.packageDigest === observed.packageDigest &&
              JSON.stringify(proof.createdShapeIds) === JSON.stringify(observed.createdShapeIds)
            )
          }
          const proposal = options.proposals.propose({
            operation: call.name,
            toolName: call.name,
            title: '继续已核对的原生对象添加',
            preview: {
              changeId,
              completed: observed.observation.completedCount,
              total: record.operations.length,
              receiptRecovery: record.inFlightIndex !== undefined,
              visualQaVerified: false,
            },
            impact: {
              host: 'powerpoint',
              targets: [record.hostSlideId],
              count: record.operations.length - record.nextIndex,
            },
            fingerprint: fingerprint(JSON.stringify(record) + observed.packageDigest),
            validate: stillCurrent,
            execute: async (s) => {
              if (!(await stillCurrent(s))) throw new Error('office_concurrent_change')
              for (let index = record.nextIndex; index < record.operations.length; index++)
                await nativeExecution.step(changeId, s)
            },
            verify: async (s) => {
              const proof = await nativeExecution.inspect(changeId, s)
              const latest = options.nativeAddSavepoint!.readExistingBatch(changeId)
              if (
                !latest ||
                latest.version !== 2 ||
                latest.state !== 'applied' ||
                proof.observation.status !== 'complete'
              )
                throw new Error('office_verify_failed')
            },
          })
          return {
            output: boundedJson(proposal),
            mutated: false,
            summary: 'Proposed durable native addition continuation',
          }
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
          let result: Awaited<ReturnType<PowerPointAdapter['screenshotSlide']>> & {
            renderer?: 'libreoffice'
          }
          try {
            result = await options.adapter.screenshotSlide(input.slide_index, signal)
          } catch (error) {
            const code =
              error && typeof error === 'object' ? (error as { code?: unknown }).code : undefined
            if (
              !options.screenshotFallback ||
              ![
                'office_screenshot_unavailable',
                'office_image_too_large',
                'ActivityLimitReached',
                'Timeout',
              ].includes(String(code))
            )
              throw error
            try {
              const expectedSlideId =
                error && typeof error === 'object'
                  ? (error as { targetSlideId?: unknown }).targetSlideId
                  : undefined
              result = await options.screenshotFallback(
                input.slide_index,
                signal,
                typeof expectedSlideId === 'string' ? expectedSlideId : undefined,
              )
              if (typeof expectedSlideId === 'string' && result.slideId !== expectedSlideId)
                throw new Error('office_concurrent_change', { cause: error })
            } catch (fallbackError) {
              if (
                signal?.aborted ||
                (fallbackError instanceof Error &&
                  [
                    'cancelled',
                    'office_concurrent_change',
                    'presentation_qa_stale',
                    'presentation_page_backup_cleanup_failed',
                    'presentation_native_add_conflict',
                    'presentation_native_add_pending',
                    'presentation_existing_batch_missing',
                    'presentation_existing_batch_stale',
                    'presentation_existing_batch_state_invalid',
                    'presentation_chart_backup_invalid',
                    'presentation_existing_backup_capacity',
                    'presentation_document_changed',
                    'presentation_baseline_changed',
                  ].includes(fallbackError.message))
              )
                throw fallbackError
              return {
                output: boundedJson({
                  status: 'waiting_screenshot',
                  slideIndex: input.slide_index,
                }),
                mutated: false,
                summary: '当前页等待宿主或备用渲染截图；稍后可重试',
              }
            }
          }
          assertNotCancelled(signal)
          if (
            result.mime !== 'image/png' ||
            typeof result.slideId !== 'string' ||
            !result.slideId ||
            result.slideId.length > 256 ||
            !validPng(result.base64)
          )
            throw new Error('office_read_failed')
          return {
            output: boundedJson({
              slideId: result.slideId,
              slideIndex: input.slide_index,
              ...(result.renderer ? { renderer: result.renderer } : {}),
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
          const input = verifyInput(call.input)
          const deck = await options.adapter.verifySlides(signal)
          if (input.slide_index !== undefined) {
            const slide = deck.slides.find((value) => value.slideIndex === input.slide_index)
            if (!slide) throw new Error('invalid_tool_input')
            return {
              output: boundedJson({
                slideWidth: deck.slideWidth,
                slideHeight: deck.slideHeight,
                slide,
                deckTruncated: deck.truncated ?? false,
              }),
              mutated: false,
              summary: 'Verified PowerPoint slide',
            }
          }
          return {
            output: boundedJson(deck),
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
          if (!options.durableTextEdit || options.durableTextEditAvailable?.() === false)
            throw new Error('presentation_existing_persistence_unavailable')
          return await options.durableTextEdit(input, signal)
        }
        if (call.name === 'duplicate_slide') {
          const input = slideInput(call.input)
          if (!options.durableDuplicate)
            throw new Error('presentation_existing_persistence_unavailable')
          return {
            output: boundedJson(
              await options.durableDuplicate(
                input.slide_index,
                input.explanation,
                signal,
                call.name,
              ),
            ),
            mutated: false,
            summary: 'Proposed durable PowerPoint slide duplication',
          }
        }
        if (call.name === 'execute_office_js' || call.name === 'add_slide_ir_objects') {
          let input: { code: string; explanation?: string }
          if (call.name === 'add_slide_ir_objects') {
            const value = exactRecord(call.input, [
              'slide_index',
              'slide',
              'style',
              'claims',
              'resolved_font_face',
              'explanation',
            ])
            if (
              !Number.isSafeInteger(value.slide_index) ||
              (value.slide_index as number) < 0 ||
              (value.slide_index as number) > 31 ||
              (value.explanation !== undefined &&
                (typeof value.explanation !== 'string' || value.explanation.length > 100))
            )
              throw new Error('invalid_tool_input')
            let serialized: string
            try {
              serialized = JSON.stringify([
                value.slide,
                value.style,
                value.claims,
                value.resolved_font_face,
              ])
            } catch {
              throw new Error('invalid_tool_input')
            }
            if (!serialized || new TextEncoder().encode(serialized).byteLength > MAX_CODE)
              throw new Error('invalid_tool_input')
            const deck = parsePresentationDeck({
              version: 1,
              id: 'office-ir',
              title: 'Office IR',
              style: value.style,
              assets: [],
              claims: value.claims,
              slides: [value.slide],
            })
            const operations = officeOperationsForSlideIR(
              deck.slides[0]!,
              deck.style,
              value.slide_index as number,
              deck.claims,
              value.resolved_font_face as string | undefined,
            )
            const code = JSON.stringify({ version: 1, operations })
            if (new TextEncoder().encode(code).byteLength > MAX_CODE)
              throw new Error('invalid_tool_input')
            input = {
              code,
              ...(typeof value.explanation === 'string' ? { explanation: value.explanation } : {}),
            }
          } else input = declarativeInput(call.input, { slide: false, explanationMax: 100 })
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
          const plannedNames =
            call.name === 'add_slide_ir_objects'
              ? new Set(
                  program.operations.flatMap((operation) =>
                    'name' in operation ? [operation.name] : [],
                  ),
                )
              : undefined
          if (plannedNames) {
            const current = await options.adapter.listSlideShapes(
              program.operations[0]!.slide_index,
              signal,
            )
            if (
              current.shapes.length + plannedNames.size > 1_000 ||
              current.shapes.some((shape) => plannedNames.has(shape.name))
            )
              throw new Error('office_concurrent_change')
          }
          const additions = program.operations.filter(
            (operation): operation is NativeAddOperation =>
              ['add_text_box', 'add_geometric_shape', 'add_native_table'].includes(operation.op),
          )
          if (
            additions.length &&
            (additions.length !== program.operations.length ||
              new Set(additions.map((operation) => operation.slide_index)).size !== 1)
          )
            throw new Error('office_api_unsupported')
          if (additions.length) {
            if (!nativeProposal) throw new Error('office_api_unsupported')
            const proposed = await nativeProposal.propose(
              additions,
              input.explanation,
              signal,
              call.name as 'add_slide_ir_objects' | 'execute_office_js',
            )
            return {
              output: boundedJson(proposed),
              mutated: false,
              summary: 'Proposed durable native SlideIR addition',
            }
          }
          if (!program.operations.some((operation) => operation.op === 'duplicate_slide')) {
            if (!options.durableModify)
              throw new Error('presentation_existing_persistence_unavailable')
            const proposed = await options.durableModify(
              structuredClone(
                program.operations,
              ) as import('./presentation-existing-batch.js').NativeModifyOperation[],
              input.explanation,
              signal,
            )
            return {
              output: boundedJson(proposed),
              mutated: false,
              summary: 'Proposed durable declarative PowerPoint modification',
            }
          }
          if (!options.durableDuplicate)
            throw new Error('presentation_existing_persistence_unavailable')
          return {
            output: boundedJson(
              await options.durableDuplicate(
                program.operations[0]!.slide_index,
                input.explanation,
                signal,
                'execute_office_js',
              ),
            ),
            mutated: false,
            summary: 'Proposed durable declarative PowerPoint duplication',
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
          if (!options.durableMaster)
            throw new Error('presentation_existing_persistence_unavailable')
          const proposal = await options.durableMaster(
            operations,
            input.explanation as string | undefined,
            signal,
          )
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
        if (call.name === 'update_slide_chart_values') {
          const input = exactRecord(call.input, [
            'slide_index',
            'shape_id',
            'values',
            'explanation',
          ])
          if (
            !Number.isInteger(input.slide_index) ||
            (input.slide_index as number) < 0 ||
            (input.slide_index as number) > MAX_SLIDE_INDEX ||
            typeof input.shape_id !== 'string' ||
            !/^[1-9]\d{0,9}$/.test(input.shape_id) ||
            !Array.isArray(input.values) ||
            input.values.length < 1 ||
            input.values.length > 8 ||
            input.values.some(
              (series) =>
                !Array.isArray(series) ||
                series.length < 1 ||
                series.length > 32 ||
                series.some(
                  (value) =>
                    typeof value !== 'string' ||
                    value.length > 32 ||
                    !/^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(value),
                ),
            ) ||
            (input.explanation !== undefined &&
              (typeof input.explanation !== 'string' || input.explanation.length > 100))
          )
            throw new Error('invalid_tool_input')
          return await proposeChartValues(
            input.slide_index as number,
            input.shape_id,
            input.values as string[][],
            input.explanation as string | undefined,
            signal,
          )
        }
        if (
          [
            'inspect_slide_chart_values_change',
            'resume_slide_chart_values_change',
            'undo_slide_chart_values_change',
            'release_slide_chart_values_change',
            'reapply_slide_chart_values_change',
          ].includes(call.name)
        ) {
          const input = exactRecord(call.input, ['change_id'])
          if (
            typeof input.change_id !== 'string' ||
            !/^[A-Za-z0-9_-]{1,128}$/.test(input.change_id)
          )
            throw new Error('invalid_tool_input')
          return await chartChangeTool(call.name, input.change_id, signal)
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
