import type { AgentSkill, ToolExecution } from '@wiswork/agent-core'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import { readUntilConverged } from './office-write-transaction.js'
import { XMLParser, XMLValidator } from 'fast-xml-parser'

export type ElevatedOfficeHost = 'word' | 'excel' | 'powerpoint'
export const ELEVATED_OFFICE_LIMITS = Object.freeze({
  sourceBytes: 32 * 1024,
  operations: 32,
  astNodes: 2048,
  statements: 32,
  officeApiCalls: 32,
  targets: 32,
  xmlBytes: 256 * 1024,
  xmlNodes: 4096,
  xmlDepth: 32,
  outputBytes: 128 * 1024,
  executionMs: 15_000,
  reconciliationMs: 2_000,
})

export interface ElevatedOfficeAuthority {
  activeMode: 'enhanced' | 'standard'
  signedIn: boolean
  paired: boolean
  hostEnabled: boolean
  rawOfficeEnabled: boolean
  rawOfficeJsEnabled: boolean
  rawOfficeOoxmlEnabled: boolean
  documentId: string
  sessionId: string
  generation: number
  revision: string
}

export type ElevatedOfficeProgram =
  | Readonly<{
      version: 1
      kind: 'office_js_ast'
      operations: readonly Readonly<{ call: string; args: Readonly<Record<string, unknown>> }>[]
    }>
  | Readonly<{
      version: 1
      kind: 'ooxml_patch'
      patches: readonly Readonly<{ part: string; xml: string }>[]
    }>

export interface ElevatedOfficeSnapshot {
  id: string
  state?: unknown
}

export interface ElevatedOfficeAdapter {
  readonly host: ElevatedOfficeHost
  captureAuthority(): ElevatedOfficeAuthority
  snapshot(program: ElevatedOfficeProgram, signal?: AbortSignal): Promise<ElevatedOfficeSnapshot>
  validateSnapshot(
    program: ElevatedOfficeProgram,
    snapshot: ElevatedOfficeSnapshot,
    signal?: AbortSignal,
  ): Promise<boolean>
  execute(
    program: ElevatedOfficeProgram,
    snapshot: ElevatedOfficeSnapshot,
    signal?: AbortSignal,
    lifecycle?: Readonly<{ markStarted(): void; markApplied(): void }>,
  ): Promise<void>
  readback(
    program: ElevatedOfficeProgram,
    snapshot: ElevatedOfficeSnapshot,
    signal?: AbortSignal,
  ): Promise<{ verified: boolean; output?: unknown }>
  rollback(snapshot: ElevatedOfficeSnapshot, signal?: AbortSignal): Promise<void>
}

const JS_CALLS: Readonly<Record<ElevatedOfficeHost, ReadonlySet<string>>> = Object.freeze({
  word: new Set(['body.insertText', 'body.replaceAll']),
  excel: new Set(['range.setValues', 'range.clear']),
  powerpoint: new Set(['shape.setText', 'shape.setGeometry', 'slide.addTextBox', 'shape.delete']),
})
const PARTS: Readonly<Record<ElevatedOfficeHost, RegExp>> = Object.freeze({
  word: /$a/,
  excel: /$a/,
  powerpoint: /^ppt\/slides\/slide[1-9]\d*\.xml$/,
})
const OPAQUE = /^[A-Za-z0-9_.:-]{1,128}$/
const forbiddenText =
  /(?:https?:|fetch\s*\(|XMLHttpRequest|WebSocket|EventSource|navigator|clipboard|localStorage|sessionStorage|document\.cookie|credentials?|authorization|\beval\s*\(|\bFunction\s*\(|\bimport\s*\(|<script|javascript:|window\.|globalThis|Office\.context|setTimeout|setInterval|requestAnimationFrame|location\.|open\s*\()/i

function invalid(): never {
  throw new Error('raw_office_program_invalid')
}

function exact(value: Record<string, unknown>, keys: readonly string[]) {
  return Object.keys(value).length === keys.length && keys.every((key) => Object.hasOwn(value, key))
}

function bytes(value: string) {
  return new TextEncoder().encode(value).byteLength
}

function boundedData(value: unknown, depth = 0, budget = { nodes: 0 }): void {
  budget.nodes += 1
  if (budget.nodes > ELEVATED_OFFICE_LIMITS.astNodes) invalid()
  if (depth > 8) invalid()
  if (value === null || typeof value === 'boolean') return
  if (typeof value === 'number') {
    if (!Number.isFinite(value) || Math.abs(value) > 1_000_000_000) invalid()
    return
  }
  if (typeof value === 'string') {
    if (bytes(value) > 12_000 || forbiddenText.test(value)) invalid()
    return
  }
  if (Array.isArray(value)) {
    if (value.length > 256) invalid()
    value.forEach((item) => boundedData(item, depth + 1, budget))
    return
  }
  if (!value || typeof value !== 'object') invalid()
  const record = value as Record<string, unknown>
  if (Object.keys(record).length > 32) invalid()
  for (const [key, item] of Object.entries(record)) {
    if (!OPAQUE.test(key) || forbiddenText.test(key)) invalid()
    boundedData(item, depth + 1, budget)
  }
}

function integer(value: unknown, min: number, max: number): value is number {
  return Number.isSafeInteger(value) && Number(value) >= min && Number(value) <= max
}

function text(value: unknown, maximum = 12_000): value is string {
  return typeof value === 'string' && value.length <= maximum && !forbiddenText.test(value)
}

function formulaLikeExcelLiteral(value: string): boolean {
  for (const character of value) {
    const code = character.codePointAt(0) ?? 0
    if (
      code <= 0x20 ||
      (code >= 0x7f && code <= 0x9f) ||
      [
        0xa0, 0x1680, 0x200b, 0x200c, 0x200d, 0x2028, 0x2029, 0x202f, 0x205f, 0x2060, 0x3000,
        0xfeff,
      ].includes(code) ||
      (code >= 0x2000 && code <= 0x200a)
    )
      continue
    return ['=', '+', '-', '@'].includes(character)
  }
  return false
}

function validateOperation(
  host: ElevatedOfficeHost,
  call: string,
  args: Record<string, unknown>,
): void {
  const range = (value: unknown) =>
    typeof value === 'string' &&
    /^(?:'[^'\r\n]{1,64}'!)?\$?[A-Z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Z]{1,3}\$?[1-9]\d{0,6})?$/.test(
      value,
    )
  if (host === 'word') {
    if (call === 'body.insertText') {
      if (
        !exact(args, ['location', 'text']) ||
        !['start', 'end', 'replace'].includes(String(args.location)) ||
        !text(args.text)
      )
        invalid()
    } else if (call === 'body.replaceAll') {
      if (
        !exact(args, ['search', 'replacement']) ||
        !text(args.search, 1_000) ||
        !args.search ||
        !text(args.replacement)
      )
        invalid()
    }
  } else if (host === 'excel') {
    if (!integer(args.sheetId, 1, 1_000_000) || !range(args.range)) invalid()
    if (call === 'range.setValues') {
      if (
        !exact(args, ['sheetId', 'range', 'values']) ||
        !Array.isArray(args.values) ||
        args.values.length < 1 ||
        args.values.length > 256
      )
        invalid()
      boundedData(args.values)
      for (const row of args.values)
        if (
          !Array.isArray(row) ||
          row.length < 1 ||
          row.length > 256 ||
          row.some((cell) => typeof cell === 'string' && formulaLikeExcelLiteral(cell))
        )
          invalid()
    } else if (call === 'range.clear') {
      if (
        !exact(args, ['sheetId', 'range', 'clearType']) ||
        !['contents', 'formats', 'all'].includes(String(args.clearType))
      )
        invalid()
    }
  } else {
    if (!integer(args.slideIndex, 0, 100_000)) invalid()
    if (call === 'shape.setText') {
      if (
        !exact(args, ['slideIndex', 'shapeId', 'text']) ||
        !OPAQUE.test(String(args.shapeId)) ||
        !text(args.text)
      )
        invalid()
    } else if (call === 'shape.setGeometry') {
      if (
        !exact(args, ['slideIndex', 'shapeId', 'left', 'top', 'width', 'height']) ||
        !OPAQUE.test(String(args.shapeId)) ||
        ![args.left, args.top, args.width, args.height].every(
          (item) =>
            typeof item === 'number' && Number.isFinite(item) && Math.abs(item) <= 1_000_000,
        ) ||
        Number(args.width) <= 0 ||
        Number(args.height) <= 0
      )
        invalid()
    } else if (call === 'slide.addTextBox') {
      if (
        !exact(args, ['slideIndex', 'text', 'left', 'top', 'width', 'height', 'style']) ||
        !text(args.text) ||
        ![args.left, args.top, args.width, args.height].every(
          (item) =>
            typeof item === 'number' && Number.isFinite(item) && Math.abs(item) <= 1_000_000,
        ) ||
        Number(args.width) <= 0 ||
        Number(args.height) <= 0 ||
        !args.style ||
        typeof args.style !== 'object' ||
        Array.isArray(args.style) ||
        !exact(args.style as Record<string, unknown>, [
          'color',
          'fontFamily',
          'fontSize',
          'bold',
          'italic',
        ]) ||
        !/^#[0-9A-F]{6}$/i.test(String((args.style as Record<string, unknown>).color)) ||
        !text((args.style as Record<string, unknown>).fontFamily, 128) ||
        !integer((args.style as Record<string, unknown>).fontSize, 1, 400) ||
        typeof (args.style as Record<string, unknown>).bold !== 'boolean' ||
        typeof (args.style as Record<string, unknown>).italic !== 'boolean'
      )
        invalid()
    } else if (call === 'shape.delete') {
      if (!exact(args, ['slideIndex', 'shapeId']) || !OPAQUE.test(String(args.shapeId))) invalid()
    }
  }
}

const SLIDE_TAGS = new Set([
  'p:sld',
  'p:cSld',
  'p:spTree',
  'p:nvGrpSpPr',
  'p:cNvPr',
  'p:cNvGrpSpPr',
  'p:nvPr',
  'p:grpSpPr',
  'p:sp',
  'p:nvSpPr',
  'p:cNvSpPr',
  'p:spPr',
  'p:txBody',
  'a:xfrm',
  'a:off',
  'a:ext',
  'a:chOff',
  'a:chExt',
  'a:prstGeom',
  'a:avLst',
  'a:bodyPr',
  'a:lstStyle',
  'a:p',
  'a:pPr',
  'a:r',
  'a:rPr',
  'a:t',
  'a:endParaRPr',
  'a:solidFill',
  'a:srgbClr',
  'a:schemeClr',
  'a:alpha',
  'a:ln',
  'a:noFill',
])
const SLIDE_ATTRIBUTES = new Set([
  'id',
  'name',
  'x',
  'y',
  'cx',
  'cy',
  'prst',
  'val',
  'lang',
  'dirty',
  'marL',
  'indent',
  'algn',
  'anchor',
  'vert',
  'sz',
  'b',
  'i',
  'typeface',
  'rot',
  'useBgFill',
])
const slideParser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  processEntities: false,
  parseTagValue: false,
  parseAttributeValue: false,
})

function validateSlideTree(value: unknown): void {
  if (Array.isArray(value)) {
    value.forEach(validateSlideTree)
    return
  }
  if (!value || typeof value !== 'object') return
  for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
    if (key === '#text') continue
    if (key === ':@') {
      for (const [rawName, rawValue] of Object.entries(child as Record<string, unknown>)) {
        const name = rawName.replace(/^@_/, '')
        if (name.startsWith('xmlns:')) {
          if (
            typeof rawValue !== 'string' ||
            !/^(?:p|a|http:\/\/schemas\.(?:openxmlformats\.org|microsoft\.com)\/)/.test(rawValue)
          )
            invalid()
          continue
        }
        if (name.startsWith('r:') || !SLIDE_ATTRIBUTES.has(name) || typeof rawValue !== 'string')
          invalid()
        if (/(?:https?:|file:|ftp:|javascript:|ppaction:|mailto:)/i.test(rawValue)) invalid()
      }
      continue
    }
    if (!SLIDE_TAGS.has(key)) invalid()
    validateSlideTree(child)
  }
}

function validateXml(xml: string): void {
  if (!xml || bytes(xml) > ELEVATED_OFFICE_LIMITS.xmlBytes) invalid()
  if (
    XMLValidator.validate(xml) !== true ||
    /<!DOCTYPE|<!ENTITY|\b(?:hlinkClick|hlinkMouseOver|fld|instrText)\b|\baction\s*=|\br:(?:id|embed|link)\s*=|\bTargetMode\s*=\s*["']External["']|\b(?:Target|src|href)\s*=\s*["'](?:https?:|file:|ftp:)/i.test(
      xml,
    )
  )
    invalid()
  try {
    validateSlideTree(slideParser.parse(xml))
  } catch {
    invalid()
  }
  const tags = xml.match(/<\/?[A-Za-z_][^>]*>/g) ?? []
  if (tags.length > ELEVATED_OFFICE_LIMITS.xmlNodes) invalid()
  let depth = 0
  for (const tag of tags) {
    if (/^<\//.test(tag)) depth -= 1
    else if (!/\/>$/.test(tag) && !/^<\?/.test(tag)) depth += 1
    if (depth < 0 || depth > ELEVATED_OFFICE_LIMITS.xmlDepth) invalid()
  }
  if (depth !== 0) invalid()
}

function frozen<T>(value: T): T {
  if (value && typeof value === 'object') {
    Object.freeze(value)
    for (const child of Object.values(value as Record<string, unknown>)) frozen(child)
  }
  return value
}

export function parseElevatedOfficeProgram(
  host: ElevatedOfficeHost,
  input: unknown,
): ElevatedOfficeProgram {
  if (!input || typeof input !== 'object' || Array.isArray(input)) invalid()
  let encoded: string
  try {
    encoded = JSON.stringify(input)
  } catch {
    invalid()
  }
  if (bytes(encoded!) > ELEVATED_OFFICE_LIMITS.sourceBytes) invalid()
  const value = input as Record<string, unknown>
  if (value.version !== 1) invalid()
  if (value.kind === 'office_js_ast') {
    if (
      !exact(value, ['version', 'kind', 'operations']) ||
      !Array.isArray(value.operations) ||
      value.operations.length < 1 ||
      value.operations.length > ELEVATED_OFFICE_LIMITS.operations ||
      (host === 'excel' && value.operations.length !== 1)
    )
      invalid()
    const operations = value.operations.map((candidate) => {
      if (!candidate || typeof candidate !== 'object' || Array.isArray(candidate)) invalid()
      const operation = candidate as Record<string, unknown>
      if (
        !exact(operation, ['call', 'args']) ||
        typeof operation.call !== 'string' ||
        !JS_CALLS[host].has(operation.call) ||
        !operation.args ||
        typeof operation.args !== 'object' ||
        Array.isArray(operation.args)
      )
        invalid()
      boundedData(operation.args)
      validateOperation(host, operation.call, operation.args as Record<string, unknown>)
      return { call: operation.call, args: operation.args as Record<string, unknown> }
    })
    return frozen({ version: 1, kind: 'office_js_ast', operations })
  }
  if (value.kind === 'ooxml_patch') {
    if (host !== 'powerpoint') invalid()
    if (
      !exact(value, ['version', 'kind', 'patches']) ||
      !Array.isArray(value.patches) ||
      value.patches.length < 1 ||
      value.patches.length !== 1
    )
      invalid()
    const seen = new Set<string>()
    const patches = value.patches.map((candidate) => {
      if (!candidate || typeof candidate !== 'object' || Array.isArray(candidate)) invalid()
      const patch = candidate as Record<string, unknown>
      if (
        !exact(patch, ['part', 'xml']) ||
        typeof patch.part !== 'string' ||
        !PARTS[host].test(patch.part) ||
        seen.has(patch.part) ||
        typeof patch.xml !== 'string'
      )
        invalid()
      seen.add(patch.part)
      validateXml(patch.xml)
      return { part: patch.part, xml: patch.xml }
    })
    return frozen({ version: 1, kind: 'ooxml_patch', patches })
  }
  invalid()
}

function permitted(value: ElevatedOfficeAuthority): boolean {
  return (
    value.activeMode === 'enhanced' &&
    value.signedIn &&
    value.paired &&
    value.hostEnabled &&
    value.rawOfficeEnabled &&
    /^[A-Za-z0-9_-]{16,128}$/.test(value.documentId) &&
    /^[A-Za-z0-9_-]{16,128}$/.test(value.sessionId) &&
    Number.isSafeInteger(value.generation) &&
    value.generation >= 0 &&
    /^[A-Za-z0-9_-]{16,128}$/.test(value.revision)
  )
}

function permitsProgram(value: ElevatedOfficeAuthority, program: ElevatedOfficeProgram): boolean {
  return (
    permitted(value) &&
    (program.kind === 'office_js_ast' ? value.rawOfficeJsEnabled : value.rawOfficeOoxmlEnabled)
  )
}

function sameAuthority(a: ElevatedOfficeAuthority, b: ElevatedOfficeAuthority): boolean {
  return (
    permitted(b) &&
    a.documentId === b.documentId &&
    a.sessionId === b.sessionId &&
    a.generation === b.generation &&
    a.revision === b.revision &&
    a.rawOfficeEnabled === b.rawOfficeEnabled &&
    a.rawOfficeJsEnabled === b.rawOfficeJsEnabled &&
    a.rawOfficeOoxmlEnabled === b.rawOfficeOoxmlEnabled
  )
}

function programTargets(host: ElevatedOfficeHost, program: ElevatedOfficeProgram): string[] {
  if (program.kind === 'ooxml_patch') return program.patches.map((patch) => patch.part)
  return program.operations.map(({ call, args }) => {
    if (host === 'word') return `body:${call}`
    if (host === 'excel') return `sheet:${String(args.sheetId)}!${String(args.range)}`
    return `slide:${String(args.slideIndex)}${args.shapeId ? `:shape:${String(args.shapeId)}` : ':new-shape'}`
  })
}

function safeFailure(code: string): ToolExecution {
  return {
    output: code,
    isError: true,
    mutated: false,
    stopToolBatch: true,
    summary: 'Raw Office edit denied',
  }
}

function elevatedProgramSchema(host: ElevatedOfficeHost): Record<string, unknown> {
  const ast = {
    type: 'object',
    properties: {
      version: { type: 'integer', enum: [1] },
      kind: { type: 'string', enum: ['office_js_ast'] },
      operations: {
        type: 'array',
        minItems: 1,
        maxItems: host === 'excel' ? 1 : ELEVATED_OFFICE_LIMITS.operations,
        items: {
          type: 'object',
          properties: {
            call: { type: 'string', enum: [...JS_CALLS[host]] },
            args: { type: 'object' },
          },
          required: ['call', 'args'],
          additionalProperties: false,
        },
      },
    },
    required: ['version', 'kind', 'operations'],
    additionalProperties: false,
  }
  if (host !== 'powerpoint') return ast
  return {
    anyOf: [
      ast,
      {
        type: 'object',
        properties: {
          version: { type: 'integer', enum: [1] },
          kind: { type: 'string', enum: ['ooxml_patch'] },
          patches: {
            type: 'array',
            minItems: 1,
            maxItems: 1,
            items: {
              type: 'object',
              properties: {
                part: { type: 'string', pattern: '^ppt/slides/slide[1-9]\\d*\\.xml$' },
                xml: { type: 'string', maxLength: ELEVATED_OFFICE_LIMITS.xmlBytes },
              },
              required: ['part', 'xml'],
              additionalProperties: false,
            },
          },
        },
        required: ['version', 'kind', 'patches'],
        additionalProperties: false,
      },
    ],
  }
}

export function createElevatedOfficeSkill(options: {
  host: ElevatedOfficeHost
  adapter: ElevatedOfficeAdapter
  proposals: StructuredProposalController
  automaticCorrection?: boolean
  confirmationTitle?: string
}): AgentSkill {
  if (options.adapter.host !== options.host) throw new Error('raw_office_adapter_invalid')
  return {
    id: `office-${options.host}-elevated`,
    systemPrompt:
      'Raw Office edits use propose_raw_office_edit. Every call is a separate elevated proposal requiring fresh user confirmation. Never invoke it from automatic correction.',
    tools: [
      {
        name: 'propose_raw_office_edit',
        description:
          options.host === 'powerpoint'
            ? 'Propose one bounded closed Office.js AST or one slide OOXML package patch for elevated confirmation.'
            : 'Propose one bounded closed Office.js AST for elevated confirmation. OOXML is unavailable in this host.',
        inputSchema: {
          type: 'object',
          properties: { program: elevatedProgramSchema(options.host) },
          required: ['program'],
          additionalProperties: false,
        },
      },
    ],
    async executeTool(call, signal) {
      if (call.name !== 'propose_raw_office_edit') return safeFailure('unknown_tool')
      if (options.automaticCorrection) return safeFailure('raw_office_confirmation_required')
      if (signal?.aborted) return safeFailure('cancelled')
      if (options.proposals.isQuarantined()) return safeFailure('office_state_uncertain')
      const initial = options.adapter.captureAuthority()
      if (!permitted(initial)) return safeFailure('raw_office_denied')
      let program: ElevatedOfficeProgram
      try {
        if (!exact(call.input, ['program'])) invalid()
        program = parseElevatedOfficeProgram(options.host, call.input.program)
      } catch {
        return safeFailure('raw_office_program_invalid')
      }
      if (!permitsProgram(initial, program)) return safeFailure('raw_office_denied')
      if (options.proposals.pending()) return safeFailure('proposal_confirmation_in_progress')
      const canonical = JSON.stringify(program)
      const digest = selectionFingerprint(canonical)
      const targets = [...new Set(programTargets(options.host, program))]
      let snapshot: ElevatedOfficeSnapshot
      let writeState: 'not_started' | 'started' | 'applied' = 'not_started'
      let executionPending = false
      try {
        snapshot = await options.adapter.snapshot(program, signal)
      } catch {
        return safeFailure('office_read_failed')
      }
      if (!/^[A-Za-z0-9_-]{16,128}$/.test(snapshot.id)) return safeFailure('office_read_failed')
      const proposal = options.proposals.propose({
        operation: 'raw_office_edit',
        toolName: 'propose_raw_office_edit',
        title: options.confirmationTitle ?? `Elevated ${options.host} Office edit`,
        preview: { kind: program.kind, operations: targets.length, digest },
        impact: { host: options.host, targets, count: targets.length },
        fingerprint: digest,
        code: canonical,
        validate: async (confirmationSignal) =>
          sameAuthority(initial, options.adapter.captureAuthority()) &&
          (await options.adapter.validateSnapshot(program, snapshot, confirmationSignal)),
        execute: async (confirmationSignal) => {
          if (
            !sameAuthority(initial, options.adapter.captureAuthority()) ||
            confirmationSignal?.aborted
          )
            throw new Error('proposal_stale')
          const lifecycle = Object.freeze({
            markStarted: () => {
              if (writeState === 'not_started') writeState = 'started'
            },
            markApplied: () => {
              writeState = 'applied'
            },
          })
          const execution = Promise.resolve().then(() =>
            options.adapter.execute(program, snapshot, confirmationSignal, lifecycle),
          )
          const settled = execution.then(
            () => {
              lifecycle.markApplied()
              return { kind: 'applied' as const }
            },
            (error: unknown) => ({ kind: 'error' as const, error }),
          )
          const waitForSettlement = async (milliseconds: number) => {
            let timer: ReturnType<typeof setTimeout> | undefined
            try {
              return await Promise.race([
                settled,
                new Promise<{ kind: 'timeout' }>((resolve) => {
                  timer = setTimeout(() => resolve({ kind: 'timeout' }), milliseconds)
                }),
              ])
            } finally {
              if (timer) clearTimeout(timer)
            }
          }
          const quarantineUntilReconciled = () => {
            const lease = options.proposals.quarantine({
              sessionId: initial.sessionId,
              generation: initial.generation,
            })
            void (async () => {
              try {
                await settled
                const current = options.adapter.captureAuthority()
                if (
                  current.documentId !== initial.documentId ||
                  current.sessionId !== initial.sessionId ||
                  current.generation !== initial.generation
                )
                  return
                const readback = await options.adapter.readback(program, snapshot)
                if (readback.verified) {
                  options.proposals.resolveQuarantine(lease, { stable: true })
                  return
                }
                await options.adapter.rollback(snapshot)
                const restored = await options.adapter.validateSnapshot(program, snapshot)
                options.proposals.resolveQuarantine(lease, { stable: restored })
              } catch {
                options.proposals.resolveQuarantine(lease, { stable: false })
              }
            })()
          }
          let outcome = await waitForSettlement(ELEVATED_OFFICE_LIMITS.executionMs)
          if (outcome.kind === 'timeout') {
            outcome = await waitForSettlement(ELEVATED_OFFICE_LIMITS.reconciliationMs)
          }
          if (outcome.kind === 'timeout') {
            quarantineUntilReconciled()
            executionPending = true
            return
          }
          if (outcome.kind === 'error') {
            if (writeState === 'not_started') throw outcome.error
            quarantineUntilReconciled()
            throw new Error('office_applied_unverified')
          }
        },
        verify: async (confirmationSignal) => {
          if (writeState === 'not_started') throw new Error('office_write_failed')
          if (!sameAuthority(initial, options.adapter.captureAuthority()))
            throw new Error('office_applied_unverified')
          if (executionPending) {
            try {
              await options.adapter.readback(program, snapshot)
            } catch {
              /* authoritative read was attempted; a pending dispatch remains uncertain */
            }
            throw new Error('office_write_pending')
          }
          let result: { verified: boolean; output?: unknown }
          try {
            result = await readUntilConverged({
              read: () => options.adapter.readback(program, snapshot, confirmationSignal),
              accept: (value) => value.verified,
              signal: confirmationSignal,
            })
          } catch {
            throw new Error('office_applied_unverified')
          }
          const serialized = (() => {
            try {
              return JSON.stringify(result.output ?? null)
            } catch {
              throw new Error('office_applied_unverified')
            }
          })()
          if (bytes(serialized) > ELEVATED_OFFICE_LIMITS.outputBytes)
            throw new Error('office_applied_unverified')
          if (!result.verified) {
            try {
              await options.adapter.rollback(snapshot)
              if (await options.adapter.validateSnapshot(program, snapshot))
                throw new Error('office_write_failed')
            } catch (error) {
              if (error instanceof Error && error.message === 'office_write_failed') throw error
            }
            throw new Error('office_applied_unverified')
          }
        },
      })
      return {
        output: JSON.stringify({
          proposalId: proposal.id,
          status: 'awaiting_user_confirmation',
          mutated: false,
          digest,
        }),
        mutated: false,
        stopToolBatch: true,
        summary: 'Awaiting elevated confirmation',
      }
    },
  }
}
