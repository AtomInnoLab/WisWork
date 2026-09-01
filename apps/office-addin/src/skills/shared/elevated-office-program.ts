import type { AgentSkill, ToolExecution } from '@wiswork/agent-core'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import { readUntilConverged } from './office-write-transaction.js'
import { XMLValidator } from 'fast-xml-parser'

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
  word: /^word\/(?:document|styles|numbering)\.xml$/,
  excel: /^xl\/(?:worksheets\/sheet[1-9]\d*|styles|sharedStrings)\.xml$/,
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
        !exact(args, ['slideIndex', 'text', 'left', 'top', 'width', 'height']) ||
        !text(args.text) ||
        ![args.left, args.top, args.width, args.height].every(
          (item) =>
            typeof item === 'number' && Number.isFinite(item) && Math.abs(item) <= 1_000_000,
        ) ||
        Number(args.width) <= 0 ||
        Number(args.height) <= 0
      )
        invalid()
    } else if (call === 'shape.delete') {
      if (!exact(args, ['slideIndex', 'shapeId']) || !OPAQUE.test(String(args.shapeId))) invalid()
    }
  }
}

function validateXml(xml: string): void {
  if (!xml || bytes(xml) > ELEVATED_OFFICE_LIMITS.xmlBytes || forbiddenText.test(xml)) invalid()
  if (
    XMLValidator.validate(xml) !== true ||
    /<!DOCTYPE|<!ENTITY|\bTargetMode\s*=\s*["']External["']|\b(?:Target|src|href)\s*=\s*["'](?:https?:|file:|ftp:)/i.test(
      xml,
    )
  )
    invalid()
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
      value.operations.length > ELEVATED_OFFICE_LIMITS.operations
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
    if (host === 'excel') invalid()
    if (
      !exact(value, ['version', 'kind', 'patches']) ||
      !Array.isArray(value.patches) ||
      value.patches.length < 1 ||
      value.patches.length > ELEVATED_OFFICE_LIMITS.targets
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
          'Propose one bounded closed Office.js AST or OOXML package patch for elevated confirmation.',
        inputSchema: {
          type: 'object',
          properties: { program: { type: 'object' } },
          required: ['program'],
          additionalProperties: false,
        },
      },
    ],
    async executeTool(call, signal) {
      if (call.name !== 'propose_raw_office_edit') return safeFailure('unknown_tool')
      if (options.automaticCorrection) return safeFailure('raw_office_confirmation_required')
      if (signal?.aborted) return safeFailure('cancelled')
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
      let applied = false
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
          let timeout: ReturnType<typeof setTimeout> | undefined
          try {
            await Promise.race([
              options.adapter.execute(program, snapshot, confirmationSignal),
              new Promise<never>((_resolve, reject) => {
                timeout = setTimeout(
                  () => reject(new Error('office_state_uncertain')),
                  ELEVATED_OFFICE_LIMITS.executionMs,
                )
              }),
            ])
          } finally {
            if (timeout) clearTimeout(timeout)
          }
          applied = true
        },
        verify: async (confirmationSignal) => {
          if (!applied) throw new Error('office_write_failed')
          if (!sameAuthority(initial, options.adapter.captureAuthority()))
            throw new Error('office_applied_unverified')
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
