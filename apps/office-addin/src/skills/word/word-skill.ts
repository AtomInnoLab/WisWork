import type { AgentSkill, ToolExecution } from '@wiswork/agent-core'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { exactObject, integerField, optionalField, stringField } from '../../agent/tool-schema.js'
import type { InMemoryVfs } from '../shared/vfs.js'
import { parseDeclarativeProgram } from '../shared/declarative-program.js'
import type { WordAdapter, WordDeclarativeOperation } from './browser-word-adapter.js'
import { MAX_WORD_RESULT_BYTES } from './browser-word-adapter.js'
import { parseWordMarkdown, parseWordPlainText } from './word-markdown.js'

const MAX_INDEX = 1_000_000
const MAX_CODE = 32 * 1024
const MAX_SCREENSHOT_BYTES = 4 * 1024 * 1024
const emptyInput = exactObject({})
const textInput = exactObject({
  startParagraph: optionalField(integerField({ min: 0, max: MAX_INDEX })),
  endParagraph: optionalField(integerField({ min: 0, max: MAX_INDEX })),
  includeFormatting: optionalField((value) => {
    if (typeof value !== 'boolean') throw new Error('invalid_tool_input')
    return value
  }),
})
const ooxmlInput = exactObject({
  startChild: optionalField(integerField({ min: 0, max: MAX_INDEX })),
  endChild: optionalField(integerField({ min: 0, max: MAX_INDEX })),
})
const screenshotInput = exactObject({
  page: optionalField(integerField({ min: 1, max: 100_000 })),
  explanation: optionalField(stringField({ maxLength: 50 })),
})
const codeInput = exactObject({
  code: stringField({ minLength: 1, maxLength: MAX_CODE }),
  explanation: optionalField(stringField({ maxLength: 100 })),
})
const writeInput = exactObject({
  text: stringField({ minLength: 1, maxLength: 12_000 }),
  mode: (value) => {
    if (!['replace', 'append', 'prepend'].includes(String(value)))
      throw new Error('invalid_tool_input')
    return value as 'replace' | 'append' | 'prepend'
  },
  explanation: optionalField(stringField({ maxLength: 100 })),
  format: optionalField((value) => {
    if (!['markdown', 'plain_text'].includes(String(value))) throw new Error('invalid_tool_input')
    return value as 'markdown' | 'plain_text'
  }),
})

const tools = [
  {
    name: 'get_document_text',
    description: 'Read bounded Word paragraphs with indices, styles, alignment, and list metadata.',
    inputSchema: {
      type: 'object',
      properties: {
        startParagraph: { type: 'integer', minimum: 0, maximum: MAX_INDEX },
        endParagraph: { type: 'integer', minimum: 0, maximum: MAX_INDEX },
        includeFormatting: { type: 'boolean' },
      },
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: 'get_document_structure',
    description:
      'Read the bounded heading, table, content-control, section, and paragraph structure.',
    inputSchema: { type: 'object', properties: {}, required: [], additionalProperties: false },
  },
  {
    name: 'get_ooxml',
    description: 'Extract bounded body OOXML into the shared VFS and return structural mappings.',
    inputSchema: {
      type: 'object',
      properties: {
        startChild: { type: 'integer', minimum: 0, maximum: MAX_INDEX },
        endChild: { type: 'integer', minimum: 0, maximum: MAX_INDEX },
      },
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: 'screenshot_document',
    description: 'Render one Word page when the host and audited browser renderer support it.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', minimum: 1, maximum: 100_000 },
        explanation: { type: 'string', maxLength: 50 },
      },
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: 'write_document',
    description:
      'Write bounded Markdown as native Word headings, paragraphs, tables, and inline formatting after explicit confirmation. Lists fail closed. Set format=plain_text to disable Markdown. Use this for normal writing instead of execute_office_js. A returned awaiting_user_confirmation status is success; do not call another write tool in the same turn.',
    inputSchema: {
      type: 'object',
      properties: {
        text: { type: 'string', minLength: 1, maxLength: 12_000 },
        mode: { type: 'string', enum: ['replace', 'append', 'prepend'] },
        explanation: { type: 'string', maxLength: 100 },
        format: { type: 'string', enum: ['markdown', 'plain_text'] },
      },
      required: ['text', 'mode'],
      additionalProperties: false,
    },
  },
  {
    name: 'execute_office_js',
    description:
      'Propose a version-1 JSON declarative Word operation program for explicit confirmation. JavaScript is rejected.',
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
] as const

function failure(name: string, code: string): ToolExecution {
  return { output: code, isError: true, mutated: false, summary: name }
}

function errorCode(error: unknown): string {
  const code = error instanceof Error ? error.message : ''
  return [
    'invalid_tool_input',
    'office_api_unsupported',
    'office_read_failed',
    'vfs_limit',
    'vfs_path_denied',
    'cancelled',
  ].includes(code)
    ? code
    : 'office_read_failed'
}

function assertNotCancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new Error('cancelled')
}

function boundedJson(value: unknown): string {
  const output = JSON.stringify(value)
  if (new TextEncoder().encode(output).byteLength > MAX_WORD_RESULT_BYTES)
    throw new Error('office_read_failed')
  return output
}

function validPngBase64(value: unknown): value is string {
  if (typeof value !== 'string' || value.length === 0) return false
  if (value.length > Math.ceil(MAX_SCREENSHOT_BYTES / 3) * 4) return false
  if (value.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(value)) return false
  if (!value.startsWith('iVBORw0KGgo')) return false
  const padding = value.endsWith('==') ? 2 : value.endsWith('=') ? 1 : 0
  return (value.length / 4) * 3 - padding <= MAX_SCREENSHOT_BYTES
}

function parseWordOperation(value: unknown): WordDeclarativeOperation {
  if (!value || typeof value !== 'object' || Array.isArray(value)) invalidProposal()
  const input = value as Record<string, unknown>
  if (input.op === 'insert_text') {
    if (
      Object.keys(input).some((key) => !['op', 'location', 'text'].includes(key)) ||
      !['start', 'end', 'replace'].includes(String(input.location)) ||
      typeof input.text !== 'string' ||
      input.text.length < 1 ||
      input.text.length > 12_000
    )
      invalidProposal()
    return {
      op: 'insert_text',
      location: input.location as 'start' | 'end' | 'replace',
      text: input.text,
    }
  }
  if (input.op === 'replace_all') {
    if (
      Object.keys(input).some(
        (key) => !['op', 'search', 'replacement', 'matchCase'].includes(key),
      ) ||
      typeof input.search !== 'string' ||
      input.search.length < 1 ||
      input.search.length > 1_000 ||
      typeof input.replacement !== 'string' ||
      input.replacement.length > 12_000 ||
      (input.matchCase !== undefined && typeof input.matchCase !== 'boolean')
    )
      invalidProposal()
    return {
      op: 'replace_all',
      search: input.search,
      replacement: input.replacement,
      matchCase: input.matchCase === true,
    }
  }
  invalidProposal()
}

function invalidProposal(): never {
  throw new Error('invalid_tool_input')
}

const MAX_WORD_COMPARISON_BYTES = 6_000
const COMPARISON_ELLIPSIS = '…\n'

function documentPreviewText(value: string, operation: WordDeclarativeOperation): string {
  const encoder = new TextEncoder()
  if (encoder.encode(value).byteLength <= MAX_WORD_COMPARISON_BYTES) return value
  const budget = MAX_WORD_COMPARISON_BYTES - encoder.encode(COMPARISON_ELLIPSIS).byteLength
  const characters = Array.from(value)
  const fromEnd = operation.op === 'insert_text' && operation.location === 'end'
  const selected: string[] = []
  let bytes = 0
  for (
    let index = fromEnd ? characters.length - 1 : 0;
    fromEnd ? index >= 0 : index < characters.length;
    index += fromEnd ? -1 : 1
  ) {
    const character = characters[index]
    const size = encoder.encode(character).byteLength
    if (bytes + size > budget) break
    if (fromEnd) selected.unshift(character)
    else selected.push(character)
    bytes += size
  }
  return fromEnd ? `${COMPARISON_ELLIPSIS}${selected.join('')}` : `${selected.join('')}\n…`
}

function previewInsert(before: string, operation: WordDeclarativeOperation): string {
  if (operation.op !== 'insert_text') return before
  if (operation.location === 'start') return `${operation.text}${before}`
  if (operation.location === 'end') return `${before}${operation.text}`
  return operation.text
}

function previewDocumentWrite(
  before: string,
  inserted: string,
  mode: 'replace' | 'append' | 'prepend',
): string {
  if (mode === 'replace') return inserted
  if (!before) return inserted
  if (!inserted) return before
  return mode === 'append' ? `${before}\n${inserted}` : `${inserted}\n${before}`
}

export function createWordSkill(options: {
  adapter: WordAdapter
  vfs: InMemoryVfs
  proposals: StructuredProposalController
}): AgentSkill {
  const awaitingConfirmation = (): ToolExecution | undefined => {
    const pending = options.proposals.pending()
    return pending
      ? {
          output: JSON.stringify({
            proposalId: pending.id,
            status: 'awaiting_user_confirmation',
            mutated: false,
          }),
          mutated: false,
          summary: 'Awaiting confirmation',
        }
      : undefined
  }

  const proposeOperations = async (
    toolName: string,
    title: string,
    operations: WordDeclarativeOperation[],
    signal?: AbortSignal,
    code?: string,
  ): Promise<ToolExecution> => {
    const existing = awaitingConfirmation()
    if (existing) return existing
    const isDirectWrite = toolName === 'write_document'
    const documentSnapshot = isDirectWrite
      ? await options.adapter.getDocumentSnapshot(signal)
      : undefined
    assertNotCancelled(signal)
    const before = isDirectWrite
      ? documentPreviewText(documentSnapshot!.text, operations[0])
      : undefined
    const fingerprint = documentSnapshot?.fingerprint ?? (await options.adapter.fingerprint(signal))
    assertNotCancelled(signal)
    const proposal = options.proposals.propose({
      operation: toolName,
      toolName,
      title,
      preview:
        toolName === 'write_document'
          ? {
              mode: {
                replace: 'replace',
                end: 'append',
                start: 'prepend',
              }[operations[0]?.op === 'insert_text' ? operations[0].location : 'replace'],
            }
          : { version: 1, operations },
      impact: {
        host: 'word',
        targets: operations.map((operation) =>
          operation.op === 'insert_text'
            ? `document:${operation.location}`
            : `document:search:${operation.search.slice(0, 64)}`,
        ),
        count: operations.length,
      },
      fingerprint,
      before,
      after: before === undefined ? undefined : previewInsert(before, operations[0]),
      code,
      validate: async (confirmSignal) =>
        (await options.adapter.fingerprint(confirmSignal)) === fingerprint,
      execute: (confirmSignal) => options.adapter.executeOperations(operations, confirmSignal),
      verify: async (confirmSignal) => {
        if (!(await options.adapter.verifyOperations(operations, confirmSignal)))
          throw new Error('office_verify_failed')
      },
    })
    return {
      output: JSON.stringify({
        proposalId: proposal.id,
        status: 'awaiting_user_confirmation',
        mutated: false,
      }),
      mutated: false,
      summary: 'Awaiting confirmation',
    }
  }

  return {
    id: 'office-word',
    systemPrompt:
      'Word reads and screenshots are bounded. For ordinary drafting and document writing, call write_document with Markdown text and mode; it atomically creates native Word headings, tables, and inline formatting. Lists fail closed. Use format=plain_text only when literal Markdown is requested. execute_office_js accepts only a version-1 JSON declarative program with allowlisted insert_text or replace_all operations; JavaScript and ambient authority are rejected. Every write requires confirmation, stale-state validation, and semantic verification. A tool result with status awaiting_user_confirmation means the write proposal succeeded: stop calling write tools and ask the user to confirm the visible proposal.',
    tools: [...tools],
    async executeTool(call, signal) {
      if (call.inputError || call.truncated) return failure(call.name, 'invalid_tool_input')
      try {
        assertNotCancelled(signal)
        if (call.name === 'get_document_text') {
          const input = textInput(call.input)
          if (input.endParagraph !== undefined && input.endParagraph < (input.startParagraph ?? 0))
            throw new Error('invalid_tool_input')
          const result = await options.adapter.getDocumentText(input, signal)
          return { output: boundedJson(result), mutated: false, summary: 'Read Word text' }
        }
        if (call.name === 'get_document_structure') {
          emptyInput(call.input)
          const result = await options.adapter.getDocumentStructure(signal)
          return { output: boundedJson(result), mutated: false, summary: 'Read Word structure' }
        }
        if (call.name === 'get_ooxml') {
          const input = ooxmlInput(call.input)
          if (input.endChild !== undefined && input.endChild < (input.startChild ?? 0))
            throw new Error('invalid_tool_input')
          const result = await options.adapter.getOoxml(input, signal)
          assertNotCancelled(signal)
          const range =
            input.startChild !== undefined || input.endChild !== undefined
              ? `-${input.startChild ?? 0}-${input.endChild ?? 'end'}`
              : ''
          const file = `/home/user/ooxml/body${range}.xml`
          const output = boundedJson({
            file,
            bytes: new TextEncoder().encode(result.xml).byteLength,
            lines: result.xml.split('\n').length,
            children: result.children,
          })
          assertNotCancelled(signal)
          options.vfs.writeFile(file, result.xml)
          return {
            output,
            mutated: false,
            summary: 'Extracted Word OOXML',
          }
        }
        if (call.name === 'screenshot_document') {
          const input = screenshotInput(call.input)
          const result = await options.adapter.screenshotDocument(input.page ?? 1, signal)
          assertNotCancelled(signal)
          if (result.mime !== 'image/png' || !validPngBase64(result.base64))
            throw new Error('office_read_failed')
          return {
            output: JSON.stringify({ mime: result.mime }),
            modelContent: [{ type: 'image', image: { mime: result.mime, base64: result.base64 } }],
            display: {
              kind: 'images',
              items: [{ url: `data:${result.mime};base64,${result.base64}` }],
            },
            mutated: false,
            summary: 'Rendered Word page',
          }
        }
        if (call.name === 'write_document') {
          const input = writeInput(call.input)
          const existing = awaitingConfirmation()
          if (existing) return existing
          const parsed =
            input.format === 'plain_text'
              ? parseWordPlainText(input.text)
              : parseWordMarkdown(input.text)
          if (parsed.structure.lists) return failure(call.name, 'office_api_unsupported')
          const write = { ...parsed, mode: input.mode }
          const snapshot = await options.adapter.getDocumentSnapshot(signal)
          assertNotCancelled(signal)
          const previewOperation: WordDeclarativeOperation = {
            op: 'insert_text',
            location:
              input.mode === 'append' ? 'end' : input.mode === 'prepend' ? 'start' : 'replace',
            text: input.text,
          }
          const before = documentPreviewText(snapshot.text, previewOperation)
          const proposal = options.proposals.propose({
            operation: call.name,
            toolName: call.name,
            title: input.explanation || 'Write drafted content to the document',
            preview: {
              mode: input.mode,
              format: input.format ?? 'markdown',
              headings: parsed.structure.headings,
              lists: parsed.structure.lists,
              tables: parsed.structure.tables,
            },
            impact: { host: 'word', targets: ['document'], count: 1 },
            fingerprint: snapshot.fingerprint,
            before,
            after: previewDocumentWrite(before, parsed.semanticText, input.mode),
            validate: async (confirmSignal) =>
              (await options.adapter.fingerprint(confirmSignal)) === snapshot.fingerprint,
            execute: (confirmSignal) => options.adapter.executeDocumentWrite(write, confirmSignal),
            verify: async (confirmSignal) => {
              if (!(await options.adapter.verifyDocumentWrite(write, confirmSignal)))
                throw new Error('office_verify_failed')
            },
          })
          return {
            output: JSON.stringify({
              proposalId: proposal.id,
              status: 'awaiting_user_confirmation',
              mutated: false,
            }),
            mutated: false,
            summary: 'Awaiting confirmation',
          }
        }
        if (call.name === 'execute_office_js') {
          const input = codeInput(call.input)
          const program = parseDeclarativeProgram(input.code, parseWordOperation)
          return await proposeOperations(
            call.name,
            input.explanation || 'Execute declarative Word operations',
            program.operations,
            signal,
            input.code,
          )
        }
        return failure(call.name, 'invalid_tool_input')
      } catch (error) {
        return failure(call.name, errorCode(error))
      }
    },
  }
}
