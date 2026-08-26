import type { AgentSkill, AgentToolCall, ToolExecution } from '@wiswork/agent-core'
import type { OfficeDocumentClient } from '../office-document.js'
import {
  MAX_PROPOSAL_SELECTION_LENGTH,
  type ProposalController,
  type ProposalOperation,
} from './proposal-controller.js'

export const MAX_SELECTION_TEXT_LENGTH = MAX_PROPOSAL_SELECTION_LENGTH
export const MAX_PROPOSAL_TEXT_LENGTH = 12_000

const tools = [
  {
    name: 'read_selection',
    description: 'Read the current Office selection as text.',
    inputSchema: { type: 'object', properties: {}, required: [], additionalProperties: false },
  },
  ...(
    [
      ['propose_replace_selection', 'Prepare replacement text for explicit user confirmation.'],
      ['propose_append_text', 'Prepare text to append for explicit user confirmation.'],
    ] as const
  ).map(([name, description]) => ({
    name,
    description,
    inputSchema: {
      type: 'object',
      properties: { text: { type: 'string', minLength: 1, maxLength: MAX_PROPOSAL_TEXT_LENGTH } },
      required: ['text'],
      additionalProperties: false,
    },
  })),
]

function isRecord(input: unknown): input is Record<string, unknown> {
  return Boolean(input && typeof input === 'object' && !Array.isArray(input))
}

function isExactEmptyObject(input: unknown): boolean {
  return isRecord(input) && Object.keys(input).length === 0
}

function proposalText(input: unknown): string | undefined {
  if (!isRecord(input)) return undefined
  if (Object.keys(input).length !== 1 || !Object.hasOwn(input, 'text')) return undefined
  const text = input.text
  return typeof text === 'string' && text.length > 0 && text.length <= MAX_PROPOSAL_TEXT_LENGTH
    ? text
    : undefined
}

function invalid(name: string): ToolExecution {
  return { output: 'invalid_tool_input', isError: true, mutated: false, summary: name }
}

export function createOfficeSkill(
  document: OfficeDocumentClient,
  proposals: ProposalController,
): AgentSkill {
  return {
    id: 'office-selection',
    systemPrompt:
      'Use read_selection for document state. Replace and append tools only create previews; never claim a document edit occurred before the user confirms it.',
    tools,
    async executeTool(call: AgentToolCall) {
      if (call.inputError || call.truncated) return invalid(call.name)
      if (call.name === 'read_selection') {
        if (!isExactEmptyObject(call.input)) return invalid(call.name)
        try {
          const selection = await document.readSelection()
          return {
            output: selection.slice(0, MAX_SELECTION_TEXT_LENGTH),
            mutated: false,
            summary: 'Read selection',
          }
        } catch {
          return {
            output: 'office_read_failed',
            isError: true,
            mutated: false,
            summary: 'Read selection',
          }
        }
      }

      const operation: ProposalOperation | undefined =
        call.name === 'propose_replace_selection'
          ? 'replace'
          : call.name === 'propose_append_text'
            ? 'append'
            : undefined
      if (!operation) return invalid(call.name)
      const value = proposalText(call.input)
      if (value === undefined) return invalid(call.name)
      try {
        const proposal = await proposals.propose(operation, value)
        return {
          output: JSON.stringify({ proposalId: proposal.id, operation, mutated: false }),
          mutated: false,
          summary: operation === 'replace' ? 'Proposed replacement' : 'Proposed append',
        }
      } catch (error) {
        if (error instanceof Error && error.message === 'selection_too_large') {
          return {
            output: 'selection_too_large',
            isError: true,
            mutated: false,
            summary: call.name,
          }
        }
        return { output: 'office_read_failed', isError: true, mutated: false, summary: call.name }
      }
    },
  }
}
