import { describe, expect, it, vi } from 'vitest'
import {
  MAX_PROPOSAL_TEXT_LENGTH,
  MAX_SELECTION_TEXT_LENGTH,
  createOfficeSkill,
} from '../src/agent/office-skill.js'
import { createProposalController } from '../src/agent/proposal-controller.js'
import type { OfficeDocumentClient } from '../src/office-document.js'

function setup(selection = 'selected') {
  const document = {
    readSelection: vi.fn().mockResolvedValue(selection),
    replaceSelection: vi.fn(),
    appendText: vi.fn(),
  } as unknown as OfficeDocumentClient
  const proposals = createProposalController(document)
  return { document, proposals, skill: createOfficeSkill(document, proposals) }
}

describe('Office Agent skill', () => {
  it('automatically reads a capped selection with exact empty input', async () => {
    const { skill } = setup('x'.repeat(MAX_SELECTION_TEXT_LENGTH + 10))
    const result = await skill.executeTool({ id: '1', name: 'read_selection', input: {} })
    expect(result.isError).not.toBe(true)
    expect(result.output).toHaveLength(MAX_SELECTION_TEXT_LENGTH)
  })

  it.each([
    ['read_selection', { extra: true }],
    ['read_selection', null],
    ['propose_replace_selection', {}],
    ['propose_replace_selection', ['draft']],
    ['propose_replace_selection', { text: 'ok', extra: true }],
    ['propose_append_text', { text: '' }],
    ['propose_append_text', { text: 'x'.repeat(MAX_PROPOSAL_TEXT_LENGTH + 1) }],
  ])('rejects invalid exact arguments for %s', async (name, input) => {
    const { skill, document } = setup()
    const result = await skill.executeTool({
      id: '1',
      name,
      input: input as unknown as Record<string, unknown>,
    })
    expect(result.isError).toBe(true)
    expect(document.replaceSelection).not.toHaveBeenCalled()
    expect(document.appendText).not.toHaveBeenCalled()
  })

  it('rejects a proposal when its selection preview would exceed the explicit cap', async () => {
    const { skill, proposals } = setup('x'.repeat(MAX_SELECTION_TEXT_LENGTH + 1))
    const result = await skill.executeTool({
      id: '1',
      name: 'propose_replace_selection',
      input: { text: 'draft' },
    })
    expect(result).toMatchObject({ isError: true, output: 'selection_too_large' })
    expect(proposals.pending()).toBeUndefined()
  })

  it.each([
    ['propose_replace_selection', 'replace'],
    ['propose_append_text', 'append'],
  ] as const)('creates a %s proposal without mutating Office', async (name, operation) => {
    const { skill, proposals, document } = setup()
    const result = await skill.executeTool({ id: '1', name, input: { text: 'draft' } })
    expect(result.mutated).toBe(false)
    expect(proposals.pending()).toMatchObject({ operation, value: 'draft', before: 'selected' })
    expect(document.replaceSelection).not.toHaveBeenCalled()
    expect(document.appendText).not.toHaveBeenCalled()
  })
})
