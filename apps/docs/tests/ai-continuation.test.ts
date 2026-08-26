import { describe, expect, it } from 'vitest'
import { DOCS_CONTINUE_INSTRUCTION } from '../src/renderer/ai/continuation'

describe('Docs AI continuation', () => {
  it('instructs continuation to preserve completed work', () => {
    expect(DOCS_CONTINUE_INSTRUCTION).toContain('existing conversation and document state')
    expect(DOCS_CONTINUE_INSTRUCTION).toContain('only the outstanding work')
    expect(DOCS_CONTINUE_INSTRUCTION).toContain('Do not repeat edits')
  })
})
