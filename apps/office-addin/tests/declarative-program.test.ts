import { describe, expect, it } from 'vitest'
import { parseDeclarativeProgram } from '../src/skills/shared/declarative-program.js'

describe('declarative Office program', () => {
  const parseOperation = (value: unknown) => {
    if (!value || typeof value !== 'object' || Array.isArray(value))
      throw new Error('invalid_tool_input')
    const operation = value as Record<string, unknown>
    if (Object.keys(operation).some((key) => !['op', 'text'].includes(key)))
      throw new Error('invalid_tool_input')
    if (operation.op !== 'append_text' || typeof operation.text !== 'string')
      throw new Error('invalid_tool_input')
    return { op: 'append_text' as const, text: operation.text }
  }

  it('parses a bounded exact JSON program with a host operation allowlist', () => {
    expect(
      parseDeclarativeProgram(
        '{"version":1,"operations":[{"op":"append_text","text":"ok"}]}',
        parseOperation,
      ),
    ).toEqual({ version: 1, operations: [{ op: 'append_text', text: 'ok' }] })
  })

  it.each([
    'context.document.body.insertText("x")',
    '{"version":1,"operations":[],"fetch":"https://example.com"}',
    '{"version":2,"operations":[{"op":"append_text","text":"x"}]}',
    '{"version":1,"operations":[{"op":"eval","text":"x"}]}',
  ])('rejects code, unknown authority, versions, and operations', (source) => {
    expect(() => parseDeclarativeProgram(source, parseOperation)).toThrow('invalid_tool_input')
  })
})
