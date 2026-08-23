import { describe, expect, it } from 'vitest'
import { parseWordMarkdown } from '../src/skills/word/word-markdown.js'

describe('bounded Word Markdown', () => {
  it('converts headings, lists, tables, and inline emphasis into native Word blocks', () => {
    const document = parseWordMarkdown(`# LLM overview

An **LLM** predicts *tokens*.

1. Train
2. Evaluate

| Model | Status |
| --- | --- |
| WisWork | Ready |`)

    expect(document.blocks).toEqual([
      { type: 'heading', level: 1, spans: [{ text: 'LLM overview' }] },
      {
        type: 'paragraph',
        spans: [
          { text: 'An ' },
          { text: 'LLM', bold: true },
          { text: ' predicts ' },
          { text: 'tokens', italic: true },
          { text: '.' },
        ],
      },
      { type: 'list', ordered: true, items: [[{ text: 'Train' }], [{ text: 'Evaluate' }]] },
      {
        type: 'table',
        rows: [
          ['Model', 'Status'],
          ['WisWork', 'Ready'],
        ],
        headerRows: 1,
      },
    ])
    expect(document.semanticText).toBe(
      'LLM overview\nAn LLM predicts tokens.\nTrain\nEvaluate\nModel\nStatus\nWisWork\nReady',
    )
    expect(document.structure).toEqual({ headings: 1, lists: 1, tables: 1 })
  })

  it('escapes HTML and leaves unsupported Markdown literal', () => {
    const document = parseWordMarkdown(
      '<script>**alert(1)**</script>\n[**link**](https://example.com)',
    )

    expect(document.blocks).toEqual([
      { type: 'paragraph', spans: [{ text: '<script>**alert(1)**</script>' }] },
      { type: 'paragraph', spans: [{ text: '[**link**](https://example.com)' }] },
    ])
    expect(document.semanticText).toBe(
      '<script>**alert(1)**</script>\n[**link**](https://example.com)',
    )
  })

  it('keeps emphasis markers literal inside inline code', () => {
    expect(parseWordMarkdown('Use `**raw**` now').blocks).toEqual([
      {
        type: 'paragraph',
        spans: [{ text: 'Use ' }, { text: '**raw**', code: true }, { text: ' now' }],
      },
    ])
  })

  it('rejects documents that exceed structural Office.js operation budgets', () => {
    expect(() =>
      parseWordMarkdown(Array.from({ length: 501 }, () => 'paragraph').join('\n')),
    ).toThrow('invalid_tool_input')
    expect(() =>
      parseWordMarkdown(
        `| ${Array.from({ length: 21 }, (_, index) => `c${index}`).join(' | ')} |\n| ${Array.from(
          { length: 21 },
          () => '---',
        ).join(' | ')} |`,
      ),
    ).toThrow('invalid_tool_input')
    expect(() => parseWordMarkdown(Array.from({ length: 1_001 }, () => '*a*').join(' '))).toThrow(
      'invalid_tool_input',
    )
    expect(() => parseWordMarkdown('| **formatted** | value |\n| --- | --- |')).toThrow(
      'invalid_tool_input',
    )
  })
})
