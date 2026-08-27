import { describe, expect, it } from 'vitest'
import { latexEditorChange } from '../src/renderer/editor/editor-commands.js'

describe('LaTeX editor toolbar commands', () => {
  it('wraps a selection and keeps the selected text active', () => {
    expect(latexEditorChange('alpha beta', 0, 5, 'bold')).toEqual({
      from: 0,
      to: 5,
      insert: '\\textbf{alpha}',
      selection: { anchor: 8, head: 13 },
    })
  })

  it('inserts editable placeholders for inline and structural commands', () => {
    expect(latexEditorChange('', 0, 0, 'inlineMath')).toEqual({
      from: 0,
      to: 0,
      insert: '$formula$',
      selection: { anchor: 1, head: 8 },
    })
    expect(latexEditorChange('', 0, 0, 'section')).toEqual({
      from: 0,
      to: 0,
      insert: '\\section{Title}\n',
      selection: { anchor: 9, head: 14 },
    })
  })

  it('turns selected lines into list items', () => {
    expect(latexEditorChange('First\nSecond', 0, 12, 'itemize')).toEqual({
      from: 0,
      to: 12,
      insert: '\\begin{itemize}\n\\item First\n\\item Second\n\\end{itemize}',
      selection: { anchor: 16, head: 40 },
    })
  })

  it('provides complete figure, table, citation, and reference snippets', () => {
    expect(latexEditorChange('', 0, 0, 'figure').insert).toContain('\\includegraphics')
    expect(latexEditorChange('', 0, 0, 'table').insert).toContain('\\begin{tabular}')
    expect(latexEditorChange('', 0, 0, 'cite').insert).toBe('\\cite{key}')
    expect(latexEditorChange('', 0, 0, 'ref').insert).toBe('\\ref{label}')
  })

  it('preserves selected content in structural and insert commands', () => {
    expect(latexEditorChange('Introduction', 0, 12, 'section').insert).toBe(
      '\\section{Introduction}\n',
    )
    expect(latexEditorChange('plot.pdf', 0, 8, 'figure').insert).toContain(
      '\\includegraphics[width=\\linewidth]{plot.pdf}',
    )
    expect(latexEditorChange('Left & Right', 0, 12, 'table').insert).toContain('Left & Right \\\\')

    for (const [selection, command] of [
      ['figure', 'figure'],
      ['table', 'table'],
    ] as const) {
      const change = latexEditorChange(selection, 0, selection.length, command)
      expect(change.insert.slice(change.selection.anchor, change.selection.head)).toBe(selection)
    }
  })
})
