export type LatexEditorCommand =
  | 'bold'
  | 'italic'
  | 'underline'
  | 'section'
  | 'subsection'
  | 'itemize'
  | 'enumerate'
  | 'inlineMath'
  | 'equation'
  | 'figure'
  | 'table'
  | 'cite'
  | 'ref'

export interface LatexEditorChange {
  from: number
  to: number
  insert: string
  selection: { anchor: number; head: number }
}

export function latexEditorChange(
  source: string,
  from: number,
  to: number,
  command: LatexEditorCommand,
): LatexEditorChange {
  const start = Math.max(0, Math.min(source.length, Math.min(from, to)))
  const end = Math.max(start, Math.min(source.length, Math.max(from, to)))
  const selected = source.slice(start, end)

  const change = (insert: string, editable: string, editableOffset = insert.indexOf(editable)) => ({
    from: start,
    to: end,
    insert,
    selection: {
      anchor: start + Math.max(0, editableOffset),
      head: start + Math.max(0, editableOffset) + editable.length,
    },
  })
  const wrap = (prefix: string, suffix: string, placeholder: string) => {
    const editable = selected || placeholder
    return change(`${prefix}${editable}${suffix}`, editable, prefix.length)
  }
  const environment = (name: 'itemize' | 'enumerate') => {
    const items = (selected || 'Item')
      .split(/\r?\n/)
      .map((line) => `\\item ${line || 'Item'}`)
      .join('\n')
    const prefix = `\\begin{${name}}\n`
    return change(`${prefix}${items}\n\\end{${name}}`, items, prefix.length)
  }

  switch (command) {
    case 'bold':
      return wrap('\\textbf{', '}', 'text')
    case 'italic':
      return wrap('\\textit{', '}', 'text')
    case 'underline':
      return wrap('\\underline{', '}', 'text')
    case 'section':
      return change('\\section{Title}\n', 'Title')
    case 'subsection':
      return change('\\subsection{Title}\n', 'Title')
    case 'itemize':
      return environment('itemize')
    case 'enumerate':
      return environment('enumerate')
    case 'inlineMath':
      return wrap('$', '$', 'formula')
    case 'equation': {
      const editable = selected || 'formula'
      return change(
        `\\begin{equation}\n${editable}\n\\end{equation}`,
        editable,
        '\\begin{equation}\n'.length,
      )
    }
    case 'figure': {
      const insert =
        '\\begin{figure}[htbp]\n' +
        '  \\centering\n' +
        '  \\includegraphics[width=\\linewidth]{image}\n' +
        '  \\caption{Caption}\n' +
        '  \\label{fig:label}\n' +
        '\\end{figure}'
      return change(insert, 'image')
    }
    case 'table': {
      const insert =
        '\\begin{table}[htbp]\n' +
        '  \\centering\n' +
        '  \\begin{tabular}{ll}\n' +
        '    A & B \\\\\n' +
        '  \\end{tabular}\n' +
        '  \\caption{Caption}\n' +
        '  \\label{tab:label}\n' +
        '\\end{table}'
      return change(insert, 'A & B')
    }
    case 'cite':
      return wrap('\\cite{', '}', 'key')
    case 'ref':
      return wrap('\\ref{', '}', 'label')
  }
}
