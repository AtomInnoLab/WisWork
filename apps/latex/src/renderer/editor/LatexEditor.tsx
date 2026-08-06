import { useEffect, useRef } from 'react'
import { autocompletion, completionKeymap } from '@codemirror/autocomplete'
import { defaultKeymap, history, historyKeymap, indentWithTab } from '@codemirror/commands'
import { syntaxHighlighting, defaultHighlightStyle } from '@codemirror/language'
import { setDiagnostics, lintGutter, type Diagnostic } from '@codemirror/lint'
import { searchKeymap } from '@codemirror/search'
import { EditorState } from '@codemirror/state'
import {
  EditorView,
  drawSelection,
  dropCursor,
  highlightActiveLine,
  highlightActiveLineGutter,
  highlightSpecialChars,
  keymap,
  lineNumbers,
} from '@codemirror/view'
import type { EditorDiagnostic } from '../compile/diagnostics.js'
import { latexLanguageExtensions } from './latex-language.js'

export interface LatexEditorProps {
  path: string
  value: string
  diagnostics: readonly EditorDiagnostic[]
  onChange: (value: string) => void
  onSave: () => void
  onCompile: () => void
  onCursorLine?: (line: number) => void
  revealLine?: number | null
}

function editorDiagnostics(
  view: EditorView,
  diagnostics: readonly EditorDiagnostic[],
): Diagnostic[] {
  return diagnostics.map((diagnostic) => {
    const line = view.state.doc.line(
      Math.min(view.state.doc.lines, Math.max(1, diagnostic.lineIndex + 1)),
    )
    const from = Math.min(line.to, line.from + diagnostic.columnIndex)
    return {
      from,
      to: Math.min(line.to, from + 1),
      severity: diagnostic.severity,
      message: diagnostic.message,
    }
  })
}

export function LatexEditor({
  path,
  value,
  diagnostics,
  onChange,
  onSave,
  onCompile,
  onCursorLine,
  revealLine,
}: LatexEditorProps) {
  const hostRef = useRef<HTMLDivElement>(null)
  const viewRef = useRef<EditorView | null>(null)
  const callbacks = useRef({ onChange, onSave, onCompile, onCursorLine })
  const currentInput = useRef({ value, diagnostics })
  callbacks.current = { onChange, onSave, onCompile, onCursorLine }
  currentInput.current = { value, diagnostics }

  useEffect(() => {
    const host = hostRef.current
    if (!host) return
    const view = new EditorView({
      parent: host,
      state: EditorState.create({
        doc: currentInput.current.value,
        extensions: [
          lineNumbers(),
          highlightActiveLineGutter(),
          highlightSpecialChars(),
          history(),
          drawSelection(),
          dropCursor(),
          highlightActiveLine(),
          syntaxHighlighting(defaultHighlightStyle, { fallback: true }),
          lintGutter(),
          autocompletion(),
          ...latexLanguageExtensions(path),
          keymap.of([
            {
              key: 'Mod-s',
              preventDefault: true,
              run: () => {
                callbacks.current.onSave()
                return true
              },
            },
            {
              key: 'Mod-Enter',
              preventDefault: true,
              run: () => {
                callbacks.current.onCompile()
                return true
              },
            },
            indentWithTab,
            ...defaultKeymap,
            ...historyKeymap,
            ...searchKeymap,
            ...completionKeymap,
          ]),
          EditorView.lineWrapping,
          EditorView.updateListener.of((update) => {
            if (update.docChanged) callbacks.current.onChange(update.state.doc.toString())
            if (update.selectionSet) {
              callbacks.current.onCursorLine?.(
                update.state.doc.lineAt(update.state.selection.main.head).number,
              )
            }
          }),
        ],
      }),
    })
    viewRef.current = view
    view.dispatch(
      setDiagnostics(view.state, editorDiagnostics(view, currentInput.current.diagnostics)),
    )
    return () => {
      view.destroy()
      viewRef.current = null
    }
    // Recreate for a different file so history remains file-local.
  }, [path])

  useEffect(() => {
    const view = viewRef.current
    if (!view) return
    const current = view.state.doc.toString()
    if (current !== value)
      view.dispatch({ changes: { from: 0, to: current.length, insert: value } })
  }, [value])

  useEffect(() => {
    const view = viewRef.current
    if (view) view.dispatch(setDiagnostics(view.state, editorDiagnostics(view, diagnostics)))
  }, [diagnostics])

  useEffect(() => {
    const view = viewRef.current
    if (!view || !revealLine) return
    const line = view.state.doc.line(Math.min(view.state.doc.lines, Math.max(1, revealLine)))
    view.dispatch({
      selection: { anchor: line.from },
      effects: EditorView.scrollIntoView(line.from, { y: 'center' }),
    })
    view.focus()
  }, [revealLine])

  return <div ref={hostRef} className="latex-editor" aria-label={`Editor: ${path}`} />
}
