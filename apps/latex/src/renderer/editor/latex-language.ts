import { bracketMatching } from '@codemirror/language'
import { latex } from 'codemirror-lang-latex'

export function latexLanguageExtensions(fileName: string) {
  return [
    bracketMatching(),
    latex({
      fileName,
      autoCloseTags: true,
      enableAutocomplete: true,
      enableLinting: false,
      enableTooltips: true,
    }),
  ]
}
