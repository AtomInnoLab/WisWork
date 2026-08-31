import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

describe('LaTeX workbench pane headers', () => {
  it('uses one exact border-box height across files, editor, PDF and AI panes', () => {
    const styles = readFileSync(new URL('../src/renderer/styles.css', import.meta.url), 'utf8')

    expect(styles).toMatch(/--latex-pane-header-height:\s*44px/)
    expect(styles).toMatch(/\.preview-stale\s*{[^}]*top:\s*var\(--latex-pane-header-height\)/s)
    expect(styles).toMatch(
      /\.project-tree header\s*{[^}]*height:\s*var\(--latex-pane-header-height\)[^}]*box-sizing:\s*border-box/s,
    )
    expect(styles).toMatch(
      /\.open-tabs\s*{[^}]*height:\s*var\(--latex-pane-header-height\)[^}]*box-sizing:\s*border-box/s,
    )
    expect(styles).toMatch(
      /\.readonly-pdf-viewer\s*{[^}]*grid-template-rows:\s*var\(--latex-pane-header-height\) minmax\(0, 1fr\)/s,
    )
    expect(styles).toMatch(
      /\.ai-panel-header\s*{[^}]*height:\s*var\(--latex-pane-header-height\)[^}]*box-sizing:\s*border-box/s,
    )
  })
})
