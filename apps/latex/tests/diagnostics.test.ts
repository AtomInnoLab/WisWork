import { describe, expect, it } from 'vitest'
import { mapCompileDiagnostics } from '../src/renderer/compile/diagnostics.js'

describe('compile diagnostics mapping', () => {
  it('maps safe project paths and one-based source locations to editor coordinates', () => {
    expect(
      mapCompileDiagnostics(
        [
          { path: 'chapters/a.tex', line: 7, column: 3, severity: 'error', message: 'bad' },
          { path: '../outside.tex', line: 1, column: 1, severity: 'warning', message: 'skip' },
          { path: 'missing.tex', line: 2, column: null, severity: 'warning', message: 'skip' },
        ],
        new Set(['main.tex', 'chapters/a.tex']),
      ),
    ).toEqual([
      {
        path: 'chapters/a.tex',
        lineIndex: 6,
        columnIndex: 2,
        severity: 'error',
        message: 'bad',
      },
    ])
  })
})
