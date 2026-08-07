import { describe, expect, it } from 'vitest'
import { parseTectonicDiagnostics } from '../src/diagnostics.js'

describe('Tectonic diagnostics', () => {
  it('parses structured file/line/column severities and TeX errors', () => {
    expect(
      parseTectonicDiagnostics(
        'chapters/a.tex:12:3: warning: Underfull box\nmain.tex:4:1: error: Missing }',
      ),
    ).toEqual([
      {
        path: 'chapters/a.tex',
        line: 12,
        column: 3,
        severity: 'warning',
        message: 'Underfull box',
      },
      { path: 'main.tex', line: 4, column: 1, severity: 'error', message: 'Missing }' },
    ])
    expect(parseTectonicDiagnostics('! Undefined control sequence.\nl.9 \\badcommand')).toEqual([
      {
        path: null,
        line: 9,
        column: null,
        severity: 'error',
        message: 'Undefined control sequence.',
      },
    ])
  })

  it('rejects absolute and traversal diagnostic paths', () => {
    expect(parseTectonicDiagnostics('/etc/passwd:1:1: error: nope')).toEqual([])
    expect(parseTectonicDiagnostics('../outside.tex:1:1: error: nope')).toEqual([])
  })
  it('parses real Tectonic severity-first lines and maps isolated absolute paths', () => {
    const log = [
      'error: /tmp/job/input/main.tex:4: Undefined control sequence',
      'warning: chapters/a.tex:8: Underfull box',
    ].join('\n')
    expect(parseTectonicDiagnostics(log, '/tmp/job/input')).toEqual([
      {
        path: 'main.tex',
        line: 4,
        column: null,
        severity: 'error',
        message: 'Undefined control sequence',
      },
      {
        path: 'chapters/a.tex',
        line: 8,
        column: null,
        severity: 'warning',
        message: 'Underfull box',
      },
    ])
  })
})
