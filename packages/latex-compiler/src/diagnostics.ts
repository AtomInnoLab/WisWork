import { isAbsolute, normalize, relative, resolve, sep } from 'node:path'

export interface TectonicDiagnostic {
  readonly path: string | null
  readonly line: number | null
  readonly column: number | null
  readonly severity: 'error' | 'warning'
  readonly message: string
}

function safePath(value: string, isolatedInputRoot?: string): string | null {
  if (value.includes('\0') || /^[a-zA-Z]:[\\/]/.test(value)) return null
  if (isAbsolute(value)) {
    if (!isolatedInputRoot) return null
    const root = resolve(isolatedInputRoot)
    const absolute = resolve(value)
    if (!absolute.startsWith(`${root}${sep}`)) return null
    return relative(root, absolute).split(sep).join('/')
  }
  const normalized = normalize(value).replaceAll('\\', '/')
  return normalized === '..' || normalized.startsWith('../') ? null : normalized
}

export function parseTectonicDiagnostics(
  log: string,
  isolatedInputRoot?: string,
): TectonicDiagnostic[] {
  const diagnostics: TectonicDiagnostic[] = []
  const lines = log.split(/\r?\n/)
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index] ?? ''
    const severityFirst = /^(error|warning):\s*(.*?):(\d+)(?::(\d+))?:\s*(.+)$/i.exec(line)
    if (severityFirst) {
      const path = safePath(severityFirst[2]!, isolatedInputRoot)
      if (path) {
        diagnostics.push({
          path,
          line: Number(severityFirst[3]),
          column: severityFirst[4] ? Number(severityFirst[4]) : null,
          severity: severityFirst[1]!.toLowerCase() as 'error' | 'warning',
          message: severityFirst[5]!.trim(),
        })
      }
      continue
    }
    const structured = /^(.*?):(\d+):(\d+):\s*(error|warning):\s*(.+)$/i.exec(line)
    if (structured) {
      const path = safePath(structured[1]!, isolatedInputRoot)
      if (path) {
        diagnostics.push({
          path,
          line: Number(structured[2]),
          column: Number(structured[3]),
          severity: structured[4]!.toLowerCase() as 'error' | 'warning',
          message: structured[5]!.trim(),
        })
      }
      continue
    }
    const tex = /^!\s*(.+)$/.exec(line)
    const location = tex ? /^l\.(\d+)\s/.exec(lines[index + 1] ?? '') : null
    if (tex) {
      diagnostics.push({
        path: null,
        line: location ? Number(location[1]) : null,
        column: null,
        severity: 'error',
        message: tex[1]!.trim(),
      })
    }
  }
  return diagnostics
}
