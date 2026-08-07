export interface CompileDiagnosticInput {
  path: string
  line: number
  column: number | null
  severity: 'error' | 'warning'
  message: string
}

export interface EditorDiagnostic {
  path: string
  lineIndex: number
  columnIndex: number
  severity: 'error' | 'warning'
  message: string
}

function safeProjectPath(path: string): string | null {
  const normalized = path.replaceAll('\\', '/')
  if (
    !normalized ||
    normalized.startsWith('/') ||
    normalized.includes('\0') ||
    normalized.split('/').some((part) => part === '..')
  ) {
    return null
  }
  return normalized.replace(/^\.\//, '')
}

export function mapCompileDiagnostics(
  diagnostics: readonly CompileDiagnosticInput[],
  projectFiles: ReadonlySet<string>,
): EditorDiagnostic[] {
  return diagnostics.flatMap((diagnostic) => {
    const path = safeProjectPath(diagnostic.path)
    if (!path || !projectFiles.has(path) || !Number.isFinite(diagnostic.line)) return []
    return [
      {
        path,
        lineIndex: Math.max(0, Math.trunc(diagnostic.line) - 1),
        columnIndex:
          diagnostic.column === null || !Number.isFinite(diagnostic.column)
            ? 0
            : Math.max(0, Math.trunc(diagnostic.column) - 1),
        severity: diagnostic.severity,
        message: diagnostic.message,
      },
    ]
  })
}
