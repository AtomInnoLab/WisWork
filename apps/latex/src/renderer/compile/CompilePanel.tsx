import type { EditorDiagnostic } from './diagnostics.js'
import { useLatexLocale } from '../i18n/locale.js'

export interface CompilePanelProps {
  compiling: boolean
  diagnostics: readonly EditorDiagnostic[]
  log: string
  onCompile: () => void
  onCancel: () => void
  onDiagnostic: (diagnostic: EditorDiagnostic) => void
}

export function CompilePanel({
  compiling,
  diagnostics,
  log,
  onCompile,
  onCancel,
  onDiagnostic,
}: CompilePanelProps) {
  const { t } = useLatexLocale()
  return (
    <section className="compile-panel">
      <header>
        <button type="button" className="compile-button" onClick={onCompile} disabled={compiling}>
          {compiling ? t('compiling') : t('compile')}
        </button>
        {compiling && (
          <button type="button" onClick={onCancel}>
            {t('cancel')}
          </button>
        )}
      </header>
      <details open={diagnostics.length > 0}>
        <summary>
          {t('diagnostics')} ({diagnostics.length})
        </summary>
        <ul className="diagnostic-list">
          {diagnostics.map((diagnostic, index) => (
            <li key={`${diagnostic.path}:${diagnostic.lineIndex}:${index}`}>
              <button type="button" onClick={() => onDiagnostic(diagnostic)}>
                <strong>{diagnostic.severity}</strong> {diagnostic.path}:{diagnostic.lineIndex + 1}{' '}
                — {diagnostic.message}
              </button>
            </li>
          ))}
        </ul>
      </details>
      <details>
        <summary>{t('logs')}</summary>
        <pre>{log || t('noDiagnostics')}</pre>
      </details>
    </section>
  )
}
