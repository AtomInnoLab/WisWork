import type { EditorDiagnostic } from './diagnostics.js'
import type { LatexBundleStatusDto } from '../../shared/ipc.js'
import { useLatexLocale } from '../i18n/locale.js'

export interface CompilePanelProps {
  compiling: boolean
  disabled?: boolean
  bundleStatus: LatexBundleStatusDto
  diagnostics: readonly EditorDiagnostic[]
  log: string
  onCompile: () => void
  onCancel: () => void
  onDiagnostic: (diagnostic: EditorDiagnostic) => void
  onAskAi: (diagnostic: EditorDiagnostic) => void
  showActions?: boolean
}

export function runCompilePanelAction(disabled: boolean, action: () => void): void {
  if (!disabled) action()
}

export function CompilePanel({
  compiling,
  disabled = false,
  bundleStatus,
  diagnostics,
  log,
  onCompile,
  onCancel,
  onDiagnostic,
  onAskAi,
  showActions = true,
}: CompilePanelProps) {
  const { t } = useLatexLocale()
  const busy = compiling || bundleStatus.state === 'downloading'
  return (
    <section className="compile-panel">
      {showActions && (
        <header>
          <button
            type="button"
            className="compile-button"
            onClick={() => runCompilePanelAction(disabled, onCompile)}
            disabled={busy || disabled}
          >
            {busy ? t('compiling') : t('compile')}
          </button>
          {busy && (
            <button
              type="button"
              onClick={() => runCompilePanelAction(disabled, onCancel)}
              disabled={disabled}
            >
              {t('cancel')}
            </button>
          )}
        </header>
      )}
      <div className="bundle-status" role="status">
        {bundleStatus.state === 'downloading'
          ? `Downloading TeX bundle (${Math.floor(
              (100 * bundleStatus.receivedBytes) / Math.max(1, bundleStatus.totalBytes),
            )}%)`
          : bundleStatus.state === 'ready'
            ? 'TeX bundle ready'
            : bundleStatus.state === 'remote'
              ? 'Remote TeX bundle configured'
              : bundleStatus.state === 'error'
                ? `TeX bundle error: ${bundleStatus.code}`
                : 'TeX bundle required on first compile'}
      </div>
      <details open={diagnostics.length > 0}>
        <summary>
          {t('diagnostics')} ({diagnostics.length})
        </summary>
        <ul className="diagnostic-list">
          {diagnostics.map((diagnostic, index) => (
            <li key={`${diagnostic.path}:${diagnostic.lineIndex}:${index}`}>
              <button
                type="button"
                aria-label={`Open ${diagnostic.path}:${diagnostic.lineIndex + 1}`}
                onClick={() => onDiagnostic(diagnostic)}
              >
                <strong>{diagnostic.severity}</strong> {diagnostic.path}:{diagnostic.lineIndex + 1}{' '}
                — {diagnostic.message}
              </button>
              <button
                type="button"
                className="diagnostic-ai-button"
                aria-label="Ask AI about this issue"
                title="Ask AI about this issue"
                onClick={() => onAskAi(diagnostic)}
              >
                Ask AI
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
