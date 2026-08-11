import type { EditorDiagnostic } from './diagnostics.js'
import type { LatexBundleStatusDto } from '../../shared/ipc.js'
import { useLatexLocale } from '../i18n/locale.js'

export interface CompilePanelProps {
  compiling: boolean
  bundleStatus: LatexBundleStatusDto
  diagnostics: readonly EditorDiagnostic[]
  log: string
  onCompile: () => void
  onCancel: () => void
  onDiagnostic: (diagnostic: EditorDiagnostic) => void
}

export function CompilePanel({
  compiling,
  bundleStatus,
  diagnostics,
  log,
  onCompile,
  onCancel,
  onDiagnostic,
}: CompilePanelProps) {
  const { t } = useLatexLocale()
  const busy = compiling || bundleStatus.state === 'downloading'
  return (
    <section className="compile-panel">
      <header>
        <button type="button" className="compile-button" onClick={onCompile} disabled={busy}>
          {busy ? t('compiling') : t('compile')}
        </button>
        {busy && (
          <button type="button" onClick={onCancel}>
            {t('cancel')}
          </button>
        )}
      </header>
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
