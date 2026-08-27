import { useEffect } from 'react'
import { useLatexLocale } from '../i18n/locale.js'

export function ExportPdfDialog({
  open,
  busy,
  onCancel,
  onCompile,
  onExportLast,
}: {
  open: boolean
  busy: boolean
  onCancel: () => void
  onCompile: () => void
  onExportLast: () => void
}) {
  const { t } = useLatexLocale()
  useEffect(() => {
    if (!open) return
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !busy) onCancel()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [busy, onCancel, open])
  if (!open) return null

  return (
    <div
      className="file-dialog-backdrop"
      role="presentation"
      onMouseDown={() => {
        if (!busy) onCancel()
      }}
    >
      <section
        className="file-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="export-pdf-dialog-title"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <h3 id="export-pdf-dialog-title">{t('exportPdfStaleTitle')}</h3>
        <p>{t('exportPdfStaleMessage')}</p>
        <footer>
          <button type="button" disabled={busy} onClick={onCancel}>
            {t('cancel')}
          </button>
          <button type="button" disabled={busy} onClick={onCompile}>
            {t('compileNow')}
          </button>
          <button type="button" className="primary-button" disabled={busy} onClick={onExportLast}>
            {busy ? t('exportingPdf') : t('exportLastPdf')}
          </button>
        </footer>
      </section>
    </div>
  )
}
