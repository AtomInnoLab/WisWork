import { useEffect, useRef } from 'react'
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
  const dialogRef = useRef<HTMLElement>(null)
  const cancelRef = useRef<HTMLButtonElement>(null)
  const busyRef = useRef(busy)
  const onCancelRef = useRef(onCancel)
  busyRef.current = busy
  onCancelRef.current = onCancel
  useEffect(() => {
    if (!open) return
    const previousFocus =
      document.activeElement instanceof HTMLElement ? document.activeElement : null
    cancelRef.current?.focus()
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !busyRef.current) {
        event.preventDefault()
        onCancelRef.current()
        return
      }
      if (event.key !== 'Tab') return
      const controls = Array.from(
        dialogRef.current?.querySelectorAll<HTMLElement>('button:not(:disabled)') ?? [],
      )
      if (!controls.length) return
      const first = controls[0]!
      const last = controls.at(-1)!
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault()
        last.focus()
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault()
        first.focus()
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => {
      window.removeEventListener('keydown', onKeyDown)
      previousFocus?.focus()
    }
  }, [open])
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
        ref={dialogRef}
        className="file-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="export-pdf-dialog-title"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <h3 id="export-pdf-dialog-title">{t('exportPdfStaleTitle')}</h3>
        <p>{t('exportPdfStaleMessage')}</p>
        <footer>
          <button ref={cancelRef} type="button" disabled={busy} onClick={onCancel}>
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
