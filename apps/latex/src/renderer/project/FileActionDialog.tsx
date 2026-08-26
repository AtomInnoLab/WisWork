import { useEffect, useState } from 'react'

export type FileAction =
  { kind: 'create' } | { kind: 'rename'; path: string } | { kind: 'delete'; path: string }

export function FileActionDialog({
  action,
  busy,
  onCancel,
  onSubmit,
}: {
  action: FileAction | null
  busy: boolean
  onCancel: () => void
  onSubmit: (value?: string) => void
}) {
  const [value, setValue] = useState('')
  useEffect(() => {
    setValue(
      action?.kind === 'rename' ? action.path : action?.kind === 'create' ? 'chapter.tex' : '',
    )
  }, [action])
  if (!action) return null
  const destructive = action.kind === 'delete'
  const title =
    action.kind === 'create'
      ? 'New LaTeX file'
      : action.kind === 'rename'
        ? 'Rename file'
        : 'Delete file'
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
        aria-labelledby="file-dialog-title"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <h3 id="file-dialog-title">{title}</h3>
        {destructive ? (
          <p>
            Delete <strong>{action.path}</strong>? This cannot be undone.
          </p>
        ) : (
          <label>
            Project-relative path
            <input
              autoFocus
              value={value}
              disabled={busy}
              onChange={(event) => setValue(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter' && value.trim()) onSubmit(value.trim())
                if (event.key === 'Escape' && !busy) onCancel()
              }}
            />
          </label>
        )}
        <footer>
          <button type="button" disabled={busy} onClick={onCancel}>
            Cancel
          </button>
          <button
            type="button"
            className={destructive ? 'danger-button' : 'primary-button'}
            disabled={busy || (!destructive && !value.trim())}
            onClick={() => onSubmit(destructive ? undefined : value.trim())}
          >
            {busy
              ? 'Working…'
              : destructive
                ? 'Delete'
                : action.kind === 'create'
                  ? 'Create'
                  : 'Rename'}
          </button>
        </footer>
      </section>
    </div>
  )
}
