import { useEffect, useState } from 'react'
import type { ReviewProposal } from './proposal-review.js'

export function ProposalReview({
  proposal,
  busy,
  onConfirm,
  onCancel,
}: {
  proposal: ReviewProposal
  busy: boolean
  onConfirm(selected: ReadonlySet<string>): void
  onCancel(): void
}) {
  const [selected, setSelected] = useState(() => new Set(proposal.files.map((file) => file.path)))
  useEffect(() => setSelected(new Set(proposal.files.map((file) => file.path))), [proposal])
  return (
    <section className="proposal-review" aria-label="AI edit proposal">
      <h3>Review AI changes</h3>
      {proposal.files.map((file) => (
        <article key={file.path}>
          <label>
            <input
              type="checkbox"
              checked={selected.has(file.path)}
              disabled={busy}
              onChange={() =>
                setSelected((current) => {
                  const next = new Set(current)
                  if (next.has(file.path)) next.delete(file.path)
                  else next.add(file.path)
                  return next
                })
              }
            />
            {file.path}
          </label>
          <div className="proposal-diff">
            <pre className="proposal-before">{file.beforeText ?? '(new file)'}</pre>
            <pre className="proposal-after">{file.afterText}</pre>
          </div>
        </article>
      ))}
      <button disabled={busy || selected.size === 0} onClick={() => onConfirm(selected)}>
        Confirm selected changes
      </button>
      <button disabled={busy} onClick={onCancel}>
        Cancel
      </button>
    </section>
  )
}
