import React from 'react'
import type { EditQueueSnapshot, SelectionScope } from './edit-queue'
import { selectionScopeSummary } from './edit-queue'

export function EditQueueCard({
  scope,
  queue,
  onCancel,
  onResume,
}: {
  scope: SelectionScope
  queue: EditQueueSnapshot
  onCancel: () => void
  onResume?: () => void
}) {
  return (
    <section className="ai-edit-queue-card" aria-label="Selection edit queue">
      <span>{selectionScopeSummary(scope)}</span>
      <span role="status">
        {queue.paused
          ? 'Paused for review'
          : queue.running
            ? `Running${queue.queued ? ` · ${queue.queued} queued` : ''}`
            : `${queue.queued} queued`}
      </span>
      {queue.paused && onResume && <button onClick={onResume}>Continue</button>}
      {(queue.running || queue.queued > 0 || queue.paused) && (
        <button onClick={onCancel}>Cancel</button>
      )}
    </section>
  )
}
