import { useEffect, useState } from 'react'
import { buildLineDiff } from './diff.js'
import { reviewAction, type ReviewProposal, type ReviewVerification } from './proposal-review.js'

function rejectionRecovery(code: Extract<ReviewVerification, { state: 'rejected' }>['code']) {
  if (code === 'LATEX_CONFLICT') return 'Save your work or regenerate the proposal before applying.'
  if (code === 'LATEX_NOT_FOUND') return 'This proposal is missing or expired. Generate it again.'
  if (code === 'LATEX_VERIFICATION_REJECTED')
    return 'The safety policy rejected this proposal. Generate a safer proposal.'
  return 'Regenerate the proposal before applying.'
}

function VerificationEvidence({ verification }: { verification: ReviewVerification }) {
  if (verification.state === 'verifying') {
    return (
      <div className="proposal-verification verification-verifying">Verifying in isolation…</div>
    )
  }
  if (verification.state === 'rejected') {
    return (
      <div className="proposal-verification verification-rejected" role="alert">
        <strong>Verification rejected</strong>
        <p>
          {verification.message} ({verification.code})
        </p>
        <p>{rejectionRecovery(verification.code)}</p>
      </div>
    )
  }
  const { evidence } = verification
  const title =
    verification.state === 'verified'
      ? 'Verified'
      : verification.state === 'failed'
        ? 'Verification failed'
        : 'Verification unavailable'
  return (
    <div className={`proposal-verification verification-${verification.state}`}>
      <strong>{title}</strong>
      {'reason' in evidence && <p>{evidence.reason}</p>}
      <div>{evidence.diagnostics.length} diagnostics</div>
      {evidence.diagnostics.length > 0 && (
        <ul className="verification-diagnostics">
          {evidence.diagnostics.map((diagnostic, index) => (
            <li key={`${diagnostic.path ?? 'project'}:${diagnostic.line ?? 0}:${index}`}>
              {diagnostic.path ?? 'Project'}
              {diagnostic.line === null ? '' : `:${diagnostic.line}`}: {diagnostic.message}
            </li>
          ))}
        </ul>
      )}
      {evidence.logSummary && (
        <details>
          <summary>Verification log</summary>
          <pre>{evidence.logSummary}</pre>
        </details>
      )}
      <time dateTime={new Date(evidence.verifiedAt).toISOString()}>
        Verified at {new Date(evidence.verifiedAt).toLocaleString()}
      </time>
    </div>
  )
}

function ProposalFileDiff({ file }: { file: ReviewProposal['files'][number] }) {
  const diff = buildLineDiff(file.beforeText, file.afterText)
  return (
    <>
      <div className="proposal-file-summary">
        {diff.summary.added} addition{diff.summary.added === 1 ? '' : 's'}, {diff.summary.removed}{' '}
        removal{diff.summary.removed === 1 ? '' : 's'}
        {diff.truncated ? ' · preview truncated' : ''}
      </div>
      {diff.hunks.map((hunk, hunkIndex) => (
        <div className="proposal-hunk" key={`${hunk.beforeStart}:${hunk.afterStart}:${hunkIndex}`}>
          <div className="proposal-hunk-header">
            @@ -{hunk.beforeStart} +{hunk.afterStart} @@
          </div>
          {hunk.lines.map((line, lineIndex) => (
            <div className={`diff-line diff-line-${line.kind}`} key={lineIndex}>
              <span>{line.beforeLine ?? ''}</span>
              <span>{line.afterLine ?? ''}</span>
              <code>
                {line.kind === 'add' ? '+' : line.kind === 'remove' ? '−' : ' '} {line.text}
              </code>
            </div>
          ))}
        </div>
      ))}
    </>
  )
}

export function ProposalReview({
  proposal,
  busy,
  verification,
  onConfirm,
  onVerifySelection,
  onCancel,
}: {
  proposal: ReviewProposal
  busy: boolean
  verification: ReviewVerification
  onConfirm(selected: ReadonlySet<string>): void
  onVerifySelection(selected: ReadonlySet<string>): void
  onCancel(): void
}) {
  const [selected, setSelected] = useState(() => new Set(proposal.files.map((file) => file.path)))
  useEffect(() => setSelected(new Set(proposal.files.map((file) => file.path))), [proposal])
  const action = reviewAction(proposal, selected, verification)
  return (
    <section className="proposal-review" aria-label="AI edit proposal">
      <h3>Review AI changes</h3>
      <VerificationEvidence verification={verification} />
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
          <ProposalFileDiff file={file} />
        </article>
      ))}
      {(verification.state === 'failed' || verification.state === 'unverifiable') &&
        action.kind === 'apply' && (
          <p className="proposal-risk-warning" role="alert">
            This change has not compiled successfully in isolation. Applying it may break the
            project.
          </p>
        )}
      <button
        disabled={busy || action.disabled}
        onClick={() =>
          action.kind === 'verify-selection' ? onVerifySelection(selected) : onConfirm(selected)
        }
      >
        {action.label}
      </button>
      <button disabled={busy} onClick={onCancel}>
        Cancel
      </button>
    </section>
  )
}
