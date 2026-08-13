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
      <div className="proposal-verification verification-verifying" role="status">
        Verifying in isolation…
      </div>
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
    <div
      className={`proposal-verification verification-${verification.state}`}
      role={verification.state === 'failed' ? 'alert' : 'status'}
    >
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
        {diff.summary.atLeast ? 'at least ' : ''}
        {diff.summary.added} addition{diff.summary.added === 1 ? '' : 's'},{' '}
        {diff.summary.atLeast ? 'at least ' : ''}
        {diff.summary.removed} removal{diff.summary.removed === 1 ? '' : 's'}
        {diff.truncated ? ' · preview truncated' : ''}
      </div>
      {diff.notice === 'change-location-beyond-preview-budget' && (
        <div className="proposal-diff-notice" role="status">
          Change location is beyond the preview budget. Review the source file before applying.
        </div>
      )}
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
  selection,
  busy,
  verification,
  riskArmed,
  onSelectionChange,
  onPrimaryAction,
  onCancel,
}: {
  proposal: ReviewProposal
  selection: ReadonlySet<string>
  busy: boolean
  verification: ReviewVerification
  riskArmed: boolean
  onSelectionChange(selected: ReadonlySet<string>): void
  onPrimaryAction(): void
  onCancel(): void
}) {
  const action = reviewAction(proposal, selection, verification, riskArmed)
  const risky =
    (verification.state === 'failed' || verification.state === 'unverifiable') &&
    action.kind !== 'verify-selection'
  return (
    <section
      className="proposal-review"
      aria-label="AI edit proposal"
      aria-busy={busy || verification.state === 'verifying'}
    >
      <h3>Review AI changes</h3>
      <VerificationEvidence verification={verification} />
      {proposal.files.map((file) => (
        <article key={file.path}>
          <label>
            <input
              type="checkbox"
              checked={selection.has(file.path)}
              disabled={busy}
              onChange={() =>
                onSelectionChange(
                  (() => {
                    const next = new Set(selection)
                    if (next.has(file.path)) next.delete(file.path)
                    else next.add(file.path)
                    return next
                  })(),
                )
              }
            />
            {file.path}
          </label>
          <ProposalFileDiff file={file} />
        </article>
      ))}
      {risky && (
        <p className="proposal-risk-warning" id="proposal-risk-warning" role="alert">
          {riskArmed
            ? 'Risk acknowledged. A separate confirmation will apply changes without successful verification.'
            : 'This change has not compiled successfully in isolation. Review this risk before enabling apply.'}
        </p>
      )}
      <button
        type="button"
        className={riskArmed && action.kind === 'apply' ? 'proposal-danger-button' : undefined}
        disabled={busy || action.disabled}
        aria-describedby={risky ? 'proposal-risk-warning' : undefined}
        onClick={onPrimaryAction}
      >
        {action.label}
      </button>
      <button type="button" disabled={busy} onClick={onCancel}>
        Cancel
      </button>
    </section>
  )
}
