import type { Session } from '../session-state'

export function blocksCanonicalPresentationTransaction(session: Session): boolean {
  // The AI host intentionally keeps a history batch open for the whole run so
  // all successful canonical writes collapse into one undo step. A history
  // batch is bookkeeping, not a competing mutation owner.
  return Boolean(session.masterEdit || session.transformPreview)
}
