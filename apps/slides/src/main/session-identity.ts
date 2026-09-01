import { randomUUID } from 'node:crypto'

export interface SessionIdentityState {
  sessionInstanceId?: string
  documentInstanceId?: string
}

export function ensureSessionInstanceIds(session: SessionIdentityState): {
  sessionInstanceId: string
  documentInstanceId: string
} {
  session.sessionInstanceId ??= randomUUID()
  session.documentInstanceId ??= randomUUID()
  return {
    sessionInstanceId: session.sessionInstanceId,
    documentInstanceId: session.documentInstanceId,
  }
}
