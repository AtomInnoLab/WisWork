import { randomUUID } from 'node:crypto'
import { resolve } from 'node:path'
import { assertPresentationId } from '@wiswork/project-store'
import { hasPresentationWorker, stopPresentationWorkers } from './presentation-jobs'

export interface PresentationProjectWorkScope {
  root: string
  projectId: string
  documentId: string
}
export interface PresentationProjectWork {
  readonly token: string
  readonly scope: Readonly<PresentationProjectWorkScope>
  readonly signal: AbortSignal
  readonly settled: Promise<void>
  /** Call only in the actual request's finally, after all work and writes have ended. */
  finish(): void
}
interface Entry {
  scope: Readonly<PresentationProjectWorkScope>
  controller: AbortController
  settled: Promise<void>
}
const foreground = new Map<string, Entry>()
function ownedScope(input: PresentationProjectWorkScope): Readonly<PresentationProjectWorkScope> {
  const { root, projectId, documentId } = input
  assertPresentationId(projectId)
  if (
    typeof root !== 'string' ||
    !root.trim() ||
    root.includes('\0') ||
    typeof documentId !== 'string' ||
    !documentId.trim() ||
    documentId.length > 4096 ||
    Array.from(documentId).some(
      (char) => char.charCodeAt(0) < 32 || (char.charCodeAt(0) >= 127 && char.charCodeAt(0) <= 159),
    )
  )
    throw Error('invalid_request')
  return Object.freeze({ root: resolve(root), projectId, documentId })
}

/** Each accepted foreground request owns a separate token; a later request never replaces it. */
export function registerPresentationProjectWork(options: {
  scope: PresentationProjectWorkScope
  signal?: AbortSignal
}): PresentationProjectWork {
  const scope = ownedScope(options.scope),
    clientSignal = options.signal,
    controller = new AbortController(),
    token = randomUUID()
  let done!: () => void
  const settled = new Promise<void>((resolve) => {
    done = resolve
  })
  const abort = () => {
    controller.abort()
  }
  if (clientSignal?.aborted) abort()
  else clientSignal?.addEventListener('abort', abort, { once: true })
  const entry = { scope, controller, settled }
  foreground.set(token, entry)
  return Object.freeze({
    token,
    scope,
    signal: controller.signal,
    settled,
    finish() {
      if (foreground.get(token) !== entry) return
      foreground.delete(token)
      clientSignal?.removeEventListener('abort', abort)
      done()
    },
  })
}

/** Read-only admission observation. A background worker for this project conservatively keeps any document scope busy. */
export function hasPresentationProjectWork(input: PresentationProjectWorkScope): boolean {
  const scope = ownedScope(input)
  return (
    Array.from(foreground.values()).some(
      (entry) =>
        entry.scope.root === scope.root &&
        entry.scope.projectId === scope.projectId &&
        entry.scope.documentId === scope.documentId,
    ) || hasPresentationWorker(`${scope.root}\0${scope.projectId}`)
  )
}

/** Call after persistent freeze and outside the project lock. Abort is a request; drain waits for actual completion. */
export async function stopPresentationProjectWork(
  input: PresentationProjectWorkScope,
): Promise<void> {
  const scope = ownedScope(input)
  const entries = Array.from(foreground.values()).filter(
    (entry) =>
      entry.scope.root === scope.root &&
      entry.scope.projectId === scope.projectId &&
      entry.scope.documentId === scope.documentId,
  )
  for (const entry of entries) entry.controller.abort()
  await Promise.all([stopPresentationWorkers(scope), ...entries.map((entry) => entry.settled)])
}
