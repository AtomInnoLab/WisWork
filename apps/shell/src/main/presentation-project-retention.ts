import { randomUUID, createHash } from 'node:crypto'
import {
  closeSync,
  constants,
  fstatSync,
  lstatSync,
  openSync,
  readSync,
  readdirSync,
} from 'node:fs'
import { join, resolve } from 'node:path'
import {
  PresentationLifecycleStore,
  PresentationStore,
  MAX_PRESENTATION_LIFECYCLE_BYTES,
  type PresentationLifecycleScope,
} from '@wiswork/project-store'
import { inspectPresentationProjectInventory } from './presentation-project-inventory'
import { createPresentationProjectDeletionService } from './presentation-project-deletion'

const MAX_CONTROLS = 32768
const hash = (value: string) => createHash('sha256').update(value).digest('hex')

export function presentationRetentionEnabled(env: NodeJS.ProcessEnv): boolean {
  return (
    env.WISWORK_PPT_PROJECT_GOVERNANCE_ENABLED === '1' &&
    env.WISWORK_PPT_RETENTION_AUTOMATION_ENABLED === '1'
  )
}

/** Hints only. The deletion service revalidates scope, policy, activity and Work under the writer lock. */
function controlScopes(root: string): PresentationLifecycleScope[] {
  const base = join(root, 'presentation-project-lifecycles')
  for (const path of [root, base]) {
    const stat = lstatSync(path, { throwIfNoEntry: false })
    if (!stat) return []
    if (!stat.isDirectory() || stat.isSymbolicLink()) throw Error('invalid_state')
  }
  const names = readdirSync(base)
  if (names.length > MAX_CONTROLS) throw Error('output_too_large')
  const result: PresentationLifecycleScope[] = []
  for (const name of names.sort()) {
    if (!/^[a-f0-9]{64}$/.test(name)) continue
    const directory = join(base, name)
    const before = lstatSync(directory, { throwIfNoEntry: false })
    if (!before || !before.isDirectory() || before.isSymbolicLink()) continue
    const path = join(directory, 'lifecycle.json')
    const file = lstatSync(path, { throwIfNoEntry: false })
    if (
      !file ||
      !file.isFile() ||
      file.isSymbolicLink() ||
      !Number.isSafeInteger(file.size) ||
      file.size > MAX_PRESENTATION_LIFECYCLE_BYTES
    )
      continue
    let fd: number | undefined
    try {
      fd = openSync(path, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK)
      const opened = fstatSync(fd)
      if (opened.ino !== file.ino || opened.dev !== file.dev || opened.size !== file.size) continue
      const bounded = Buffer.alloc(file.size + 1)
      let length = 0
      while (length < bounded.length) {
        const size = readSync(fd, bounded, length, bounded.length - length, length)
        if (size === 0) break
        length += size
      }
      if (length !== file.size) continue
      const after = fstatSync(fd)
      const current = lstatSync(directory)
      if (
        after.ino !== file.ino ||
        after.dev !== file.dev ||
        after.size !== file.size ||
        after.mtimeMs !== file.mtimeMs ||
        after.ctimeMs !== file.ctimeMs ||
        current.ino !== before.ino ||
        current.dev !== before.dev
      )
        continue
      const raw: unknown = JSON.parse(bounded.subarray(0, length).toString('utf8'))
      if (!raw || typeof raw !== 'object' || Array.isArray(raw)) continue
      const { projectId, documentId } = raw as Record<string, unknown>
      if (
        typeof projectId !== 'string' ||
        !/^[A-Za-z0-9_-]{1,128}$/.test(projectId) ||
        hash(projectId) !== name ||
        typeof documentId !== 'string' ||
        !documentId.trim() ||
        documentId.length > 4096
      )
        continue
      result.push({ projectId, documentId })
    } catch {
      // A malformed or replaced control is ineligible. Never repair it during a retention scan.
    } finally {
      if (fd !== undefined) closeSync(fd)
    }
  }
  return result
}

/** One bounded local retention pass. Default null policy performs no deletion. */
export function createPresentationProjectRetentionService(options: {
  userDataPath: string
  acquireProjectLock: (projectId: string) => Promise<() => void>
  now?: () => Date
}) {
  const root = resolve(options.userDataPath)
  const life = new PresentationLifecycleStore(root)
  const store = new PresentationStore(root)
  const deletion = createPresentationProjectDeletionService({
    userDataPath: root,
    acquireProjectLock: options.acquireProjectLock,
  })
  let running = false
  return {
    async tick() {
      if (running) return { considered: 0, started: 0, resumed: 0, skipped: 0 }
      running = true
      const counts = { considered: 0, started: 0, resumed: 0, skipped: 0 }
      try {
        const now = (options.now ?? (() => new Date()))()
        if (!(now instanceof Date) || !Number.isFinite(now.getTime()))
          throw Error('invalid_request')
        for (const scope of controlScopes(root)) {
          counts.considered++
          try {
            const record = life.readControl(scope)
            if (!record) {
              counts.skipped++
              continue
            }
            if (record.state === 'deleting' && record.deletion?.reason === 'retention') {
              await deletion.resume({
                scope,
                expectedRevision: record.revision,
                deletionId: record.deletion.deletionId,
              })
              counts.resumed++
              continue
            }
            if (record.state !== 'active' || record.policy.contentRetentionDays === null) {
              counts.skipped++
              continue
            }
            if (!store.projectScope(scope.projectId, scope.documentId)) {
              counts.skipped++
              continue
            }
            const saved = [...record.audit]
              .reverse()
              .find((event) => event.action === 'policy_updated')
            const policySavedAt = saved?.at ?? record.audit[0]?.at
            if (!policySavedAt) {
              counts.skipped++
              continue
            }
            const inventory = await inspectPresentationProjectInventory({
              userDataPath: root,
              ...scope,
              observeActivity: { policySavedAt },
            })
            const activity = inventory.activity
            const latest = activity?.latestActivityAt
            if (
              !activity?.decidable ||
              !latest ||
              Date.parse(latest) > now.getTime() ||
              now.getTime() - Date.parse(latest) < record.policy.contentRetentionDays * 86400000
            ) {
              counts.skipped++
              continue
            }
            await deletion.retentionConfirm({
              scope,
              expectedRevision: record.revision,
              policySavedAt,
              latestActivityAt: latest,
              now: now.toISOString(),
              deletionId: randomUUID(),
            })
            counts.started++
          } catch {
            counts.skipped++
          }
        }
        return counts
      } finally {
        running = false
      }
    },
  }
}
