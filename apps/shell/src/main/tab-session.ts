import { randomBytes } from 'node:crypto'
import { constants } from 'node:fs'
import { lstat, mkdir, open, rename, rm } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { openLatexProject } from '@wiswork/latex-project'

const SESSION_VERSION = 1
const MAX_SESSION_BYTES = 64 * 1024
const MAX_OPEN_PROJECTS = 50
const MAX_PATH_LENGTH = 4096

export interface TabSessionState {
  projectPaths: string[]
  activeProjectPath: string | null
}

const EMPTY_SESSION: TabSessionState = { projectPaths: [], activeProjectPath: null }

interface StoredTabSession {
  version: typeof SESSION_VERSION
  projectPaths: string[]
  activeProjectPath: string | null
}

function validPath(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0 && value.length <= MAX_PATH_LENGTH
}

function parseStored(value: unknown): StoredTabSession | null {
  if (!value || typeof value !== 'object') return null
  const record = value as Record<string, unknown>
  if (
    record.version !== SESSION_VERSION ||
    !Array.isArray(record.projectPaths) ||
    record.projectPaths.length > MAX_OPEN_PROJECTS ||
    !record.projectPaths.every(validPath) ||
    !(record.activeProjectPath === null || validPath(record.activeProjectPath))
  ) {
    return null
  }
  return {
    version: SESSION_VERSION,
    projectPaths: [...record.projectPaths],
    activeProjectPath: record.activeProjectPath as string | null,
  }
}

async function readBounded(path: string): Promise<string> {
  const before = await lstat(path)
  if (!before.isFile() || before.isSymbolicLink() || before.size > MAX_SESSION_BYTES) {
    throw new Error('Invalid tab session file')
  }
  const noFollow = 'O_NOFOLLOW' in constants ? constants.O_NOFOLLOW : 0
  const handle = await open(path, constants.O_RDONLY | noFollow)
  try {
    const opened = await handle.stat()
    if (!opened.isFile() || opened.size > MAX_SESSION_BYTES) {
      throw new Error('Invalid tab session file')
    }
    const bytes = Buffer.allocUnsafe(opened.size + 1)
    const { bytesRead } = await handle.read(bytes, 0, bytes.length, 0)
    if (bytesRead !== opened.size || (await handle.stat()).size !== opened.size) {
      throw new Error('Tab session changed while reading')
    }
    return bytes.subarray(0, bytesRead).toString('utf8')
  } finally {
    await handle.close()
  }
}

export async function loadTabSession(path: string): Promise<TabSessionState> {
  let stored: StoredTabSession | null
  try {
    stored = parseStored(JSON.parse(await readBounded(path)))
  } catch {
    return { ...EMPTY_SESSION }
  }
  if (!stored) return { ...EMPTY_SESSION }

  const projectPaths: string[] = []
  const canonicalByInput = new Map<string, string>()
  const seen = new Set<string>()
  for (const input of stored.projectPaths) {
    try {
      const project = await openLatexProject(input)
      if (!project.mainFile) continue
      canonicalByInput.set(input, project.rootPath)
      if (seen.has(project.rootPath)) continue
      seen.add(project.rootPath)
      projectPaths.push(project.rootPath)
    } catch {
      // A missing, malformed, linked-out, or concurrently changed project is skipped.
    }
  }
  const activeProjectPath = stored.activeProjectPath
    ? (canonicalByInput.get(stored.activeProjectPath) ?? null)
    : null
  return { projectPaths, activeProjectPath }
}

export async function saveTabSession(path: string, state: TabSessionState): Promise<void> {
  const projectPaths = state.projectPaths.filter(validPath).slice(0, MAX_OPEN_PROJECTS)
  const activeProjectPath =
    state.activeProjectPath && projectPaths.includes(state.activeProjectPath)
      ? state.activeProjectPath
      : null
  const bytes = Buffer.from(
    JSON.stringify({ version: SESSION_VERSION, projectPaths, activeProjectPath }),
    'utf8',
  )
  if (bytes.length > MAX_SESSION_BYTES) throw new Error('Tab session exceeds size limit')
  const parent = dirname(path)
  await mkdir(parent, { recursive: true })
  const temporary = join(parent, `.open-tabs-${randomBytes(12).toString('hex')}.tmp`)
  try {
    const handle = await open(temporary, 'wx', 0o600)
    try {
      await handle.writeFile(bytes)
      await handle.sync()
    } finally {
      await handle.close()
    }
    await rename(temporary, path)
    const directory = await open(parent, 'r')
    try {
      await directory.sync()
    } finally {
      await directory.close()
    }
  } catch (error) {
    await rm(temporary, { force: true }).catch(() => undefined)
    throw error
  }
}

type SaveSession = (state: TabSessionState) => Promise<void>

export class TabSessionPersistenceCoordinator {
  private tail: Promise<void> = Promise.resolve()
  private lastWriteSucceeded = true
  private generation = 0

  constructor(private readonly save: SaveSession) {}

  enqueue(state: TabSessionState): void {
    this.generation += 1
    const snapshot: TabSessionState = {
      projectPaths: [...state.projectPaths],
      activeProjectPath: state.activeProjectPath,
    }
    this.tail = this.tail
      .catch(() => undefined)
      .then(() => this.save(snapshot))
      .then(
        () => {
          this.lastWriteSucceeded = true
        },
        () => {
          this.lastWriteSucceeded = false
        },
      )
  }

  async flush(state: TabSessionState): Promise<boolean> {
    this.enqueue(state)
    let observedGeneration = -1
    while (observedGeneration !== this.generation) {
      observedGeneration = this.generation
      const observedTail = this.tail
      await observedTail
    }
    return this.lastWriteSucceeded
  }
}
