import { createHash, randomUUID } from 'node:crypto'
import {
  constants,
  openSync,
  closeSync,
  fsyncSync,
  fstatSync,
  lstatSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmdirSync,
  unlinkSync,
  writeFileSync,
  type BigIntStats,
} from 'node:fs'
import { open } from 'node:fs/promises'
import { dirname, join, relative, resolve, sep } from 'node:path'
import {
  PresentationLifecycleStore,
  type PresentationDeletionResult,
  type PresentationLifecycleScope,
} from '@wiswork/project-store'
import {
  inspectPresentationProjectInventory,
  MAX_PRESENTATION_INVENTORY_FILES,
  MAX_PRESENTATION_INVENTORY_BYTES,
} from './presentation-project-inventory'
const MAX_PROOF_BYTES = 16 * 1024 * 1024
const busy = new Set<string>()
const hash = (v: string) => createHash('sha256').update(v).digest('hex')
interface Identity {
  dev: string
  ino: string
  mode: string
}
interface Entry {
  relative: string
  identity: Identity
  type: 'file' | 'directory'
  size?: number
  mtime?: string
  sha256?: string
}
interface Proof {
  version: 1
  scope: PresentationLifecycleScope
  deletionId: string
  resourceId: string
  source: string
  entries: Entry[]
}
const identity = (s: BigIntStats): Identity => ({
  dev: String(s.dev),
  ino: String(s.ino),
  mode: String(s.mode),
})
const sameIdentity = (s: BigIntStats, i: Identity) =>
  JSON.stringify(identity(s)) === JSON.stringify(i)
const stat = (p: string) => lstatSync(p, { bigint: true, throwIfNoEntry: false })
function fail(code = 'ownership_unproven'): never {
  throw Error(code)
}
function checkedPath(p: string) {
  for (let current = resolve(p); ; current = dirname(current)) {
    const s = stat(current)
    if (s && (s.isSymbolicLink() || (current !== resolve(p) && !s.isDirectory()))) fail()
    if (dirname(current) === current) break
  }
}
function walk(root: string, guard: () => void): Entry[] {
  const entries: Entry[] = []
  let bytes = 0,
    files = 0
  const visit = (path: string, depth: number) => {
    guard()
    checkedPath(path)
    const s = stat(path)
    if (!s) return
    if (depth > 16 || entries.length >= 2 * MAX_PRESENTATION_INVENTORY_FILES) fail()
    if (s.isSymbolicLink() || (!s.isFile() && !s.isDirectory())) fail()
    const name = relative(root, path).split(sep).join('/')
    if (
      name
        .split('/')
        .some(
          (part) =>
            part === '..' || part.includes('\\') || part.includes('\0') || part.includes(':'),
        )
    )
      fail()
    if (s.isFile()) {
      const size = Number(s.size)
      if (
        !Number.isSafeInteger(size) ||
        ++files > MAX_PRESENTATION_INVENTORY_FILES ||
        (bytes += size) > MAX_PRESENTATION_INVENTORY_BYTES
      )
        fail()
      entries.push({
        relative: name,
        identity: identity(s),
        type: 'file',
        size,
        mtime: String(s.mtimeNs),
      })
    } else {
      entries.push({ relative: name, identity: identity(s), type: 'directory' })
      for (const child of readdirSync(path).sort()) visit(join(path, child), depth + 1)
    }
  }
  visit(root, 0)
  return entries
}
async function digest(path: string, entry: Entry, guard: () => void) {
  guard()
  checkedPath(path)
  const fd = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK)
  try {
    guard()
    const first = await fd.stat({ bigint: true })
    guard()
    if (
      !first.isFile() ||
      !sameIdentity(first, entry.identity) ||
      Number(first.size) !== entry.size ||
      String(first.mtimeNs) !== entry.mtime
    )
      fail()
    const h = createHash('sha256'),
      buffer = Buffer.alloc(512 * 1024)
    let offset = 0
    while (offset < entry.size!) {
      guard()
      const { bytesRead } = await fd.read(
        buffer,
        0,
        Math.min(buffer.length, entry.size! - offset),
        offset,
      )
      guard()
      if (!bytesRead) fail()
      h.update(buffer.subarray(0, bytesRead))
      offset += bytesRead
    }
    const last = await fd.stat({ bigint: true })
    guard()
    const current = stat(path)
    if (
      !sameIdentity(last, entry.identity) ||
      Number(last.size) !== entry.size ||
      String(last.mtimeNs) !== entry.mtime ||
      last.ctimeNs !== first.ctimeNs ||
      !current ||
      !sameIdentity(current, entry.identity) ||
      Number(current.size) !== entry.size ||
      String(current.mtimeNs) !== entry.mtime ||
      current.ctimeNs !== first.ctimeNs
    )
      fail()
    return { sha256: h.digest('hex'), ctime: String(last.ctimeNs) }
  } finally {
    await fd.close()
  }
}
function sourceFor(
  root: string,
  scope: PresentationLifecycleScope,
  kind: string,
): string | undefined {
  if (kind === 'project') return join(root, 'projects', 'presentations', hash(scope.projectId))
  const namespaces: Record<string, string> = {
    research: 'presentation-research',
    delivery_bundles: 'presentation-delivery-bundles',
    page_backups: 'presentation-page-backups',
    preferences: 'presentation-preferences',
    comments: 'presentation-comments',
    manual_observations: 'presentation-manual-observations',
  }
  const namespace = namespaces[kind]
  if (!namespace) return
  if (kind === 'page_backups') return join(root, namespace, hash(scope.projectId))
  if (['preferences', 'comments', 'manual_observations'].includes(kind))
    return join(
      root,
      namespace,
      hash(JSON.stringify([scope.documentId, scope.projectId])) + '.json',
    )
  return join(root, namespace, hash(scope.documentId), hash(scope.projectId))
}
function readProof(path: string, expected: Omit<Proof, 'entries'>): Proof {
  checkedPath(path)
  const s = stat(path)
  if (!s?.isFile() || s.size > BigInt(MAX_PROOF_BYTES)) fail()
  const raw = readFileSync(path, 'utf8'),
    envelope = JSON.parse(raw) as { proof: Proof; checksum: string }
  const p = envelope.proof
  if (
    !p ||
    envelope.checksum !== hash(JSON.stringify(p)) ||
    p.version !== 1 ||
    JSON.stringify(p.scope) !== JSON.stringify(expected.scope) ||
    p.deletionId !== expected.deletionId ||
    p.resourceId !== expected.resourceId ||
    p.source !== expected.source ||
    !Array.isArray(p.entries) ||
    !p.entries.length ||
    p.entries.length > 2 * MAX_PRESENTATION_INVENTORY_FILES
  )
    fail()
  const names = new Set<string>()
  for (const e of p.entries) {
    if (
      !e ||
      typeof e.relative !== 'string' ||
      e.relative.startsWith('/') ||
      e.relative
        .split('/')
        .some(
          (v) => v === '..' || v === '.' || v.includes('\\') || v.includes('\0') || v.includes(':'),
        ) ||
      names.has(e.relative) ||
      !['file', 'directory'].includes(e.type) ||
      !e.identity ||
      Object.values(e.identity).some((v) => typeof v !== 'string' || !/^[0-9]+$/.test(v))
    )
      fail()
    if (
      e.type === 'file' &&
      (!Number.isSafeInteger(e.size) ||
        e.size! < 0 ||
        typeof e.mtime !== 'string' ||
        !/^-?[0-9]+$/.test(e.mtime) ||
        typeof e.sha256 !== 'string' ||
        !/^[a-f0-9]{64}$/.test(e.sha256))
    )
      fail()
    names.add(e.relative)
  }
  if (!names.has('')) fail()
  return p
}
/** Internal deletion primitive. Caller freezes/drains and holds the project lock; no public route is enabled here. */
export async function removePresentationProjectDeletionResource(options: {
  userDataPath: string
  scope: PresentationLifecycleScope
  deletionId: string
  resourceId: string
  expectedRevision: number
  assertDeleting?: () => void
}): Promise<PresentationDeletionResult> {
  const root = resolve(options.userDataPath),
    scope = Object.freeze({
      projectId: options.scope.projectId,
      documentId: options.scope.documentId,
    }),
    deletionId = options.deletionId,
    resourceId = options.resourceId,
    revision = options.expectedRevision,
    extraGuard = options.assertDeleting
  const store = new PresentationLifecycleStore(root),
    key = root + '\0' + scope.projectId + '\0' + resourceId
  const result = (
    status: PresentationDeletionResult['status'],
    code?: PresentationDeletionResult['code'],
  ): PresentationDeletionResult => ({ deletionId, resourceId, status, ...(code ? { code } : {}) })
  const anchors = new Map<string, Identity>()
  const remember = (path: string) => {
    checkedPath(path)
    for (let current = resolve(path); ; current = dirname(current)) {
      const s = stat(current)
      if (s) {
        if (!s.isDirectory()) fail()
        const prior = anchors.get(current)
        if (prior && !sameIdentity(s, prior)) fail()
        anchors.set(current, identity(s))
      }
      if (dirname(current) === current) break
    }
  }
  const checkAnchors = () => {
    for (const [path, expected] of anchors) {
      const current = stat(path)
      if (!current?.isDirectory() || !sameIdentity(current, expected)) fail()
    }
  }
  const guard = () => {
    checkAnchors()
    const r = store.read(scope)
    if (
      !r ||
      r.revision !== revision ||
      r.state !== 'deleting' ||
      r.deletion?.deletionId !== deletionId
    )
      fail('resource_busy')
    extraGuard?.()
    checkAnchors()
    const next = store.read(scope)
    if (
      !next ||
      next.revision !== revision ||
      next.state !== 'deleting' ||
      next.deletion?.deletionId !== deletionId
    )
      fail('resource_busy')
  }
  const syncDirectory = (path: string) => {
    guard()
    checkedPath(path)
    if (process.platform === 'win32') return
    const fd = openSync(path, constants.O_RDONLY | constants.O_NOFOLLOW)
    try {
      guard()
      const actual = fstatSync(fd, { bigint: true }),
        expected = anchors.get(resolve(path))
      if (!actual.isDirectory() || !expected || !sameIdentity(actual, expected)) fail()
      fsyncSync(fd)
    } finally {
      closeSync(fd)
    }
  }
  const syncProofParents = (path: string) => {
    for (let current = resolve(path); ; current = dirname(current)) {
      syncDirectory(current)
      if (current === root) break
      if (dirname(current) === current) fail()
    }
  }
  if (busy.has(key)) return result('failed', 'resource_busy')
  busy.add(key)
  try {
    remember(root)
    guard()
    const item = store.read(scope)!.deletion!.resources.find((r) => r.resourceId === resourceId)
    if (!item) fail()
    if (item.ownership !== 'project_exclusive')
      return result(
        'retained',
        item.ownership === 'shared_reference' ? 'shared_resource' : 'ownership_unproven',
      )
    const source = sourceFor(root, scope, item.kind)
    if (!source) return result('retained', 'ownership_unproven')
    const expectedId =
      'resource_' +
      hash(
        JSON.stringify([
          'presentation-inventory-v1',
          scope.documentId,
          scope.projectId,
          item.kind,
          relative(root, source),
        ]),
      )
    if (expectedId !== resourceId) fail()
    const work = join(
        root,
        'presentation-project-deletion-work',
        hash(scope.projectId),
        hash(deletionId),
        hash(resourceId),
      ),
      holding = join(work, 'data'),
      proofPath = join(work, 'proof.json')
    checkedPath(source)
    remember(dirname(source))
    remember(work)
    checkedPath(work)
    checkedPath(holding)
    if (stat(source) && stat(holding)) fail('resource_busy')
    if (!stat(source) && !stat(holding) && !stat(proofPath)) return result('removed', 'not_found')
    const expected = {
      version: 1 as const,
      scope,
      deletionId,
      resourceId,
      source: relative(root, source),
    }
    let proof: Proof
    if (stat(proofPath)) proof = readProof(proofPath, expected)
    else {
      if (stat(holding)) fail()
      const inventory = await inspectPresentationProjectInventory({ userDataPath: root, ...scope })
      guard()
      const candidate = inventory.resources.find((r) => r.resourceId === resourceId)
      if (candidate?.ownership !== 'project_exclusive' || candidate.kind !== item.kind) fail()
      const entries = walk(source, guard)
      for (const e of entries)
        if (e.type === 'file')
          e.sha256 = (await digest(join(source, ...e.relative.split('/')), e, guard)).sha256
      proof = { ...expected, entries }
      const raw = JSON.stringify({ proof, checksum: hash(JSON.stringify(proof)) })
      if (Buffer.byteLength(raw) > MAX_PROOF_BYTES) fail()
      guard()
      checkedPath(work)
      mkdirSync(work, { recursive: true, mode: 0o700 })
      checkedPath(work)
      remember(work)
      const temporary = join(work, randomUUID() + '.tmp')
      let owned = false
      let fd: number | undefined
      let ownedStat: BigIntStats | undefined
      try {
        guard()
        fd = openSync(
          temporary,
          constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | constants.O_NOFOLLOW,
          0o600,
        )
        owned = true
        ownedStat = fstatSync(fd, { bigint: true })
        guard()
        checkedPath(temporary)
        const linked = stat(temporary)
        if (!linked || !sameIdentity(linked, identity(ownedStat))) fail()
        writeFileSync(fd, raw)
        ownedStat = fstatSync(fd, { bigint: true })
        guard()
        fsyncSync(fd)
        guard()
        checkedPath(proofPath)
        if (stat(proofPath)) fail('resource_busy')
        const current = fstatSync(fd, { bigint: true }),
          finalLinked = stat(temporary)
        if (
          !finalLinked?.isFile() ||
          !sameIdentity(finalLinked, identity(ownedStat)) ||
          !sameIdentity(current, identity(ownedStat)) ||
          current.size !== ownedStat.size ||
          current.ctimeNs !== ownedStat.ctimeNs ||
          finalLinked.size !== current.size ||
          finalLinked.ctimeNs !== current.ctimeNs
        )
          fail()
        renameSync(temporary, proofPath)
        owned = false
        syncProofParents(work)
      } finally {
        if (owned && fd !== undefined && ownedStat) {
          try {
            const current = fstatSync(fd, { bigint: true }),
              linked = stat(temporary)
            // A held exclusive fd proves only this invocation's unchanged staging leaf.
            if (
              linked?.isFile() &&
              sameIdentity(linked, identity(ownedStat)) &&
              sameIdentity(current, identity(ownedStat)) &&
              current.size === ownedStat.size &&
              current.ctimeNs === ownedStat.ctimeNs &&
              linked.size === current.size &&
              linked.ctimeNs === current.ctimeNs
            )
              unlinkSync(temporary)
          } catch {
            /* uncertain staging remains intact */
          }
        }
        if (fd !== undefined) closeSync(fd)
      }
    }
    const currentRoot = stat(source) ? source : holding
    const current = walk(currentRoot, guard),
      allowed = new Map(proof.entries.map((e) => [e.relative, e]))
    for (const e of current) {
      const expectedEntry = allowed.get(e.relative)
      if (
        !expectedEntry ||
        e.type !== expectedEntry.type ||
        JSON.stringify(e.identity) !== JSON.stringify(expectedEntry.identity)
      )
        fail()
      if (
        e.type === 'file' &&
        (await digest(join(currentRoot, ...e.relative.split('/')), expectedEntry, guard)).sha256 !==
          expectedEntry.sha256
      )
        fail()
    }
    // Missing entries are accepted only in a previously quarantined partial cleanup.
    if (currentRoot === source && current.length !== proof.entries.length) fail()
    if (currentRoot === source) {
      guard()
      checkedPath(source)
      checkedPath(holding)
      if (stat(holding)) fail('resource_busy')
      const original = stat(source),
        rootEntry = allowed.get('')!
      if (!original || !sameIdentity(original, rootEntry.identity)) fail()
      renameSync(source, holding)
    }
    // Persist both rename parents before beginning irreversible leaf cleanup.
    syncDirectory(work)
    syncDirectory(dirname(source))
    for (const e of current.filter((e) => e.type === 'file')) {
      guard()
      const path = join(holding, ...e.relative.split('/'))
      checkedPath(path)
      const original = allowed.get(e.relative)!
      const verified = await digest(path, original, guard)
      if (verified.sha256 !== original.sha256) fail()
      guard()
      const s = stat(path)
      if (
        !s?.isFile() ||
        !sameIdentity(s, original.identity) ||
        Number(s.size) !== original.size ||
        String(s.mtimeNs) !== original.mtime ||
        String(s.ctimeNs) !== verified.ctime
      )
        fail()
      unlinkSync(path)
    }
    for (const e of current
      .filter((e) => e.type === 'directory')
      .sort(
        (a, b) =>
          b.relative.split('/').length - a.relative.split('/').length ||
          b.relative.length - a.relative.length,
      )) {
      guard()
      const path = join(holding, ...e.relative.split('/'))
      checkedPath(path)
      const s = stat(path)
      if (!s?.isDirectory() || !sameIdentity(s, e.identity)) fail()
      rmdirSync(path)
    }
    guard()
    if (stat(source) || stat(holding)) fail('resource_busy')
    checkedPath(proofPath)
    checkAnchors()
    syncDirectory(work)
    guard()
    unlinkSync(proofPath)
    syncDirectory(work)
    return result('removed')
  } catch (error) {
    let code = ''
    try {
      code = error instanceof Error ? error.message : ''
    } catch {
      /* finite result only */
    }
    return code === 'resource_busy'
      ? result('failed', 'resource_busy')
      : code === 'ownership_unproven' || code.startsWith('presentation_inventory_')
        ? result('retained', 'ownership_unproven')
        : result('failed', 'io_failed')
  } finally {
    busy.delete(key)
  }
}
