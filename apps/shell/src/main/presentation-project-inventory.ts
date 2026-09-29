import { createHash } from 'node:crypto'
import {
  constants,
  closeSync,
  fstatSync,
  lstatSync,
  openSync,
  readFileSync,
  readdirSync,
  type Stats,
} from 'node:fs'
import { dirname, join, relative, resolve, sep } from 'node:path'
import { parsePresentationAcquisitionHistory } from '@wiswork/project-store/presentation-acquisition'
import { parsePresentationResearchRecord } from '@wiswork/project-store/presentation-research'
import { parsePresentationDeliveryBundleReceipt } from '@wiswork/project-store/presentation-delivery-bundle'
import { parsePresentationTeamLedger } from '@wiswork/pptx-engine/presentation-team'
import { parseSavedPresentationPreference } from '@wiswork/pptx-engine/presentation-preference'
import { parsePresentationManualObservation } from '@wiswork/pptx-engine/presentation-manual-observation'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
export const MAX_PRESENTATION_INVENTORY_BYTES = 32 * 1024 * 1024 * 1024
export const MAX_PRESENTATION_INVENTORY_FILES = 32768
const PROJECT_JSON_BYTES = 17 * 1024 * 1024
const RESEARCH_JSON_BYTES = 64 * 1024 * 1024
export type PresentationInventoryKind =
  | 'project'
  | 'research'
  | 'delivery_bundles'
  | 'page_backups'
  | 'preferences'
  | 'comments'
  | 'manual_observations'
  | 'teams'
  | 'attachments'
  | 'acquisition_history'
  | 'existing_page_backups'
  | 'existing_page_releases'
  | 'master_backups'
  | 'package_backups'
  | 'brand_reference'
  | 'unbound_page_staging'
  | 'unbound_existing_staging'
  | 'lifecycle_control'
export interface PresentationInventoryResource {
  kind: PresentationInventoryKind
  ownership: 'project_exclusive' | 'document_shared' | 'global_shared' | 'unproven'
  disposition: 'candidate' | 'retained'
  resourceId: string
  fileCount: number
  bytes: number
}
export interface PresentationProjectInventory {
  version: 1
  scope: 'known_presentation_namespaces'
  complete: true
  deletionPerformed: false
  resources: PresentationInventoryResource[]
  totals: { fileCount: number; bytes: number }
}
const hash = (value: string) => createHash('sha256').update(value).digest('hex')
function fail(code = 'presentation_inventory_invalid'): never {
  throw Error(code)
}
const identity = (a: Stats, b: Stats) => a.dev === b.dev && a.ino === b.ino && a.mode === b.mode
/** Read-only bounded namespace preview. A candidate still requires lifecycle freeze and mutation checks before deletion. */
export async function inspectPresentationProjectInventory(options: {
  userDataPath: string
  documentId: string
  projectId: string
  signal?: AbortSignal
}): Promise<PresentationProjectInventory> {
  try {
    const { documentId, projectId, signal } = options
    const check = () => {
      if (signal?.aborted) throw Error('cancelled')
    }
    check()
    if (
      typeof options.userDataPath !== 'string' ||
      !options.userDataPath ||
      typeof documentId !== 'string' ||
      !documentId.trim() ||
      documentId.length > 2048 ||
      typeof projectId !== 'string' ||
      !/^[A-Za-z0-9_-]{1,128}$/.test(projectId)
    )
      fail('invalid_request')
    const base = resolve(options.userDataPath),
      resources: PresentationInventoryResource[] = [],
      total = { fileCount: 0, bytes: 0 }
    const stat = (path: string) => lstatSync(path, { throwIfNoEntry: false })
    const chain = (path: string) => {
      if (path !== base && !path.startsWith(base + sep)) fail()
      const paths: string[] = []
      for (let p = path; ; p = dirname(p)) {
        paths.push(p)
        if (dirname(p) === p) break
      }
      return paths.reverse().flatMap((p) => {
        const s = stat(p)
        if (!s) return []
        if (s.isSymbolicLink() || (p !== path && !s.isDirectory())) fail()
        return [{ path: p, info: s }]
      })
    }
    const unchanged = (proof: ReturnType<typeof chain>) => {
      for (const entry of proof) {
        const now = stat(entry.path)
        if (!now || !identity(now, entry.info) || now.isSymbolicLink()) fail()
      }
    }
    const walk = (path: string, depth = 0): { path: string; info: Stats }[] => {
      check()
      if (depth > 16) fail('presentation_inventory_budget')
      const proof = chain(path),
        s = stat(path)
      if (!s) return []
      if (s.isSymbolicLink() || (!s.isFile() && !s.isDirectory())) fail()
      if (s.isFile()) {
        if (
          ++total.fileCount > MAX_PRESENTATION_INVENTORY_FILES ||
          !Number.isSafeInteger(s.size) ||
          (total.bytes += s.size) > MAX_PRESENTATION_INVENTORY_BYTES
        )
          fail('presentation_inventory_budget')
        unchanged(proof)
        return [{ path, info: s }]
      }
      const names = readdirSync(path).sort()
      if (names.length > MAX_PRESENTATION_INVENTORY_FILES) fail('presentation_inventory_budget')
      const files = names.flatMap((name) => walk(join(path, name), depth + 1))
      unchanged(proof)
      return files
    }
    const json = (path: string, limit = PROJECT_JSON_BYTES): any => {
      check()
      const proof = chain(path),
        s = stat(path)
      if (!s || !s.isFile() || s.size > limit) fail()
      let fd: number | undefined
      try {
        fd = openSync(path, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK)
        const opened = fstatSync(fd)
        if (!identity(opened, s) || !opened.isFile() || opened.size > limit) fail()
        unchanged(proof)
        const raw = readFileSync(fd)
        if (raw.length !== s.size || raw.length > limit) fail()
        unchanged(proof)
        const after = fstatSync(fd)
        if (
          !identity(after, s) ||
          after.size !== s.size ||
          after.mtimeMs !== s.mtimeMs ||
          after.ctimeMs !== s.ctimeMs
        )
          fail()
        return JSON.parse(raw.toString('utf8'))
      } catch {
        fail()
      } finally {
        if (fd !== undefined) closeSync(fd)
      }
    }
    const bound = (v: any) => {
      if (
        !v ||
        typeof v !== 'object' ||
        Array.isArray(v) ||
        v.documentId !== documentId ||
        v.projectId !== projectId
      )
        fail()
    }
    const add = (
      kind: PresentationInventoryKind,
      path: string,
      ownership: PresentationInventoryResource['ownership'],
      files: ReturnType<typeof walk>,
    ) => {
      if (!files.length) return
      resources.push({
        kind,
        ownership,
        disposition: ownership === 'project_exclusive' ? 'candidate' : 'retained',
        resourceId:
          'resource_' +
          hash(
            JSON.stringify([
              'presentation-inventory-v1',
              documentId,
              projectId,
              kind,
              relative(base, path),
            ]),
          ),
        fileCount: files.length,
        bytes: files.reduce((n, f) => n + f.info.size, 0),
      })
    }
    const project = join(base, 'projects', 'presentations', hash(projectId)),
      main = walk(project)
    let anchor = false
    if (main.length) {
      const v = json(join(project, 'project.json'))
      bound(v)
      if (v.version !== 1 || Object.keys(v).sort().join(',') !== 'documentId,projectId,version')
        fail()
      anchor = true
      for (const f of main.filter((f) => f.path.endsWith('.json'))) {
        const v = json(f.path)
        if (v?.documentId !== undefined || v?.projectId !== undefined) bound(v)
      }
      add('project', project, 'project_exclusive', main)
    }
    const exclusive = anchor ? 'project_exclusive' : 'unproven'
    for (const [kind, namespace] of [
      ['research', 'presentation-research'],
      ['delivery_bundles', 'presentation-delivery-bundles'],
      ['page_backups', 'presentation-page-backups'],
    ] as const) {
      const path =
          kind === 'page_backups'
            ? join(base, namespace, hash(projectId))
            : join(base, namespace, hash(documentId), hash(projectId)),
        files = walk(path)
      if (files.length) {
        if (kind === 'research') {
          const v = json(join(path, 'state.json'), RESEARCH_JSON_BYTES)
          bound(v.state)
          if (
            v.checksum !== hash(JSON.stringify(v.state)) ||
            ![1, 2].includes(v.state.version) ||
            !Array.isArray(v.state.records)
          )
            fail()
          if (
            !Number.isSafeInteger(v.state.revision) ||
            v.state.revision < 0 ||
            v.state.records.length !== v.state.totalRecords ||
            v.state.totalRecords > 128
          )
            fail()
          for (const record of v.state.records) bound(parsePresentationResearchRecord(record))
        } else {
          const metadata = files.filter((f) => f.path.endsWith(sep + 'metadata.json'))
          if (!metadata.length) fail()
          for (const f of metadata) {
            const v = json(f.path, 256 * 1024)
            if (kind === 'delivery_bundles') {
              if (v.checksum !== hash(JSON.stringify(v.receipt))) fail()
              const receipt = parsePresentationDeliveryBundleReceipt(v.receipt)
              bound(receipt)
            } else {
              bound(v)
              if (!['uploading', 'ready'].includes(v.status)) fail()
            }
          }
        }
        add(kind, path, exclusive, files)
      }
    }
    for (const [kind, namespace] of [
      ['preferences', 'presentation-preferences'],
      ['comments', 'presentation-comments'],
      ['manual_observations', 'presentation-manual-observations'],
    ] as const) {
      const path = join(base, namespace, hash(JSON.stringify([documentId, projectId])) + '.json'),
        files = walk(path)
      if (files.length) {
        const v = json(
          path,
          kind === 'preferences' ? 64 * 1024 : kind === 'comments' ? 320 * 1024 : 1024 * 1024,
        )
        bound(v)
        if (v.version !== 1) fail()
        if (kind === 'preferences') {
          if (!Array.isArray(v.preferences)) fail()
          for (const p of v.preferences) {
            if (parseSavedPresentationPreference(p).projectId !== projectId) fail()
          }
        }
        if (kind === 'comments' && !Array.isArray(v.comments)) fail()
        if (kind === 'manual_observations') {
          if (
            !Array.isArray(v.observations) ||
            v.checksum !== hash(canonicalPresentationValue(v.observations))
          )
            fail()
          for (const o of v.observations) bound(parsePresentationManualObservation(o))
        }
        add(kind, path, exclusive, files)
      }
    }
    for (const [kind, namespace] of [
      ['attachments', 'presentation-attachments'],
      ['existing_page_backups', 'presentation-existing-page-backups'],
      ['master_backups', 'presentation-master-backups'],
      ['package_backups', 'presentation-package-backups'],
    ] as const) {
      const path = join(base, namespace, hash(documentId))
      const files = walk(path)
      for (const file of files.filter((f) => f.path.endsWith('.json'))) {
        const value = json(file.path)
        if (value.documentId !== undefined && value.documentId !== documentId) fail()
      }
      add(kind, path, 'document_shared', files)
    }
    const releases = join(base, 'presentation-existing-page-backups', '.released', hash(documentId))
    add('existing_page_releases', releases, 'document_shared', walk(releases))
    const acquisition = join(base, 'presentation-acquisition-history', hash(documentId) + '.json'),
      acquisitionFiles = walk(acquisition)
    if (acquisitionFiles.length) {
      const value = json(acquisition, 256 * 1024)
      if (
        value.checksum !== hash(JSON.stringify(value.history)) ||
        parsePresentationAcquisitionHistory(value.history).documentId !== documentId
      )
        fail()
      add('acquisition_history', acquisition, 'document_shared', acquisitionFiles)
    }
    const teamRoot = join(base, 'presentation-teams')
    if (stat(teamRoot)) {
      const proof = chain(teamRoot)
      if (!stat(teamRoot)!.isDirectory()) fail()
      const names = readdirSync(teamRoot)
      if (names.length > MAX_PRESENTATION_INVENTORY_FILES) fail('presentation_inventory_budget')
      for (const name of names.sort()) {
        const path = join(teamRoot, name)
        if (!/^team_[a-f0-9]{64}\.json$/.test(name)) {
          add('teams', path, 'unproven', walk(path))
          continue
        }
        const v = parsePresentationTeamLedger(json(path, 768 * 1024))
        if (
          v.teamId + '.json' !== name ||
          'team_' + hash(JSON.stringify([v.ownerSubject, v.documentId, v.projectId])) !== v.teamId
        )
          fail()
        if (v.documentId === documentId && v.projectId === projectId)
          add('teams', path, 'unproven', walk(path))
      }
      unchanged(proof)
    }
    for (const [kind, namespace] of [
      ['unbound_page_staging', 'presentation-page-backups'],
      ['unbound_existing_staging', 'presentation-existing-page-backups'],
    ] as const) {
      const staging = join(base, namespace)
      if (stat(staging)) {
        const proof = chain(staging)
        const names = readdirSync(staging).sort()
        if (names.length > MAX_PRESENTATION_INVENTORY_FILES) fail('presentation_inventory_budget')
        for (const name of names)
          if (name.startsWith('.tmp-')) {
            const path = join(staging, name)
            add(kind, path, 'unproven', walk(path))
          }
        unchanged(proof)
      }
    }
    const brandIds = new Set<string>()
    for (const f of main.filter((f) => /plan(?:-revision-[0-9]+)?\.json$/.test(f.path))) {
      const v = json(f.path)
      if (v.plan?.brandKit?.id) {
        const id = v.plan.brandKit.id
        if (typeof id !== 'string' || !/^[A-Za-z0-9_-]{1,80}$/.test(id)) fail()
        brandIds.add(id)
      }
    }
    for (const id of [...brandIds].sort()) {
      const path = join(base, 'presentation-brand-kits', hash(id) + '.json'),
        files = walk(path)
      if (files.length) {
        const v = json(path, 256 * 1024)
        if (v.version !== 1 || v.id !== id || !Array.isArray(v.revisions)) fail()
        add('brand_reference', path, 'global_shared', files)
      }
    }
    const lifecycle = join(base, 'presentation-project-lifecycles', hash(projectId)),
      controls = walk(lifecycle)
    if (controls.length) {
      const v = json(join(lifecycle, 'lifecycle.json'))
      bound(v)
      if (v.version !== 1) fail()
      add('lifecycle_control', lifecycle, 'unproven', controls)
    }
    check()
    return {
      version: 1,
      scope: 'known_presentation_namespaces',
      complete: true,
      deletionPerformed: false,
      resources,
      totals: total,
    }
  } catch (error) {
    let code = ''
    try {
      if (error instanceof Error) code = error.message
    } catch {
      // An untrusted error getter cannot escape the finite public error boundary.
    }
    if (
      [
        'invalid_request',
        'cancelled',
        'presentation_inventory_invalid',
        'presentation_inventory_budget',
      ].includes(code)
    )
      // eslint-disable-next-line preserve-caught-error -- Public errors must not retain untrusted secret-bearing causes.
      throw Error(code)
    // eslint-disable-next-line preserve-caught-error -- Public errors must not retain untrusted secret-bearing causes.
    throw Error('presentation_inventory_invalid')
  }
}
