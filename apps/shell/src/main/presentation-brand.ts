import { createHash } from 'node:crypto'
import { randomUUID } from 'node:crypto'
import { lstatSync, mkdirSync, readFileSync, readdirSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { basename, join } from 'node:path'
import type { PresentationInlineAsset } from '@wiswork/pptx-engine/presentation'
import { parsePresentationBrandKit, type PresentationBrandKit, type PresentationPlan } from '@wiswork/pptx-engine/presentation-plan'

const present = (path: string) => lstatSync(path, { throwIfNoEntry: false })
type BrandRecord = { version: 1; id: string; revisions: PresentationBrandKit[] }

/** Local reusable kit catalog. Every revision is immutable and plans retain their own snapshot. */
export class PresentationBrandLibrary {
  private readonly root: string
  constructor(userDataPath: string) { this.root = join(userDataPath, 'presentation-brand-kits') }
  private directory(create: boolean): string | undefined {
    const stat = present(this.root)
    if (stat && (!stat.isDirectory() || stat.isSymbolicLink())) throw new Error('invalid_state')
    if (!stat && create) mkdirSync(this.root, { recursive: true, mode: 0o700 })
    return present(this.root) ? this.root : undefined
  }
  private path(id: string): string {
    if (!/^[A-Za-z0-9_-]{1,80}$/.test(id)) throw new Error('invalid_request')
    return join(this.root, `${createHash('sha256').update(id).digest('hex')}.json`)
  }
  private read(id: string): BrandRecord | undefined {
    const path = this.path(id)
    const stat = present(path)
    if (!stat) return undefined
    if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 256 * 1024) throw new Error('invalid_state')
    try {
      const record = JSON.parse(readFileSync(path, 'utf8')) as BrandRecord
      if (!record || record.version !== 1 || record.id !== id ||
        Object.keys(record).sort().join(',') !== 'id,revisions,version' ||
        !Array.isArray(record.revisions) || !record.revisions.length || record.revisions.length > 256)
        throw new Error('invalid_state')
      for (const [index, kit] of record.revisions.entries()) {
        parsePresentationBrandKit(kit)
        if (kit.id !== id || kit.revision !== index + 1) throw new Error('invalid_state')
      }
      return record
    } catch { throw new Error('invalid_state') }
  }
  get(id: string, revision?: number): PresentationBrandKit | undefined {
    this.directory(false)
    const record = this.read(id)
    if (!record) return undefined
    if (revision !== undefined && (!Number.isSafeInteger(revision) || revision < 1)) throw new Error('invalid_request')
    const kit = record.revisions[(revision ?? record.revisions.length) - 1]
    return kit ? structuredClone(kit) : undefined
  }
  list(): PresentationBrandKit[] {
    const root = this.directory(false)
    if (!root) return []
    const names = readdirSync(root).filter((name) => !/^[a-f0-9]{64}\.json\.[0-9a-f-]+\.tmp$/.test(name))
    if (names.length > 64 || names.some((name) => !/^[a-f0-9]{64}\.json$/.test(name))) throw new Error('invalid_state')
    return names.map((name) => {
      const stat = present(join(root, name))
      if (!stat?.isFile() || stat.isSymbolicLink() || stat.size > 256 * 1024) throw new Error('invalid_state')
      let parsed: BrandRecord
      try { parsed = JSON.parse(readFileSync(join(root, name), 'utf8')) as BrandRecord }
      catch { throw new Error('invalid_state') }
      if (typeof parsed?.id !== 'string') throw new Error('invalid_state')
      const latest = this.read(parsed.id)
      if (!latest || name !== `${createHash('sha256').update(parsed.id).digest('hex')}.json`) throw new Error('invalid_state')
      return structuredClone(latest.revisions.at(-1)!)
    }).sort((a, b) => a.id.localeCompare(b.id))
  }
  save(expectedRevision: number, input: unknown): PresentationBrandKit {
    const kit = parsePresentationBrandKit(input)
    if (!Number.isSafeInteger(expectedRevision) || expectedRevision < 0 || kit.revision !== expectedRevision + 1)
      throw new Error('invalid_request')
    const root = this.directory(true)!
    const prior = this.read(kit.id)
    if ((prior?.revisions.length ?? 0) !== expectedRevision) throw new Error('revision_conflict')
    if (prior && prior.revisions.length >= 256) throw new Error('quota_exceeded')
    if (!prior && this.list().length >= 64) throw new Error('quota_exceeded')
    const record: BrandRecord = { version: 1, id: kit.id, revisions: [...(prior?.revisions ?? []), kit] }
    const serialized = JSON.stringify(record)
    if (Buffer.byteLength(serialized, 'utf8') > 256 * 1024) throw new Error('quota_exceeded')
    const path = this.path(kit.id)
    const temporary = join(root, `${basename(path)}.${randomUUID()}.tmp`)
    try {
      writeFileSync(temporary, serialized, { encoding: 'utf8', flag: 'wx', mode: 0o600 })
      renameSync(temporary, path)
    } finally { rmSync(temporary, { force: true }) }
    return structuredClone(kit)
  }
}

/** Bind the brand logo to the bytes actually passed to the compiler, after attachment conversion. */
export function assertBrandLogoAsset(plan: PresentationPlan, assets: PresentationInlineAsset[]): void {
  const logo = plan.brandKit?.logo
  if (!logo) return
  const asset = assets.find((item) => item.id === logo.assetId)
  if (!asset || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(asset.base64))
    throw new Error('plan_mismatch')
  const bytes = Buffer.from(asset.base64, 'base64')
  if (!bytes.length || createHash('sha256').update(bytes).digest('hex') !== logo.assetDigest)
    throw new Error('plan_mismatch')
}
