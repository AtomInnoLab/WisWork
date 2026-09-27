import { createHash, randomUUID } from 'node:crypto'
import { lstatSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import type { PresentationPlan } from '@wiswork/pptx-engine/presentation-plan'

type TargetKind = 'slide' | 'claim' | 'source'
export interface PresentationComment {
  id: string
  targetKind: TargetKind
  targetId: string
  authorLabel: string
  text: string
  planRevision: number
  state: 'open' | 'resolved'
  createdAt: string
  updatedAt: string
}
export interface PresentationCommentLedger {
  version: 1
  documentId: string
  projectId: string
  revision: number
  comments: PresentationComment[]
}
const present = (path: string) => lstatSync(path, { throwIfNoEntry: false })
const id = (value: unknown) => typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const safeText = (value: unknown, max: number) => typeof value === 'string' && value.trim().length > 0 &&
  value.length <= max && !Array.from(value).some((char) => {
    const code = char.charCodeAt(0)
    return code < 32 || (code >= 127 && code <= 159)
  })
const time = (value: unknown) => typeof value === 'string' && Number.isFinite(Date.parse(value)) &&
  new Date(value).toISOString() === value
function validComment(value: unknown): value is PresentationComment {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const c = value as PresentationComment
  return Object.keys(c).sort().join(',') === 'authorLabel,createdAt,id,planRevision,state,targetId,targetKind,text,updatedAt' &&
    id(c.id) && id(c.targetId) && ['slide', 'claim', 'source'].includes(c.targetKind) &&
    safeText(c.authorLabel, 80) && safeText(c.text, 2000) &&
    Number.isSafeInteger(c.planRevision) && c.planRevision >= 1 &&
    ['open', 'resolved'].includes(c.state) && time(c.createdAt) && time(c.updatedAt) && c.updatedAt >= c.createdAt
}

/** Local review annotations stay bound to their plan revision and never count as QA approval. */
export class PresentationCommentLibrary {
  private root: string
  constructor(userDataPath: string) { this.root = join(userDataPath, 'presentation-comments') }
  private path(documentId: string, projectId: string, create: boolean): string {
    if (typeof documentId !== 'string' || !documentId || documentId.length > 2048 || !id(projectId))
      throw new Error('invalid_request')
    const stat = present(this.root)
    if (stat && (!stat.isDirectory() || stat.isSymbolicLink())) throw new Error('invalid_state')
    if (!stat && create) mkdirSync(this.root, { recursive: true, mode: 0o700 })
    return join(this.root, `${createHash('sha256').update(JSON.stringify([documentId, projectId])).digest('hex')}.json`)
  }
  list(documentId: string, projectId: string): PresentationCommentLedger {
    const path = this.path(documentId, projectId, false)
    const stat = present(path)
    if (!stat) return { version: 1, documentId, projectId, revision: 0, comments: [] }
    if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 320 * 1024) throw new Error('invalid_state')
    try {
      const ledger = JSON.parse(readFileSync(path, 'utf8')) as PresentationCommentLedger
      if (!ledger || Object.keys(ledger).sort().join(',') !== 'comments,documentId,projectId,revision,version' ||
        ledger.version !== 1 || ledger.documentId !== documentId || ledger.projectId !== projectId ||
        !Number.isSafeInteger(ledger.revision) || ledger.revision < 1 || ledger.revision > 256 ||
        !Array.isArray(ledger.comments) || ledger.comments.length < 1 || ledger.comments.length > 128 ||
        ledger.comments.some((comment) => !validComment(comment)) ||
        new Set(ledger.comments.map((comment) => comment.id)).size !== ledger.comments.length)
        throw new Error('invalid_state')
      return structuredClone(ledger)
    } catch { throw new Error('invalid_state') }
  }
  private write(ledger: PresentationCommentLedger): PresentationCommentLedger {
    const path = this.path(ledger.documentId, ledger.projectId, true)
    const serialized = JSON.stringify(ledger)
    if (Buffer.byteLength(serialized, 'utf8') > 320 * 1024) throw new Error('quota_exceeded')
    const temporary = `${path}.${randomUUID()}.tmp`
    try {
      writeFileSync(temporary, serialized, { encoding: 'utf8', flag: 'wx', mode: 0o600 })
      renameSync(temporary, path)
    } finally { rmSync(temporary, { force: true }) }
    return structuredClone(ledger)
  }
  add(documentId: string, projectId: string, expectedRevision: number, planRevision: number,
    input: unknown, current: { revision: number; plan: PresentationPlan }): PresentationCommentLedger {
    if (!Number.isSafeInteger(expectedRevision) || expectedRevision < 0 ||
      !Number.isSafeInteger(planRevision) || planRevision !== current.revision ||
      !input || typeof input !== 'object' || Array.isArray(input)) throw new Error('invalid_request')
    const comment = input as Pick<PresentationComment, 'id' | 'targetKind' | 'targetId' | 'authorLabel' | 'text'>
    if (Object.keys(comment).sort().join(',') !== 'authorLabel,id,targetId,targetKind,text' ||
      !id(comment.id) || !id(comment.targetId) ||
      !['slide', 'claim', 'source'].includes(comment.targetKind) ||
      !safeText(comment.authorLabel, 80) || !safeText(comment.text, 2000) ||
      !(comment.targetKind === 'slide' ? current.plan.slides :
        comment.targetKind === 'claim' ? current.plan.claims : current.plan.sources)
        .some((item) => item.id === comment.targetId)) throw new Error('invalid_request')
    const ledger = this.list(documentId, projectId)
    if (ledger.revision !== expectedRevision) throw new Error('revision_conflict')
    if (ledger.comments.some((item) => item.id === comment.id)) throw new Error('revision_conflict')
    if (ledger.comments.length >= 128 || ledger.revision >= 256) throw new Error('quota_exceeded')
    const now = new Date().toISOString()
    return this.write({ ...ledger, revision: ledger.revision + 1,
      comments: [...ledger.comments, { ...comment, planRevision, state: 'open', createdAt: now, updatedAt: now }] })
  }
  resolve(documentId: string, projectId: string, expectedRevision: number, commentId: string): PresentationCommentLedger {
    if (!Number.isSafeInteger(expectedRevision) || expectedRevision < 0 || !id(commentId)) throw new Error('invalid_request')
    const ledger = this.list(documentId, projectId)
    if (ledger.revision !== expectedRevision) throw new Error('revision_conflict')
    const target = ledger.comments.find((comment) => comment.id === commentId)
    if (!target) throw new Error('not_found')
    if (target.state !== 'open') throw new Error('invalid_state')
    if (ledger.revision >= 256) throw new Error('quota_exceeded')
    return this.write({ ...ledger, revision: ledger.revision + 1,
      comments: ledger.comments.map((comment) => comment.id === commentId
        ? { ...comment, state: 'resolved' as const, updatedAt: new Date().toISOString() } : comment) })
  }
}
