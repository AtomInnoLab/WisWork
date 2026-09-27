import { createHash, randomUUID } from 'node:crypto'
import { lstatSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

export interface SavedPresentationPreference {
  projectId: string
  changeId: string
  text: string
}
type RecordFile = { version: 1; documentId: string; projectId: string; preferences: SavedPresentationPreference[] }
const present = (path: string) => lstatSync(path, { throwIfNoEntry: false })
const id = (value: unknown, max: number) => typeof value === 'string' && new RegExp(`^[A-Za-z0-9_-]{1,${max}}$`).test(value)
const valid = (value: unknown): value is SavedPresentationPreference => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const p = value as SavedPresentationPreference
  return Object.keys(p).sort().join(',') === 'changeId,projectId,text' &&
    id(p.projectId, 80) && id(p.changeId, 128) &&
    typeof p.text === 'string' && p.text.trim().length > 0 &&
    p.text.length <= 240 && !Array.from(p.text).some((char) => {
      const code = char.charCodeAt(0)
      return code < 32 || (code >= 127 && code <= 159)
    })
}

/** Separate, document-scoped observations; never reads or writes the brand library. */
export class PresentationPreferenceLibrary {
  private root: string
  constructor(userDataPath: string) { this.root = join(userDataPath, 'presentation-preferences') }
  private path(documentId: string, projectId: string, create: boolean): string {
    if (typeof documentId !== 'string' || !documentId || documentId.length > 2048 || !id(projectId, 80))
      throw new Error('invalid_request')
    const stat = present(this.root)
    if (stat && (!stat.isDirectory() || stat.isSymbolicLink())) throw new Error('invalid_state')
    if (!stat && create) mkdirSync(this.root, { recursive: true, mode: 0o700 })
    return join(this.root, `${createHash('sha256').update(JSON.stringify([documentId, projectId])).digest('hex')}.json`)
  }
  private read(documentId: string, projectId: string): RecordFile | undefined {
    const path = this.path(documentId, projectId, false)
    const stat = present(path)
    if (!stat) return undefined
    if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 64 * 1024) throw new Error('invalid_state')
    try {
      const record = JSON.parse(readFileSync(path, 'utf8')) as RecordFile
      if (!record || Object.keys(record).sort().join(',') !== 'documentId,preferences,projectId,version' ||
        record.version !== 1 || record.documentId !== documentId || record.projectId !== projectId ||
        !Array.isArray(record.preferences) || record.preferences.length > 64 ||
        record.preferences.some((p) => !valid(p) || p.projectId !== projectId) ||
        new Set(record.preferences.map((p) => p.changeId)).size !== record.preferences.length)
        throw new Error('invalid_state')
      return record
    } catch { throw new Error('invalid_state') }
  }
  list(documentId: string, projectId: string): SavedPresentationPreference[] {
    return structuredClone(this.read(documentId, projectId)?.preferences ?? [])
  }
  save(documentId: string, input: unknown): SavedPresentationPreference {
    if (!valid(input)) throw new Error('invalid_request')
    const path = this.path(documentId, input.projectId, true)
    const prior = this.read(documentId, input.projectId)
    const existing = prior?.preferences.find((p) => p.changeId === input.changeId)
    if (existing) {
      if (existing.text !== input.text) throw new Error('revision_conflict')
      return structuredClone(existing)
    }
    if ((prior?.preferences.length ?? 0) >= 64) throw new Error('quota_exceeded')
    const record: RecordFile = { version: 1, documentId, projectId: input.projectId, preferences: [...(prior?.preferences ?? []), input] }
    const serialized = JSON.stringify(record)
    if (Buffer.byteLength(serialized, 'utf8') > 64 * 1024) throw new Error('quota_exceeded')
    const temporary = `${path}.${randomUUID()}.tmp`
    try {
      writeFileSync(temporary, serialized, { encoding: 'utf8', flag: 'wx', mode: 0o600 })
      renameSync(temporary, path)
    } finally { rmSync(temporary, { force: true }) }
    return structuredClone(input)
  }
  delete(documentId: string, projectId: string, changeId: string): boolean {
    if (!id(changeId, 128)) throw new Error('invalid_request')
    const prior = this.read(documentId, projectId)
    const preferences = prior?.preferences.filter((p) => p.changeId !== changeId) ?? []
    if (!prior || preferences.length === prior.preferences.length) return false
    const path = this.path(documentId, projectId, true)
    const temporary = `${path}.${randomUUID()}.tmp`
    try {
      writeFileSync(temporary, JSON.stringify({ ...prior, preferences }), { encoding: 'utf8', flag: 'wx', mode: 0o600 })
      renameSync(temporary, path)
    } finally { rmSync(temporary, { force: true }) }
    return true
  }
}
