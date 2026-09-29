export type PresentationLibraryWriteGuard = (
  scope: Readonly<{ documentId: string; projectId: string }>,
) => void

import type { PresentationManualObservation } from '@wiswork/pptx-engine/presentation-manual-observation'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import { createHash, randomUUID } from 'node:crypto'
import { lstatSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import {
  parsePresentationPreferenceSource,
  parsePresentationPreferenceOrigin,
  parseSavedPresentationPreference,
  type PresentationPreferenceSource,
  type SavedPresentationPreference,
} from '@wiswork/pptx-engine/presentation-preference'
export type { SavedPresentationPreference } from '@wiswork/pptx-engine/presentation-preference'
type RecordFile = {
  version: 1
  documentId: string
  projectId: string
  preferences: SavedPresentationPreference[]
}
const present = (path: string) => lstatSync(path, { throwIfNoEntry: false })
const id = (value: unknown, max: number) =>
  typeof value === 'string' && new RegExp(`^[A-Za-z0-9_-]{1,${max}}$`).test(value)
const hash = (text: string) => createHash('sha256').update(text).digest('hex')
const reuseId = (
  documentId: string,
  projectId: string,
  source: PresentationPreferenceSource,
  digest: string,
) =>
  'reuse_' +
  hash(
    JSON.stringify([
      documentId,
      projectId,
      source.documentId,
      source.projectId,
      source.changeId,
      digest,
    ]),
  )
const valid = (value: unknown): value is SavedPresentationPreference => {
  try {
    parseSavedPresentationPreference(value)
    return true
  } catch {
    return false
  }
}

/** Separate, document-scoped observations; never reads or writes the brand library. */
export class PresentationPreferenceLibrary {
  private root: string
  constructor(userDataPath: string) {
    this.root = join(userDataPath, 'presentation-preferences')
  }
  private path(documentId: string, projectId: string, create: boolean): string {
    if (
      typeof documentId !== 'string' ||
      !documentId ||
      documentId.length > 2048 ||
      !id(projectId, 80)
    )
      throw new Error('invalid_request')
    const stat = present(this.root)
    if (stat && (!stat.isDirectory() || stat.isSymbolicLink())) throw new Error('invalid_state')
    if (!stat && create) mkdirSync(this.root, { recursive: true, mode: 0o700 })
    return join(
      this.root,
      `${createHash('sha256')
        .update(JSON.stringify([documentId, projectId]))
        .digest('hex')}.json`,
    )
  }
  private read(documentId: string, projectId: string): RecordFile | undefined {
    const path = this.path(documentId, projectId, false)
    const stat = present(path)
    if (!stat) return undefined
    if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 64 * 1024)
      throw new Error('invalid_state')
    try {
      const record = JSON.parse(readFileSync(path, 'utf8')) as RecordFile
      if (
        !record ||
        Object.keys(record).sort().join(',') !== 'documentId,preferences,projectId,version' ||
        record.version !== 1 ||
        record.documentId !== documentId ||
        record.projectId !== projectId ||
        !Array.isArray(record.preferences) ||
        record.preferences.length > 64 ||
        record.preferences.some(
          (p) =>
            !valid(p) ||
            p.projectId !== projectId ||
            (p.reuse &&
              (p.reuse.sourceTextDigest !== hash(p.text) ||
                p.changeId !==
                  reuseId(documentId, projectId, p.reuse.source, p.reuse.sourceTextDigest) ||
                (p.reuse.source.documentId === documentId &&
                  p.reuse.source.projectId === projectId))),
        ) ||
        new Set(record.preferences.map((p) => p.changeId)).size !== record.preferences.length
      )
        throw new Error('invalid_state')
      return record
    } catch {
      throw new Error('invalid_state')
    }
  }
  list(documentId: string, projectId: string): SavedPresentationPreference[] {
    return structuredClone(this.read(documentId, projectId)?.preferences ?? [])
  }
  save(
    documentId: string,
    inputValue: unknown,
    assertWritable?: PresentationLibraryWriteGuard,
  ): SavedPresentationPreference {
    if (!valid(inputValue) || inputValue.reuse !== undefined || inputValue.origin !== undefined)
      throw new Error('invalid_request')
    const input = structuredClone(inputValue)
    this.path(documentId, input.projectId, false)
    const check = () => assertWritable?.(Object.freeze({ documentId, projectId: input.projectId }))
    check()
    const path = this.path(documentId, input.projectId, true)
    const prior = this.read(documentId, input.projectId)
    const existing = prior?.preferences.find((p) => p.changeId === input.changeId)
    if (existing) {
      if (
        existing.text !== input.text ||
        existing.reuse !== undefined ||
        existing.origin !== undefined
      )
        throw new Error('revision_conflict')
      return structuredClone(existing)
    }
    if ((prior?.preferences.length ?? 0) >= 64) throw new Error('quota_exceeded')
    const record: RecordFile = {
      version: 1,
      documentId,
      projectId: input.projectId,
      preferences: [...(prior?.preferences ?? []), input],
    }
    const serialized = JSON.stringify(record)
    if (Buffer.byteLength(serialized, 'utf8') > 64 * 1024) throw new Error('quota_exceeded')
    const temporary = `${path}.${randomUUID()}.tmp`
    try {
      check()
      writeFileSync(temporary, serialized, { encoding: 'utf8', flag: 'wx', mode: 0o600 })
      check()
      renameSync(temporary, path)
    } finally {
      rmSync(temporary, { force: true })
    }
    return structuredClone(input)
  }
  saveObservation(
    documentId: string,
    observation: PresentationManualObservation,
    text: unknown,
    assertWritable?: PresentationLibraryWriteGuard,
  ): SavedPresentationPreference {
    if (
      observation.documentId !== documentId ||
      !observation.after ||
      observation.before.digest === observation.after.digest
    )
      throw new Error('invalid_request')
    const input = parseSavedPresentationPreference({
      projectId: observation.projectId,
      changeId: 'manual_' + observation.observationId,
      text,
      origin: {
        version: 1,
        observationId: observation.observationId,
        beforeDigest: observation.before.digest,
        afterDigest: observation.after.digest,
      },
    })
    const prior = this.read(documentId, input.projectId),
      existing = prior?.preferences.find((p) => p.changeId === input.changeId)
    if (existing) {
      if (canonicalPresentationValue(existing) !== canonicalPresentationValue(input))
        throw new Error('revision_conflict')
      return structuredClone(existing)
    }
    if ((prior?.preferences.length ?? 0) >= 64) throw new Error('quota_exceeded')
    const record: RecordFile = {
      version: 1,
      documentId,
      projectId: input.projectId,
      preferences: [...(prior?.preferences ?? []), input],
    }
    const serialized = JSON.stringify(record)
    if (Buffer.byteLength(serialized) > 64 * 1024) throw new Error('quota_exceeded')
    this.path(documentId, input.projectId, false)
    const check = () => assertWritable?.(Object.freeze({ documentId, projectId: input.projectId }))
    check()
    const path = this.path(documentId, input.projectId, true),
      temporary = path + '.' + randomUUID() + '.tmp'
    try {
      check()
      writeFileSync(temporary, serialized, { flag: 'wx', mode: 0o600 })
      check()
      renameSync(temporary, path)
    } finally {
      rmSync(temporary, { force: true })
    }
    return structuredClone(input)
  }
  get(
    documentId: string,
    projectId: string,
    changeId: string,
  ): SavedPresentationPreference | undefined {
    if (!id(changeId, 128)) throw new Error('invalid_request')
    const found = this.read(documentId, projectId)?.preferences.find((p) => p.changeId === changeId)
    return found ? structuredClone(found) : undefined
  }
  import(
    documentId: string,
    projectId: string,
    sourceValue: unknown,
    expectedTextDigest: unknown,
    approvalId: unknown,
    expectedOrigin?: unknown,
    guards?: {
      assertWritable?: PresentationLibraryWriteGuard
      assertSourceCurrent?: PresentationLibraryWriteGuard
    },
  ): SavedPresentationPreference {
    const source = structuredClone(parsePresentationPreferenceSource(sourceValue))
    const assertWritable = guards?.assertWritable,
      assertSourceCurrent = guards?.assertSourceCurrent
    const sourceScope = Object.freeze({
      documentId: source.documentId,
      projectId: source.projectId,
    })
    const origin =
      expectedOrigin === undefined || expectedOrigin === null
        ? expectedOrigin
        : parsePresentationPreferenceOrigin(expectedOrigin)
    const check = () => {
      assertSourceCurrent?.(sourceScope)
      assertWritable?.(Object.freeze({ documentId, projectId }))
    }
    const path = this.path(documentId, projectId, false)
    if (source.documentId === documentId && source.projectId === projectId)
      throw new Error('invalid_request')
    if (typeof expectedTextDigest !== 'string' || !/^[a-f0-9]{64}$/.test(expectedTextDigest))
      throw new Error('invalid_request')
    assertSourceCurrent?.(sourceScope)
    const original = this.get(source.documentId, source.projectId, source.changeId)
    assertSourceCurrent?.(sourceScope)
    if (!original) throw new Error('not_found')
    if (original.reuse) throw new Error('invalid_request')
    if (
      (original.origin && origin === undefined) ||
      (origin !== undefined &&
        canonicalPresentationValue(origin) !== canonicalPresentationValue(original.origin ?? null))
    )
      throw new Error('revision_conflict')
    if (hash(original.text) !== expectedTextDigest) throw new Error('revision_conflict')
    const preference = parseSavedPresentationPreference({
      projectId,
      changeId: reuseId(documentId, projectId, source, expectedTextDigest),
      text: original.text,
      ...(original.origin ? { origin: original.origin } : {}),
      reuse: {
        version: 1,
        source,
        sourceTextDigest: expectedTextDigest,
        approvedAt: new Date().toISOString(),
        approvalId,
      },
    })
    const prior = this.read(documentId, projectId)
    const existing = prior?.preferences.find((p) => p.changeId === preference.changeId)
    if (existing) {
      if (
        !existing.reuse ||
        existing.text !== preference.text ||
        existing.reuse.source.documentId !== source.documentId ||
        existing.reuse.source.projectId !== source.projectId ||
        existing.reuse.source.changeId !== source.changeId ||
        existing.reuse.sourceTextDigest !== expectedTextDigest ||
        canonicalPresentationValue(existing.origin ?? null) !==
          canonicalPresentationValue(original.origin ?? null)
      )
        throw new Error('revision_conflict')
      return structuredClone(existing)
    }
    if ((prior?.preferences.length ?? 0) >= 64) throw new Error('quota_exceeded')
    const record: RecordFile = {
      version: 1,
      documentId,
      projectId,
      preferences: [...(prior?.preferences ?? []), preference],
    }
    const serialized = JSON.stringify(record)
    if (Buffer.byteLength(serialized, 'utf8') > 64 * 1024) throw new Error('quota_exceeded')
    check()
    this.path(documentId, projectId, true)
    const temporary = `${path}.${randomUUID()}.tmp`
    try {
      check()
      writeFileSync(temporary, serialized, { encoding: 'utf8', flag: 'wx', mode: 0o600 })
      check()
      renameSync(temporary, path)
    } finally {
      rmSync(temporary, { force: true })
    }
    return structuredClone(preference)
  }
  delete(
    documentId: string,
    projectId: string,
    changeId: string,
    assertWritable?: PresentationLibraryWriteGuard,
  ): boolean {
    if (!id(changeId, 128)) throw new Error('invalid_request')
    const prior = this.read(documentId, projectId)
    const preferences = prior?.preferences.filter((p) => p.changeId !== changeId) ?? []
    if (!prior || preferences.length === prior.preferences.length) return false
    const check = () => assertWritable?.(Object.freeze({ documentId, projectId }))
    check()
    const path = this.path(documentId, projectId, true)
    const temporary = `${path}.${randomUUID()}.tmp`
    try {
      check()
      writeFileSync(temporary, JSON.stringify({ ...prior, preferences }), {
        encoding: 'utf8',
        flag: 'wx',
        mode: 0o600,
      })
      check()
      renameSync(temporary, path)
    } finally {
      rmSync(temporary, { force: true })
    }
    return true
  }
}
