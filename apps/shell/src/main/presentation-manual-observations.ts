export type PresentationLibraryWriteGuard = (
  scope: Readonly<{ documentId: string; projectId: string }>,
) => void

import { createHash, randomUUID } from 'node:crypto'
import {
  constants,
  closeSync,
  fstatSync,
  lstatSync,
  mkdirSync,
  openSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { join } from 'node:path'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import {
  MAX_PRESENTATION_MANUAL_OBSERVATIONS,
  MAX_PRESENTATION_MANUAL_OBSERVATIONS_BYTES,
  parsePresentationManualObservation,
  parsePresentationManualObservationShape,
  type PresentationManualObservation,
} from '@wiswork/pptx-engine/presentation-manual-observation'
const hash = (value: unknown) =>
  createHash('sha256').update(canonicalPresentationValue(value)).digest('hex')
const present = (path: string) => lstatSync(path, { throwIfNoEntry: false })
/** Synchronous scoped transactions keep each immutable observation and its checksum atomic. */
export class PresentationManualObservationLibrary {
  private root: string
  constructor(userDataPath: string) {
    this.root = join(userDataPath, 'presentation-manual-observations')
  }
  private path(documentId: string, projectId: string, create = false) {
    if (
      typeof documentId !== 'string' ||
      !documentId ||
      documentId.length > 2048 ||
      typeof projectId !== 'string' ||
      !/^[A-Za-z0-9_-]{1,80}$/.test(projectId)
    )
      throw new Error('invalid_request')
    const stat = present(this.root)
    if (stat && (!stat.isDirectory() || stat.isSymbolicLink())) throw new Error('invalid_state')
    if (!stat && create) mkdirSync(this.root, { recursive: true, mode: 0o700 })
    return join(this.root, hash([documentId, projectId]) + '.json')
  }
  private id(value: string) {
    if (typeof value !== 'string' || !/^[A-Za-z0-9_-]{1,80}$/.test(value))
      throw new Error('invalid_request')
  }
  list(documentId: string, projectId: string): PresentationManualObservation[] {
    const path = this.path(documentId, projectId),
      stat = present(path)
    if (!stat) return []
    if (
      !stat.isFile() ||
      stat.isSymbolicLink() ||
      stat.size > MAX_PRESENTATION_MANUAL_OBSERVATIONS_BYTES
    )
      throw new Error('invalid_state')
    let fd: number | undefined
    try {
      fd = openSync(path, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0))
      const opened = fstatSync(fd)
      if (
        !opened.isFile() ||
        opened.ino !== stat.ino ||
        opened.dev !== stat.dev ||
        opened.size > MAX_PRESENTATION_MANUAL_OBSERVATIONS_BYTES
      )
        throw new Error('invalid_state')
      const v = JSON.parse(readFileSync(fd, 'utf8'))
      if (
        !v ||
        Object.keys(v).sort().join(',') !== 'checksum,documentId,observations,projectId,version' ||
        v.version !== 1 ||
        v.documentId !== documentId ||
        v.projectId !== projectId ||
        !Array.isArray(v.observations) ||
        v.observations.length > MAX_PRESENTATION_MANUAL_OBSERVATIONS ||
        v.checksum !== hash(v.observations)
      )
        throw new Error('invalid_state')
      const records = v.observations.map(
        parsePresentationManualObservation,
      ) as PresentationManualObservation[]
      if (
        new Set(records.map((r) => r.observationId)).size !== records.length ||
        records.some(
          (r) =>
            r.documentId !== documentId ||
            r.projectId !== projectId ||
            r.before.digest !== hash(r.before.shape) ||
            (r.after && r.after.digest !== hash(r.after.shape)),
        )
      )
        throw new Error('invalid_state')
      return records
    } catch {
      throw new Error('invalid_state')
    } finally {
      if (fd !== undefined) closeSync(fd)
    }
  }
  get(documentId: string, projectId: string, observationId: string) {
    this.id(observationId)
    return this.list(documentId, projectId).find((r) => r.observationId === observationId)
  }
  private write(
    documentId: string,
    projectId: string,
    observations: PresentationManualObservation[],
    assertWritable?: PresentationLibraryWriteGuard,
  ) {
    const serialized = JSON.stringify({
      version: 1,
      documentId,
      projectId,
      observations,
      checksum: hash(observations),
    })
    if (
      observations.length > MAX_PRESENTATION_MANUAL_OBSERVATIONS ||
      Buffer.byteLength(serialized) > MAX_PRESENTATION_MANUAL_OBSERVATIONS_BYTES
    )
      throw new Error('quota_exceeded')
    const check = () => assertWritable?.(Object.freeze({ documentId, projectId }))
    this.path(documentId, projectId, false)
    check()
    const path = this.path(documentId, projectId, true),
      temporary = path + '.' + randomUUID() + '.tmp'
    try {
      check()
      writeFileSync(temporary, serialized, { flag: 'wx', mode: 0o600 })
      check()
      renameSync(temporary, path)
    } finally {
      rmSync(temporary, { force: true })
    }
  }
  begin(
    documentId: string,
    projectId: string,
    observationId: string,
    slideId: string,
    shapeValue: unknown,
    assertWritable?: PresentationLibraryWriteGuard,
  ) {
    this.id(observationId)
    const shape = parsePresentationManualObservationShape(shapeValue),
      records = this.list(documentId, projectId),
      old = records.find((r) => r.observationId === observationId)
    if (old) {
      if (
        old.slideId !== slideId ||
        canonicalPresentationValue(old.before.shape) !== canonicalPresentationValue(shape)
      )
        throw new Error('revision_conflict')
      return old
    }
    const record = parsePresentationManualObservation({
      version: 1,
      source: 'host_difference_unattributed',
      observationId,
      documentId,
      projectId,
      slideId,
      shapeId: shape.id,
      before: { capturedAt: new Date().toISOString(), shape, digest: hash(shape) },
      atomicSnapshot: false,
      coverage: 'text_geometry_aggregate_font',
    })
    this.write(documentId, projectId, [...records, record], assertWritable)
    return record
  }
  complete(
    documentId: string,
    projectId: string,
    observationId: string,
    expectedBeforeDigest: unknown,
    shapeValue: unknown,
    assertWritable?: PresentationLibraryWriteGuard,
  ) {
    this.id(observationId)
    const records = this.list(documentId, projectId),
      old = records.find((r) => r.observationId === observationId)
    if (!old) throw new Error('not_found')
    if (expectedBeforeDigest !== old.before.digest) throw new Error('revision_conflict')
    const shape = parsePresentationManualObservationShape(shapeValue)
    if (old.after) {
      if (canonicalPresentationValue(old.after.shape) !== canonicalPresentationValue(shape))
        throw new Error('revision_conflict')
      return old
    }
    const record = parsePresentationManualObservation({
      ...old,
      after: {
        capturedAt: new Date(Math.max(Date.now(), Date.parse(old.before.capturedAt))).toISOString(),
        shape,
        digest: hash(shape),
      },
    })
    this.write(
      documentId,
      projectId,
      records.map((r) => (r.observationId === observationId ? record : r)),
      assertWritable,
    )
    return record
  }
  delete(
    documentId: string,
    projectId: string,
    observationId: string,
    assertWritable?: PresentationLibraryWriteGuard,
  ) {
    this.id(observationId)
    const records = this.list(documentId, projectId),
      next = records.filter((r) => r.observationId !== observationId)
    if (next.length === records.length) return false
    this.write(documentId, projectId, next, assertWritable)
    return true
  }
}
