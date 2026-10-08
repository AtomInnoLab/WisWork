import { createHash } from 'node:crypto'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import { mkdtempSync, rmSync, readFileSync, writeFileSync, readdirSync, symlinkSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import { createPresentationService } from '../src/main/presentation-service.js'

it('persists a before snapshot and completes after reopening the actual service', async () => {
  const root = mkdtempSync(join(tmpdir(), 'manual-observation-'))
  try {
    let service = createPresentationService({ userDataPath: root })
    const request = async (body: object) =>
      JSON.parse(Buffer.from(await service(body, new AbortController().signal)).toString())
    const scope = { documentId: 'doc', projectId: 'project', observationId: 'observation' }
    const shape = {
      id: 'shape',
      name: 'Title',
      type: 'TextBox',
      left: 1,
      top: 2,
      width: 300,
      height: 40,
      text: 'before',
    }
    const begun = await request({
      operation: 'manual_observation_begin',
      ...scope,
      slideId: 'slide',
      shape,
    })
    expect(begun.observation?.before.shape).toEqual(shape)
    service = createPresentationService({ userDataPath: root })
    const complete = await request({
      operation: 'manual_observation_complete',
      ...scope,
      expectedBeforeDigest: begun.observation.before.digest,
      shape: { ...shape, text: 'after' },
    })
    expect(complete.observation.after.shape.text).toBe('after')
    expect(complete.observation.source).toBe('host_difference_unattributed')
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})

it('enforces immutable retries, strict scope and approved provenance while observation deletion is independent', async () => {
  const root = mkdtempSync(join(tmpdir(), 'manual-provenance-'))
  try {
    let service = createPresentationService({ userDataPath: root })
    const request = async (body: object) =>
      JSON.parse(Buffer.from(await service(body, new AbortController().signal)).toString())
    const scope = { documentId: 'doc', projectId: 'p', observationId: 'one' },
      shape = {
        id: 'shape',
        name: 'Title',
        type: 'TextBox',
        left: 0,
        top: 0,
        width: 100,
        height: 30,
        text: 'before',
      }
    const { observation } = await request({
      operation: 'manual_observation_begin',
      ...scope,
      slideId: 'slide',
      shape,
    })
    expect(
      (
        await request({
          operation: 'preference_save_observation',
          ...scope,
          expectedBeforeDigest: observation.before.digest,
          expectedAfterDigest: null,
          text: 'short titles',
        })
      ).error,
    ).toBe('invalid_request')
    const body = {
      operation: 'manual_observation_complete',
      ...scope,
      expectedBeforeDigest: observation.before.digest,
      shape: { ...shape, text: 'after' },
    }
    const completed = await request(body)
    expect(completed.observation.after.shape.text).toBe('after')
    expect(await request(body)).toEqual(completed)
    expect((await request({ ...body, shape: { ...shape, text: 'different' } })).error).toBe(
      'revision_conflict',
    )
    expect((await request({ ...body, expectedBeforeDigest: 'a'.repeat(64) })).error).toBe(
      'revision_conflict',
    )
    expect((await request({ ...body, documentId: 'other' })).error).toBe('not_found')
    expect((await request({ ...body, author: 'user' })).error).toBe('invalid_request')
    const saved = await request({
      operation: 'preference_save_observation',
      ...scope,
      expectedBeforeDigest: observation.before.digest,
      expectedAfterDigest: completed.observation.after.digest,
      text: 'short titles',
    })
    expect(saved.preference.origin).toEqual({
      version: 1,
      observationId: 'one',
      beforeDigest: observation.before.digest,
      afterDigest: completed.observation.after.digest,
    })
    expect(
      (
        await request({
          operation: 'preference_save',
          documentId: 'doc',
          preference: saved.preference,
        })
      ).error,
    ).toBe('invalid_request')
    expect(
      await request({
        operation: 'preference_save_observation',
        ...scope,
        expectedBeforeDigest: observation.before.digest,
        expectedAfterDigest: completed.observation.after.digest,
        text: 'short titles',
      }),
    ).toEqual(saved)
    expect(
      (
        await request({
          operation: 'preference_save_observation',
          ...scope,
          expectedBeforeDigest: observation.before.digest,
          expectedAfterDigest: completed.observation.after.digest,
          text: 'different',
        })
      ).error,
    ).toBe('revision_conflict')
    const imported = await request({
      operation: 'preference_import',
      documentId: 'target',
      projectId: 'target',
      source: { documentId: 'doc', projectId: 'p', changeId: 'manual_one' },
      expectedTextDigest: (await import('node:crypto'))
        .createHash('sha256')
        .update('short titles')
        .digest('hex'),
      approvalId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
      expectedOrigin: saved.preference.origin,
    })
    expect(imported.preference.origin).toEqual(saved.preference.origin)
    expect(
      await request({
        operation: 'manual_observation_delete',
        ...scope,
        expectedBeforeDigest: observation.before.digest,
        expectedAfterDigest: completed.observation.after.digest,
      }),
    ).toEqual({
      deleted: true,
    })
    service = createPresentationService({ userDataPath: root })
    expect(
      await request({ operation: 'manual_observation_list', documentId: 'doc', projectId: 'p' }),
    ).toEqual({ observations: [] })
    expect(
      await request({
        operation: 'preference_get',
        documentId: 'doc',
        projectId: 'p',
        changeId: 'manual_one',
      }),
    ).toEqual(saved)
    expect(
      (
        await request({
          operation: 'preference_get',
          documentId: 'target',
          projectId: 'target',
          changeId: imported.preference.changeId,
        })
      ).preference,
    ).toEqual(imported.preference)
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
it('requires a changed observation before approval and the exact origin precondition when importing it', async () => {
  const root = mkdtempSync(join(tmpdir(), 'manual-origin-'))
  try {
    const service = createPresentationService({ userDataPath: root }),
      request = async (body: object) =>
        JSON.parse(Buffer.from(await service(body, new AbortController().signal)).toString())
    const scope = { documentId: 'doc', projectId: 'p', observationId: 'one' },
      shape = {
        id: 'shape',
        name: 'title',
        type: 'TextBox',
        left: 0,
        top: 0,
        width: 100,
        height: 30,
        text: 'before',
      }
    const { observation } = await request({
      operation: 'manual_observation_begin',
      ...scope,
      slideId: 'slide',
      shape,
    })
    await request({
      operation: 'manual_observation_complete',
      ...scope,
      expectedBeforeDigest: observation.before.digest,
      shape,
    })
    expect(
      await request({
        operation: 'preference_save_observation',
        ...scope,
        expectedBeforeDigest: observation.before.digest,
        expectedAfterDigest: observation.before.digest,
        text: 'short',
      }),
    ).toEqual({ error: 'invalid_request' })
    const scope2 = { ...scope, observationId: 'two' },
      begun = await request({
        operation: 'manual_observation_begin',
        ...scope2,
        slideId: 'slide',
        shape,
      })
    const completed2 = await request({
      operation: 'manual_observation_complete',
      ...scope2,
      expectedBeforeDigest: begun.observation.before.digest,
      shape: { ...shape, text: 'after' },
    })
    const { preference } = await request({
      operation: 'preference_save_observation',
      ...scope2,
      expectedBeforeDigest: begun.observation.before.digest,
      expectedAfterDigest: completed2.observation.after.digest,
      text: 'short',
    })
    const body = {
      operation: 'preference_import',
      documentId: 'target',
      projectId: 'target',
      source: { documentId: 'doc', projectId: 'p', changeId: 'manual_two' },
      expectedTextDigest: (await import('node:crypto'))
        .createHash('sha256')
        .update('short')
        .digest('hex'),
      approvalId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
    }
    expect((await request(body)).error).toBe('revision_conflict')
    expect(
      (
        await request({
          ...body,
          expectedOrigin: { ...preference.origin, beforeDigest: 'f'.repeat(64) },
        })
      ).error,
    ).toBe('revision_conflict')
    expect(
      (await request({ ...body, expectedOrigin: { ...preference.origin, author: 'user' } })).error,
    ).toBe('invalid_request')
    expect(
      (await request({ ...body, expectedOrigin: preference.origin })).preference.origin,
    ).toEqual(preference.origin)
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
it('rejects modified snapshot bytes, broken files and symbolic links without returning fake empty history', async () => {
  for (const mode of ['digest', 'checksum', 'json', 'symlink', 'rootlink']) {
    const root = mkdtempSync(join(tmpdir(), 'manual-corrupt-'))
    try {
      const service = createPresentationService({ userDataPath: root }),
        request = async (body: object) =>
          JSON.parse(Buffer.from(await service(body, new AbortController().signal)).toString())
      const scope = { documentId: 'doc', projectId: 'p', observationId: 'one' },
        shape = {
          id: 'shape',
          name: 'title',
          type: 'TextBox',
          left: 0,
          top: 0,
          width: 100,
          height: 30,
          text: 'before',
        }
      expect(
        (
          await request({
            operation: 'manual_observation_begin',
            ...scope,
            slideId: 'slide',
            shape,
          })
        ).observation,
      ).toBeDefined()
      const folder = join(root, 'presentation-manual-observations'),
        file = join(folder, readdirSync(folder)[0]!)
      if (mode === 'rootlink') {
        const moved = join(root, 'elsewhere')
        ;(await import('node:fs')).renameSync(folder, moved)
        symlinkSync(moved, folder, 'dir')
      } else if (mode === 'symlink') {
        const elsewhere = join(root, 'elsewhere.json')
        writeFileSync(elsewhere, readFileSync(file))
        rmSync(file)
        symlinkSync(elsewhere, file)
      } else if (mode === 'json') writeFileSync(file, '{broken')
      else {
        const value = JSON.parse(readFileSync(file, 'utf8'))
        value.observations[0].before.shape.text = 'forged'
        if (mode === 'digest')
          value.checksum = createHash('sha256')
            .update(canonicalPresentationValue(value.observations))
            .digest('hex')
        writeFileSync(file, JSON.stringify(value))
      }
      expect(
        await request({ operation: 'manual_observation_list', documentId: 'doc', projectId: 'p' }),
      ).toEqual({ error: 'invalid_state' })
    } finally {
      rmSync(root, { recursive: true, force: true })
    }
  }
})
it('enforces 32 retained records, per-record UTF8 bytes and total bytes without evicting history', async () => {
  const root = mkdtempSync(join(tmpdir(), 'manual-quota-'))
  try {
    const service = createPresentationService({ userDataPath: root }),
      request = async (body: object) =>
        JSON.parse(Buffer.from(await service(body, new AbortController().signal)).toString())
    const scope = { documentId: 'doc', projectId: 'p' },
      shape = {
        id: 'shape',
        name: 'title',
        type: 'TextBox',
        left: 0,
        top: 0,
        width: 100,
        height: 30,
        text: 'before',
      }
    for (let i = 0; i < 32; i++)
      expect(
        (
          await request({
            operation: 'manual_observation_begin',
            ...scope,
            observationId: 'item' + i,
            slideId: 'slide',
            shape,
          })
        ).observation,
      ).toBeDefined()
    expect(
      (
        await request({
          operation: 'manual_observation_begin',
          ...scope,
          observationId: 'overflow',
          slideId: 'slide',
          shape,
        })
      ).error,
    ).toBe('quota_exceeded')
    expect(
      (await request({ operation: 'manual_observation_list', ...scope })).observations,
    ).toHaveLength(32)
    expect(
      (
        await request({
          operation: 'manual_observation_begin',
          ...scope,
          projectId: 'oversized',
          observationId: 'one',
          slideId: 'slide',
          shape: { ...shape, text: '汉'.repeat(23000) },
        })
      ).error,
    ).toBe('invalid_request')
    let accepted = 0
    for (let i = 0; i < 32; i++) {
      const value = await request({
        operation: 'manual_observation_begin',
        ...scope,
        projectId: 'bytes',
        observationId: 'item' + i,
        slideId: 'slide',
        shape: { ...shape, text: '汉'.repeat(21000) },
      })
      if (value.error) {
        expect(value.error).toBe('quota_exceeded')
        break
      }
      accepted++
    }
    expect(accepted).toBeGreaterThan(0)
    expect(accepted).toBeLessThan(32)
    expect(
      (await request({ operation: 'manual_observation_list', ...scope, projectId: 'bytes' }))
        .observations,
    ).toHaveLength(accepted)
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
it('requires exact before/after preconditions before deletion and preference approval', async () => {
  const root = mkdtempSync(join(tmpdir(), 'manual-cas-'))
  try {
    const service = createPresentationService({ userDataPath: root }),
      request = async (body: object) =>
        JSON.parse(Buffer.from(await service(body, new AbortController().signal)).toString())
    const scope = { documentId: 'doc', projectId: 'p', observationId: 'one' },
      shape = {
        id: 'shape',
        name: 'title',
        type: 'TextBox',
        left: 0,
        top: 0,
        width: 100,
        height: 30,
        text: 'before',
      }
    const { observation } = await request({
      operation: 'manual_observation_begin',
      ...scope,
      slideId: 'slide',
      shape,
    })
    expect((await request({ operation: 'manual_observation_delete', ...scope })).error).toBe(
      'invalid_request',
    )
    const expected = { expectedBeforeDigest: observation.before.digest, expectedAfterDigest: null }
    const completed = await request({
      operation: 'manual_observation_complete',
      ...scope,
      expectedBeforeDigest: expected.expectedBeforeDigest,
      shape: { ...shape, text: 'after' },
    })
    expect(
      (await request({ operation: 'manual_observation_delete', ...scope, ...expected })).error,
    ).toBe('revision_conflict')
    expect(
      (
        await request({
          operation: 'preference_save_observation',
          ...scope,
          ...expected,
          text: 'short',
        })
      ).error,
    ).toBe('revision_conflict')
    expect((await request({ operation: 'manual_observation_get', ...scope })).observation).toEqual(
      completed.observation,
    )
    expect(
      await request({
        operation: 'manual_observation_delete',
        ...scope,
        ...expected,
        expectedAfterDigest: completed.observation.after.digest,
      }),
    ).toEqual({ deleted: true })
    await request({
      operation: 'manual_observation_begin',
      ...scope,
      slideId: 'slide',
      shape: { ...shape, text: 'recreated' },
    })
    expect(
      (
        await request({
          operation: 'manual_observation_delete',
          ...scope,
          ...expected,
          expectedAfterDigest: null,
        })
      ).error,
    ).toBe('revision_conflict')
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
it('does not approve an old origin after same-text source provenance changes on disk', async () => {
  const root = mkdtempSync(join(tmpdir(), 'manual-origin-race-'))
  try {
    const service = createPresentationService({ userDataPath: root }),
      request = async (body: object) =>
        JSON.parse(Buffer.from(await service(body, new AbortController().signal)).toString())
    const scope = { documentId: 'doc', projectId: 'p', observationId: 'one' },
      shape = {
        id: 'shape',
        name: 'title',
        type: 'TextBox',
        left: 0,
        top: 0,
        width: 100,
        height: 30,
        text: 'before',
      }
    const { observation } = await request({
      operation: 'manual_observation_begin',
      ...scope,
      slideId: 'slide',
      shape,
    })
    const after = await request({
      operation: 'manual_observation_complete',
      ...scope,
      expectedBeforeDigest: observation.before.digest,
      shape: { ...shape, text: 'after' },
    })
    const { preference } = await request({
      operation: 'preference_save_observation',
      ...scope,
      expectedBeforeDigest: observation.before.digest,
      expectedAfterDigest: after.observation.after.digest,
      text: 'short',
    })
    const body = {
      operation: 'preference_import',
      documentId: 'target',
      projectId: 'target',
      source: { documentId: 'doc', projectId: 'p', changeId: 'manual_one' },
      expectedTextDigest: createHash('sha256').update('short').digest('hex'),
      expectedOrigin: preference.origin,
      approvalId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
    }
    const first = await request(body)
    expect(first.preference.origin).toEqual(preference.origin)
    const folder = join(root, 'presentation-preferences'),
      sourceFile = join(
        folder,
        readdirSync(folder).find(
          (name) => JSON.parse(readFileSync(join(folder, name), 'utf8')).documentId === 'doc',
        )!,
      )
    const source = JSON.parse(readFileSync(sourceFile, 'utf8'))
    source.preferences[0].origin.beforeDigest = 'f'.repeat(64)
    writeFileSync(sourceFile, JSON.stringify(source))
    expect(await request(body)).toEqual({ error: 'revision_conflict' })
    expect(
      (
        await request({
          operation: 'preference_get',
          documentId: 'target',
          projectId: 'target',
          changeId: first.preference.changeId,
        })
      ).preference,
    ).toEqual(first.preference)
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
