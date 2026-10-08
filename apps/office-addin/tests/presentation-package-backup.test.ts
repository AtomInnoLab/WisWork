import { describe, expect, it } from 'vitest'
import {
  readPackageBackup,
  savePackageBackup,
} from '../src/skills/powerpoint/presentation-package-backup.js'
const encode = (bytes: Uint8Array) =>
  btoa(Array.from(bytes, (x) => String.fromCharCode(x)).join(''))
function fixture(transform: (body: Record<string, unknown>, op: string) => unknown = (x) => x) {
  const calls: Record<string, unknown>[] = []
  let bytes = new Uint8Array(0)
  let meta: Record<string, unknown> = {}
  const request = async (raw: unknown) => {
    const body = raw as Record<string, unknown>
    calls.push(body)
    const op = String(body.operation)
    if (op === 'package_backup_begin') {
      bytes = new Uint8Array(Number(body.sizeBytes))
      meta = {
        documentId: body.documentId,
        changeId: body.changeId,
        key: body.key,
        sha256: body.sha256,
        sizeBytes: body.sizeBytes,
        status: 'uploading',
        receivedBytes: 0,
      }
    }
    if (op === 'package_backup_chunk') {
      const part = Uint8Array.from(atob(String(body.base64)), (x) => x.charCodeAt(0))
      bytes.set(part, Number(body.offset))
      meta.receivedBytes = Number(body.offset) + part.length
    }
    if (op === 'package_backup_finish') meta.status = 'ready'
    const result =
      op === 'package_backup_read'
        ? {
            documentId: meta.documentId,
            changeId: meta.changeId,
            key: meta.key,
            sha256: meta.sha256,
            sizeBytes: meta.sizeBytes,
            offset: body.offset,
            base64: encode(
              bytes.slice(Number(body.offset), Number(body.offset) + Number(body.length)),
            ),
          }
        : { ...meta }
    return new Response(JSON.stringify(transform(result, op)))
  }
  return { request, calls }
}
const scope = { documentId: 'document', changeId: 'change_1', key: 'snapshot' }
describe('master backup client', () => {
  it('chunks uploads and rereads before returning a verified receipt', async () => {
    const f = fixture()
    const bytes = Uint8Array.from({ length: 300000 }, (_, i) => i % 251)
    const backup = await savePackageBackup({ ...scope, request: f.request, bytes })
    expect(backup.sizeBytes).toBe(bytes.length)
    expect(
      f.calls.filter((x) => x.operation === 'package_backup_chunk').map((x) => x.offset),
    ).toEqual([0, 131072, 262144])
    expect(f.calls.filter((x) => x.operation === 'package_backup_read')).toHaveLength(3)
    expect(await readPackageBackup({ ...scope, request: f.request, backup })).toEqual(bytes)
  })
  it('snapshots caller bytes and scope before the first await', async () => {
    const f = fixture()
    const input = { ...scope, request: f.request, bytes: new Uint8Array([1, 2, 3]) }
    const pending = savePackageBackup(input)
    input.bytes.fill(9)
    input.documentId = 'other'
    input.key = 'page-2'
    const backup = await pending
    expect(await readPackageBackup({ ...scope, request: f.request, backup })).toEqual(
      new Uint8Array([1, 2, 3]),
    )
    expect(f.calls.every((x) => x.documentId === 'document' && x.key === 'snapshot')).toBe(true)
  })
  it.each(['documentId', 'changeId', 'key', 'sha256', 'sizeBytes'])(
    'rejects mismatched %s on read',
    async (field) => {
      let corrupt = false
      const f = fixture((body, op) =>
        corrupt && op === 'package_backup_read' ? { ...body, [field]: 'wrong' } : body,
      )
      const backup = await savePackageBackup({
        ...scope,
        request: f.request,
        bytes: new Uint8Array([1]),
      })
      corrupt = true
      await expect(readPackageBackup({ ...scope, request: f.request, backup })).rejects.toThrow(
        'presentation_package_backup_invalid',
      )
    },
  )
  it('rejects damaged bytes with correct metadata', async () => {
    const f = fixture((body, op) =>
      op === 'package_backup_read' ? { ...body, base64: 'Ag==' } : body,
    )
    await expect(
      savePackageBackup({ ...scope, request: f.request, bytes: new Uint8Array([1]) }),
    ).rejects.toThrow('presentation_package_backup_invalid')
  })
  it.each([
    { key: '../snapshot' },
    { key: 'page-' + '1'.repeat(124) },
    { changeId: '../change' },
    { documentId: '' },
    { documentId: 'x'.repeat(2049) },
  ])('rejects invalid scope without requests', async (invalid) => {
    const f = fixture()
    await expect(
      savePackageBackup({ ...scope, ...invalid, request: f.request, bytes: new Uint8Array([1]) }),
    ).rejects.toThrow('presentation_package_backup_invalid')
    expect(f.calls).toHaveLength(0)
  })
  it('rejects oversized and empty blobs', async () => {
    const f = fixture()
    for (const bytes of [new Uint8Array(0), new Uint8Array(8 * 1024 * 1024 + 1)])
      await expect(savePackageBackup({ ...scope, request: f.request, bytes })).rejects.toThrow(
        'presentation_package_backup_invalid',
      )
    expect(f.calls).toHaveLength(0)
  })
  it('rejects extra metadata, stalled progress and oversized responses', async () => {
    for (const transform of [
      (body: Record<string, unknown>) => ({ ...body, unexpected: true }),
      (body: Record<string, unknown>, op: string) =>
        op === 'package_backup_chunk' ? { ...body, receivedBytes: 0 } : body,
      () => ({ padding: 'x'.repeat(256 * 1024) }),
    ]) {
      const f = fixture(transform)
      await expect(
        savePackageBackup({ ...scope, request: f.request, bytes: new Uint8Array([1]) }),
      ).rejects.toThrow('presentation_package_backup_invalid')
    }
  })
  it('maps quota failures and preserves cancellation', async () => {
    await expect(
      savePackageBackup({
        ...scope,
        bytes: new Uint8Array([1]),
        request: async () =>
          new Response(JSON.stringify({ error: 'quota_exceeded' }), { status: 413 }),
      }),
    ).rejects.toThrow('presentation_package_backup_capacity')
    const controller = new AbortController()
    const f = fixture()
    const request = async (body: unknown) => {
      const response = await f.request(body)
      controller.abort()
      return response
    }
    await expect(
      savePackageBackup({
        ...scope,
        bytes: new Uint8Array([1]),
        request,
        signal: controller.signal,
      }),
    ).rejects.toThrow('cancelled')
    expect(f.calls).toHaveLength(1)
  })
  it('preserves cancellation when the response stream rejects during an abort', async () => {
    const controller = new AbortController()
    const request = async () =>
      new Response(
        new ReadableStream({
          async pull(stream) {
            await new Promise((resolve) => setTimeout(resolve, 0))
            controller.abort()
            stream.error(new DOMException('stream aborted', 'AbortError'))
          },
        }),
      )
    await expect(
      savePackageBackup({
        ...scope,
        bytes: new Uint8Array([1]),
        request,
        signal: controller.signal,
      }),
    ).rejects.toThrow('cancelled')
  })
})
