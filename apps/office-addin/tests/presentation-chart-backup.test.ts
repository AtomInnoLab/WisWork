import { describe, expect, it } from 'vitest'
import {
  readChartPackageBackup,
  saveChartPackageBackup,
} from '../src/skills/powerpoint/presentation-chart-backup'

const scope = { documentId: 'doc', hostSlideId: 'slide', slideIds: ['slide'] }
const base64 = btoa('small chart package')

function transport(options: { tamper?: boolean; wrongScope?: boolean } = {}) {
  let received = new Uint8Array(0)
  let record: Record<string, unknown> = {}
  const request = async (body: unknown) => {
    const data = body as Record<string, unknown>
    const operation = data.operation
    if (operation === 'existing_page_backup_begin') {
      record = {
        backupId: data.backupId,
        documentId: data.documentId,
        hostSlideId: data.hostSlideId,
        slideIds: data.slideIds,
        sha256: data.sha256,
        sizeBytes: data.sizeBytes,
        receivedBytes: 0,
        status: 'pending',
      }
    } else if (operation === 'existing_page_backup_chunk') {
      const chunk = Uint8Array.from(atob(data.base64 as string), (x) => x.charCodeAt(0))
      const next = new Uint8Array(received.length + chunk.length)
      next.set(received)
      next.set(chunk, received.length)
      received = next
      record.receivedBytes = received.length
    } else if (operation === 'existing_page_backup_finish') {
      record.status = 'ready'
    } else if (operation === 'existing_page_backup_read') {
      const offset = data.offset as number
      const length = data.length as number
      const chunk = received.subarray(offset, offset + length)
      return new Response(
        JSON.stringify({
          backupId: record.backupId,
          offset,
          sizeBytes: record.sizeBytes,
          sha256: record.sha256,
          base64: btoa(
            String.fromCharCode(...(options.tamper ? Uint8Array.from(chunk, (x) => x ^ 1) : chunk)),
          ),
        }),
      )
    }
    return new Response(
      JSON.stringify(options.wrongScope ? { ...record, hostSlideId: 'other' } : record),
    )
  }
  return { request }
}

describe('chart package backup transport', () => {
  it('stores and reads the exact package bytes with scope and hash checks', async () => {
    const { request } = transport()
    const backup = await saveChartPackageBackup({
      ...scope,
      request,
      base64,
      backupId: 'chart_backup',
    })
    expect(backup.sizeBytes).toBe(19)
    expect(await readChartPackageBackup({ ...scope, request, backup })).toBe(base64)
  })
  it('rejects a mismatched host slide', async () => {
    const { request } = transport({ wrongScope: true })
    await expect(
      saveChartPackageBackup({ ...scope, request, base64, backupId: 'chart_backup' }),
    ).rejects.toThrow('presentation_chart_backup_invalid')
  })
  it('rejects tampered chunk bytes and wrong semantic digest', async () => {
    const { request } = transport({ tamper: true })
    const backup = await saveChartPackageBackup({
      ...scope,
      request,
      base64,
      backupId: 'chart_backup',
    })
    await expect(readChartPackageBackup({ ...scope, request, backup })).rejects.toThrow(
      'presentation_chart_backup_invalid',
    )
  })
  it('rejects malformed or oversized data before transport', async () => {
    const { request } = transport()
    await expect(
      saveChartPackageBackup({ ...scope, request, base64: 'bad', backupId: 'chart_backup' }),
    ).rejects.toThrow('presentation_chart_backup_invalid')
    await expect(
      saveChartPackageBackup({
        ...scope,
        request,
        base64: 'A'.repeat(12_000_000),
        backupId: 'chart_backup',
      }),
    ).rejects.toThrow('presentation_chart_backup_invalid')
  })
})

it.each(['quota_exceeded', 'quota_exceeded: /private/secret', 'access_denied'])(
  'classifies only exact backup-begin quota response %s',
  async (error) => {
    const request = async () =>
      new Response(JSON.stringify({ error, detail: '/private/secret' }), { status: 409 })
    await expect(
      saveChartPackageBackup({ ...scope, request, base64, backupId: 'chart_backup' }),
    ).rejects.toThrow(
      error === 'quota_exceeded'
        ? 'presentation_existing_backup_capacity'
        : 'presentation_chart_backup_invalid',
    )
  },
)
it('does not classify quota errors in later backup operations as capacity', async () => {
  const backend = transport()
  const request = async (body: unknown) =>
    (body as Record<string, unknown>).operation === 'existing_page_backup_chunk'
      ? new Response(JSON.stringify({ error: 'quota_exceeded' }), { status: 409 })
      : backend.request(body)
  await expect(
    saveChartPackageBackup({ ...scope, request, base64, backupId: 'chart_backup' }),
  ).rejects.toThrow('presentation_chart_backup_invalid')
})
