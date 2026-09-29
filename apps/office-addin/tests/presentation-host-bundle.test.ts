import JSZip from 'jszip'
import { createHash } from 'node:crypto'
import { describe, expect, it, vi } from 'vitest'
import { deliveryReportFixture } from './presentation-delivery-fixture.js'
import { createPresentationHostBundleSkill } from '../src/skills/powerpoint/presentation-host-bundle.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex')
async function setup(vfs = new InMemoryVfs()) {
  const report = await deliveryReportFixture()
  const pptx = await new JSZip()
    .file('ppt/slides/slide1.xml', '<title>用户修改后的当前文稿</title>')
    .generateAsync({ type: 'uint8array' })
  let receipt: Record<string, unknown> | undefined
  let stored = new Uint8Array()
  let document = report.documentId
  const request = vi.fn(async (body: Record<string, unknown>) => {
    const op = body.operation
    if (op === 'status') return Response.json({ deliveryBundlesAvailable: true })
    if (op === 'production_delivery_report') return Response.json(report)
    if (op === 'delivery_bundle_begin') {
      receipt = {
        version: 1,
        documentId: document,
        projectId: report.projectId,
        requestId: report.requestId,
        bundleId: body.bundleId,
        sha256: body.sha256,
        sizeBytes: body.sizeBytes,
        receivedBytes: 0,
        state: 'uploading',
        createdAt: (body.manifest as { createdAt: string }).createdAt,
        manifest: body.manifest,
      }
    } else if (op === 'delivery_bundle_chunk') {
      const chunk = Buffer.from(body.base64 as string, 'base64')
      const next = new Uint8Array(stored.length + chunk.length)
      next.set(stored)
      next.set(chunk, stored.length)
      stored = next
      receipt!.receivedBytes = stored.length
    } else if (op === 'delivery_bundle_finish') {
      receipt!.state = 'ready'
      receipt!.completedAt = new Date().toISOString()
    } else if (op === 'delivery_bundle_read') {
      return Response.json({
        bundleId: receipt!.bundleId,
        offset: body.offset,
        totalBytes: stored.length,
        base64: Buffer.from(
          stored.slice(Number(body.offset), Number(body.offset) + Number(body.length)),
        ).toString('base64'),
      })
    }
    return Response.json(receipt)
  })
  const exportDocument = vi.fn(async (format: 'pptx' | 'pdf') =>
    format === 'pptx' ? pptx : new TextEncoder().encode('%PDF-1.7\ncurrent host pdf'),
  )
  const options = {
    available: () => true,
    nativeAvailable: () => true,
    request,
    documentId: async () => document,
    vfs,
    exportDocument,
    readCheckpoints: () => [],
  }
  const skill = createPresentationHostBundleSkill(options)
  const input = { project_id: report.projectId, request_id: report.requestId }
  const call = (name = 'export_current_presentation_bundle', extras = {}, signal?: AbortSignal) =>
    skill.executeTool({ id: 'x', name, input: { ...input, ...extras } }, signal)
  return {
    report,
    pptx,
    request,
    exportDocument,
    skill,
    call,
    vfs,
    options,
    bytes: () => stored,
    changeDocument: () => {
      document = 'other'
    },
  }
}
describe('current native host delivery package', () => {
  it('persists actual edited host bytes and every evidence file, then restores without exporting again', async () => {
    const f = await setup()
    const result = await f.call(undefined, { include_pdf: true })
    expect(result.isError).toBeFalsy()
    const value = JSON.parse(result.output)
    expect(value.receipt.state).toBe('ready')
    expect(value.bundleId).toBe(hash(f.bytes()))
    const zip = await JSZip.loadAsync(f.bytes())
    expect(Object.keys(zip.files).sort()).toEqual(
      [
        'README.md',
        'checkpoints.json',
        'claims.json',
        'evidence.json',
        'evidence.md',
        'manifest.json',
        'presentation.pdf',
        'presentation.pptx',
        'quality.json',
        'sources.json',
      ].sort(),
    )
    expect(await zip.file('presentation.pptx')!.async('uint8array')).toEqual(f.pptx)
    const manifest = JSON.parse(await zip.file('manifest.json')!.async('string'))
    expect(manifest.checks).toEqual({
      completion: 'not_verified',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
      roundTrip: 'not_run',
      hostQa: 'not_checked',
      pdf: 'included',
    })
    for (const file of manifest.files) {
      const bytes = await zip.file(file.name)!.async('uint8array')
      expect(hash(bytes)).toBe(file.sha256)
      expect(bytes.length).toBe(file.sizeBytes)
    }
    expect(JSON.parse(await zip.file('evidence.json')!.async('string'))).toEqual(f.report)
    expect(await zip.file('README.md')!.async('string')).toContain('整个当前')
    const reopenedVfs = new InMemoryVfs()
    const reopened = createPresentationHostBundleSkill({ ...f.options, vfs: reopenedVfs })
    const restored = await reopened.executeTool({
      id: 'restore',
      name: 'restore_presentation_delivery_bundle',
      input: {
        project_id: f.report.projectId,
        request_id: f.report.requestId,
        bundle_id: value.bundleId,
      },
    })
    expect(restored.isError).toBeFalsy()
    expect(f.exportDocument).toHaveBeenCalledTimes(2)
    expect(reopenedVfs.readBytes(JSON.parse(restored.output).paths[0])).toEqual(f.bytes())
  })
  it('hides old PC and rejects invalid input before native export', async () => {
    const f = await setup()
    f.request.mockResolvedValue(Response.json({ error: 'invalid_request' }))
    expect((await f.call()).output).toBe('presentation_upgrade_required')
    expect(f.exportDocument).not.toHaveBeenCalled()
    expect((await f.call(undefined, { extra: true })).output).toBe('invalid_tool_input')
  })
  it('retains PPTX when optional native PDF is unavailable and never substitutes compiled PDF', async () => {
    const f = await setup()
    f.exportDocument.mockImplementation(async (format) => {
      if (format === 'pdf') throw Error('office_document_export_unavailable')
      return f.pptx
    })
    const result = await f.call(undefined, { include_pdf: true })
    expect(result.isError).toBeFalsy()
    const zip = await JSZip.loadAsync(f.bytes())
    expect(zip.file('presentation.pdf')).toBeNull()
    expect(JSON.parse(await zip.file('manifest.json')!.async('string')).checks.pdf).toBe(
      'unavailable',
    )
  })
  it('rejects changed documents and clear during host export before publishing or upload', async () => {
    for (const change of [true, false]) {
      const f = await setup()
      f.exportDocument.mockImplementation(async () => {
        if (change) f.changeDocument()
        else f.skill.clear()
        return f.pptx
      })
      expect((await f.call()).output).toBe(change ? 'presentation_document_changed' : 'cancelled')
      expect(
        f.request.mock.calls.some(([body]) => body.operation === 'delivery_bundle_begin'),
      ).toBe(false)
      expect(f.vfs.list('/home/user')).toEqual([])
    }
  })
  it('does not replay native export after a lost finish response; the persisted package remains recoverable', async () => {
    const f = await setup()
    const actual = f.request.getMockImplementation()!
    f.request.mockImplementation(async (body) => {
      const response = await actual(body)
      if (body.operation === 'delivery_bundle_finish') throw Error('private network error')
      return response
    })
    expect((await f.call()).isError).toBe(true)
    expect(f.exportDocument).toHaveBeenCalledTimes(1)
    expect(f.vfs.list('/home/user')).toEqual([])
    const bundleId = hash(f.bytes())
    const restored = await f.call('restore_presentation_delivery_bundle', { bundle_id: bundleId })
    expect(restored.isError).toBeFalsy()
    expect(f.exportDocument).toHaveBeenCalledTimes(1)
  })
  it('keeps PC package when session storage is full and verifies read bytes before any VFS write', async () => {
    const f = await setup(new InMemoryVfs({ maxFiles: 0 }))
    expect((await f.call()).output).toBe('presentation_session_storage_full')
    const actual = f.request.getMockImplementation()!
    f.request.mockImplementation(async (body) =>
      body.operation === 'delivery_bundle_read'
        ? Response.json({
            bundleId: body.bundleId,
            offset: body.offset,
            totalBytes: f.bytes().length,
            base64: Buffer.alloc(Number(body.length)).toString('base64'),
          })
        : actual(body),
    )
    expect(
      (await f.call('restore_presentation_delivery_bundle', { bundle_id: hash(f.bytes()) })).output,
    ).toBe('presentation_response_invalid')
  })
})

it('packages only validated same-task historical QA and refuses wrong-document metadata', async () => {
  const f = await setup()
  const qa = {
    version: 1 as const,
    source: 'production' as const,
    documentId: f.report.documentId,
    projectId: f.report.projectId,
    requestId: f.report.requestId,
    artifactDigest: 'a'.repeat(64),
    pages: [],
  }
  const skill = createPresentationHostBundleSkill({ ...f.options, readQuality: () => qa })
  const result = await skill.executeTool({
    id: 'qa',
    name: 'export_current_presentation_bundle',
    input: { project_id: f.report.projectId, request_id: f.report.requestId },
  })
  expect(result.isError, result.output).toBeFalsy()
  const zip = await JSZip.loadAsync(f.bytes())
  expect(JSON.parse(await zip.file('manifest.json')!.async('string')).checks.hostQa).toBe(
    'historical_records_only',
  )
  expect(JSON.parse(await zip.file('quality.json')!.async('string'))).toMatchObject({
    record: qa,
    needsRecapture: true,
    checks: { completion: 'not_verified' },
  })
  const wrong = await setup()
  const wrongSkill = createPresentationHostBundleSkill({
    ...wrong.options,
    readQuality: () => ({ ...qa, documentId: 'another' }),
  })
  expect(
    (
      await wrongSkill.executeTool({
        id: 'bad',
        name: 'export_current_presentation_bundle',
        input: { project_id: wrong.report.projectId, request_id: wrong.report.requestId },
      })
    ).output,
  ).toBe('presentation_delivery_bundle_history_invalid')
  expect(
    wrong.request.mock.calls.some(([body]) => body.operation === 'delivery_bundle_begin'),
  ).toBe(false)
})
it('aborts optional PDF without converting cancellation into an unavailable warning', async () => {
  const f = await setup(),
    controller = new AbortController()
  f.exportDocument.mockImplementation(async (format) => {
    if (format === 'pdf') {
      controller.abort()
      throw Error('cancelled')
    }
    return f.pptx
  })
  expect((await f.call(undefined, { include_pdf: true }, controller.signal)).output).toBe(
    'cancelled',
  )
  expect(f.request.mock.calls.some(([body]) => body.operation === 'delivery_bundle_begin')).toBe(
    false,
  )
})
it('rejects invalid checkpoint history instead of silently dropping save points', async () => {
  const f = await setup()
  const skill = createPresentationHostBundleSkill({
    ...f.options,
    readCheckpoints: () => [{ kind: 'text', record: { documentId: 'other' } }],
  })
  expect(
    (
      await skill.executeTool({
        id: 'bad-history',
        name: 'export_current_presentation_bundle',
        input: { project_id: f.report.projectId, request_id: f.report.requestId },
      })
    ).output,
  ).toBe('presentation_delivery_bundle_history_invalid')
  expect(f.request.mock.calls.some(([body]) => body.operation === 'delivery_bundle_begin')).toBe(
    false,
  )
})
