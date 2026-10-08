import { afterEach, it, expect } from 'vitest'
import {
  packageEditingFixture as fixture,
  cleanupPackageFixtures,
} from './helpers/package-editing-fixture.js'
afterEach(cleanupPackageFixtures)
it('does not upload PC backups for a rejected proposal', async () => {
  const f = await fixture()
  await f.propose()
  expect(f.request).not.toHaveBeenCalled()
  f.proposals.reject()
  expect(f.request).not.toHaveBeenCalled()
  expect(f.data.size).toBe(0)
})
it('releases an uploaded page if the next backup fails before intent persistence', async () => {
  const f = await fixture()
  const p = await f.propose()
  const originalRequest = f.request.getMockImplementation()!
  f.request.mockImplementation(async (body, signal) => {
    if (
      (body as { operation?: string; key?: string }).operation === 'package_backup_begin' &&
      (body as { key?: string }).key === 'page-1'
    )
      throw Error('transport_failed')
    return originalRequest(body, signal)
  })
  await expect(f.confirm()).rejects.toThrow('transport_failed')
  const response = await f.request({
    operation: 'package_backup_list',
    documentId: 'doc',
    changeId: String(p.preview.changeId),
  })
  expect((await response.json()).backups).toEqual([])
  expect(f.data.size).toBe(0)
  expect(f.adapter.stage).not.toHaveBeenCalled()
})
it('releases all backups when the first settings write fails before commit', async () => {
  const f = await fixture()
  const p = await f.propose()
  f.write.mockImplementationOnce(async () => {
    throw Error('settings_failed')
  })
  await expect(f.confirm()).rejects.toThrow('settings_failed')
  const response = await f.request({
    operation: 'package_backup_list',
    documentId: 'doc',
    changeId: String(p.preview.changeId),
  })
  expect((await response.json()).backups).toEqual([])
  expect(f.data.size).toBe(0)
  expect(f.adapter.stage).not.toHaveBeenCalled()
})
it('keeps savepoints when the first settings write commits but loses its ACK', async () => {
  const f = await fixture()
  const p = await f.propose()
  const write = f.write.getMockImplementation()!
  f.write.mockImplementationOnce(async (next, expected) => {
    await write(next, expected)
    throw Error('settings_ack_lost')
  })
  await expect(f.confirm()).rejects.toThrow('settings_ack_lost')
  const id = String(p.preview.changeId)
  expect(f.data.get(id)?.state).toBe('prepared')
  const response = await f.request({
    operation: 'package_backup_list',
    documentId: 'doc',
    changeId: id,
  })
  expect((await response.json()).backups).toHaveLength(4)
  expect(f.adapter.stage).not.toHaveBeenCalled()
})
it('persists PC originals, imports then deletes, and restores after reopen', async () => {
  const f = await fixture()
  const p = await f.propose()
  await f.confirm()
  const id = String(p.preview.changeId)
  expect(f.data.get(id)!.state).toBe('applied')
  f.reopen()
  await f.tool('undo', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('undone')
  expect(f.order).toEqual(['import-2', 's1'])
})
it.each(['stage_before', 'stage_after', 'callback_after'])(
  'retains %s import pending with no automatic delete/reimport',
  async (mode) => {
    const f = await fixture()
    const p = await f.propose()
    f.mode(mode)
    await expect(f.confirm()).rejects.toThrow()
    const id = String(p.preview.changeId)
    expect(f.data.get(id)!.pending?.action).toBe('import')
    expect(f.adapter.remove).not.toHaveBeenCalled()
    f.reopen()
    expect((await f.tool('resume', { change_id: id })).isError).toBe(true)
    f.mode('')
    await f.tool('reconcile', { change_id: id })
    await f.confirm()
    expect(f.data.get(id)!.state).toBe(mode === 'stage_before' ? 'discarded' : 'staged')
    expect(f.adapter.stage).toHaveBeenCalledTimes(1)
  },
)
it('closes lost deletion ACK only by explicit metadata reconciliation', async () => {
  const f = await fixture()
  const p = await f.propose()
  f.mode('remove_after')
  await expect(f.confirm()).rejects.toThrow()
  const id = String(p.preview.changeId)
  expect(f.data.get(id)!.pending?.action).toBe('delete_source')
  f.reopen()
  f.mode('')
  await f.tool('reconcile', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('applied')
  expect(f.adapter.remove).toHaveBeenCalledTimes(1)
})
it('refuses third party content drift before undo', async () => {
  const f = await fixture()
  const p = await f.propose()
  await f.confirm()
  f.packages.set('s1', f.packages.get('import-1')!)
  expect((await f.tool('undo', { change_id: String(p.preview.changeId) })).isError).toBe(true)
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
})
it('supports all600 pages with source at end index', async () => {
  const f = await fixture(600)
  const p = await f.propose('slide', 599)
  await f.confirm()
  expect(f.order).toHaveLength(600)
  expect(f.order.slice(0, 599)).toEqual(Array.from({ length: 599 }, (_, i) => `s${i}`))
  expect(f.data.get(String(p.preview.changeId))!.state).toBe('applied')
}, 120000)
it('preserves32 chart replacement paths', async () => {
  const f = await fixture()
  const p = await f.propose(
    'chart',
    0,
    Array.from({ length: 32 }, (_, i) => ({
      path: `ppt/charts/style${i + 1}.xml`,
      xml: '<c:style xmlns:c="urn:c" val="2"/>',
    })),
  )
  await f.confirm()
  expect(f.data.get(String(p.preview.changeId))!.state).toBe('applied')
})
it('discards prepared metadata after failed intent persistence with zero native writes', async () => {
  const f = await fixture()
  const p = await f.propose()
  f.writeMode('before_pending')
  await expect(f.confirm()).rejects.toThrow()
  const id = String(p.preview.changeId)
  expect(f.data.get(id)!.state).toBe('prepared')
  f.writeMode('')
  await f.tool('discard', { change_id: id })
  expect(f.proposals.pending()!.impact.host).toBe('local_checkpoint')
  expect(f.proposals.pending()!.operation).toBe(f.proposals.pending()!.toolName)
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('discarded')
  expect(f.adapter.stage).not.toHaveBeenCalled()
  expect(f.adapter.remove).not.toHaveBeenCalled()
})
it('explicitly discards a proven staged import while preserving source and other pages', async () => {
  const f = await fixture()
  const p = await f.propose()
  f.mode('callback_after')
  await expect(f.confirm()).rejects.toThrow()
  const id = String(p.preview.changeId)
  f.mode('')
  await f.tool('reconcile', { change_id: id })
  await f.confirm()
  await f.tool('discard', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('discarded')
  expect(f.order).toEqual(['s0', 's1'])
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
})
it('reconciles lost proof settings ACK without any native replay', async () => {
  const f = await fixture()
  const p = await f.propose()
  f.writeMode('observed_ack')
  await expect(f.confirm()).rejects.toThrow()
  const id = String(p.preview.changeId)
  expect(f.data.get(id)!.pending!.afterProofRef).toBeDefined()
  f.reopen()
  f.writeMode('')
  await f.tool('reconcile', { change_id: id })
  expect(f.proposals.pending()!.impact.host).toBe('local_checkpoint')
  expect(f.proposals.pending()!.operation).toBe(f.proposals.pending()!.toolName)
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('staged')
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
  expect(f.adapter.remove).not.toHaveBeenCalled()
  await f.tool('resume', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('applied')
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
})
it('reconciles unknown restoration ACK then deletes applied page without reimport', async () => {
  const f = await fixture()
  const p = await f.propose()
  await f.confirm()
  const id = String(p.preview.changeId)
  f.mode('restore_after')
  await f.tool('undo', { change_id: id })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.data.get(id)!.pending!.action).toBe('restore')
  f.reopen()
  f.mode('')
  await f.tool('reconcile', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('restore_staged')
  await f.tool('undo', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('undone')
  expect(f.adapter.stage).toHaveBeenCalledTimes(2)
})
it.each(['doc', 'disconnect', 'clear'])(
  'refuses %s change after an awaited live read',
  async (kind) => {
    const f = await fixture()
    await f.propose()
    f.setRead(() => {
      if (kind === 'doc') f.setDocument('different')
      else if (kind === 'disconnect') f.disconnect()
      else f.clear()
    })
    await expect(f.confirm()).rejects.toThrow()
    expect(f.adapter.stage).not.toHaveBeenCalled()
    expect(f.data.size).toBe(0)
  },
)
it('copies caller XML replacement aliases before awaited host reads', async () => {
  const f = await fixture()
  const replacements = structuredClone(f.replacements)
  f.setRead(() => {
    replacements[0]!.xml = 'mutated caller XML'
  })
  const p = await f.propose('slide', 0, replacements)
  f.setRead(() => {})
  await f.confirm()
  expect(f.data.get(String(p.preview.changeId))!.state).toBe('applied')
})
it('keeps no-op XML imports separated by actual native IDs', async () => {
  const f = await fixture()
  const zip = await (await import('jszip')).default.loadAsync(f.base64, { base64: true })
  const p = await f.propose('slide', 0, [
    {
      path: 'ppt/slides/slide1.xml',
      xml: await zip.file('ppt/slides/slide1.xml')!.async('string'),
    },
  ])
  await f.confirm()
  const saved = f.data.get(String(p.preview.changeId))!
  expect(saved.sourceSlideId).toBe('s0')
  expect(saved.packageSourceSlideId).toBe('256#')
  expect(saved.replacementSlideId).toBe('import-1')
})
it('keeps page screenshots/reviews historical and clears them on confirmed restoration', async () => {
  const f = await fixture()
  const p = await f.propose()
  await f.confirm()
  const id = String(p.preview.changeId),
    slideId = f.data.get(id)!.replacementSlideId!
  const captured = await f.tool('capture', { change_id: id, slide_id: slideId })
  const parsed = JSON.parse(captured.output)
  expect(parsed.qaPassed).toBe(false)
  const result = await f.tool('review', {
    change_id: id,
    slide_id: slideId,
    screenshot_digest: parsed.screenshotDigest,
    status: 'pass',
    notes: 'Observed source page',
  })
  expect(result.isError).toBeUndefined()
  expect(f.data.get(id)!.reviews).toHaveLength(1)
  await f.tool('undo', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.reviews).toHaveLength(0)
})
it('refuses source deletion when a third party changes a staged package after native insert', async () => {
  const f = await fixture()
  const p = await f.propose()
  f.mode('callback_after')
  await expect(f.confirm()).rejects.toThrow()
  const id = String(p.preview.changeId)
  f.packages.set('s1', f.packages.get('import-1')!)
  expect((await f.tool('reconcile', { change_id: id })).isError).toBe(true)
  expect(f.adapter.remove).not.toHaveBeenCalled()
})
it('refuses business document changes at last native synchronous guard', async () => {
  const f = await fixture()
  await f.propose()
  f.setBeforeSdk(() => f.setDocument('other'))
  await expect(f.confirm()).rejects.toThrow('presentation_document_changed')
  expect(f.order).toEqual(['s0', 's1'])
  expect(f.data.values().next().value!.pending!.action).toBe('import')
})
it('does not claim ambiguous added pages under an unresolved import', async () => {
  const f = await fixture()
  const p = await f.propose()
  f.mode('stage_after')
  await expect(f.confirm()).rejects.toThrow()
  const id = String(p.preview.changeId)
  f.order.splice(2, 0, 'foreign')
  f.packages.set('foreign', f.packages.get('import-1')!)
  expect((await f.tool('reconcile', { change_id: id })).isError).toBe(true)
  expect(f.adapter.remove).not.toHaveBeenCalled()
  expect(f.data.get(id)!.pending).toBeDefined()
})
it('closes a proven before-delete pending without writes and permits explicit staged discard', async () => {
  const f = await fixture()
  const p = await f.propose()
  f.mode('remove_before')
  await expect(f.confirm()).rejects.toThrow()
  const id = String(p.preview.changeId)
  f.mode('')
  await f.tool('reconcile', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('staged')
  expect(f.data.get(id)!.pending).toBeUndefined()
  expect(f.adapter.remove).toHaveBeenCalledTimes(1)
  await f.tool('discard', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('discarded')
  expect(f.order).toEqual(['s0', 's1'])
})
it('copies the entire review tool call before awaited screenshot reads', async () => {
  const f = await fixture()
  const p = await f.propose()
  await f.confirm()
  const id = String(p.preview.changeId),
    slideId = f.data.get(id)!.replacementSlideId!
  const capture = JSON.parse((await f.tool('capture', { change_id: id, slide_id: slideId })).output)
  const input = {
    change_id: id,
    slide_id: slideId,
    screenshot_digest: capture.screenshotDigest,
    status: 'pass',
    notes: 'Original notes',
  }
  const original = f.adapter.screenshotSlide.getMockImplementation()!
  f.adapter.screenshotSlide.mockImplementationOnce(async (i) => {
    input.status = 'fail'
    input.notes = 'Aliased notes'
    return original(i)
  })
  expect((await f.tool('review', input)).isError).toBeUndefined()
  const saved = f.data.get(id)!
  const { readPackageBackup } =
    await import('../src/skills/powerpoint/presentation-package-backup.js')
  const bytes = await readPackageBackup({
    request: f.request,
    documentId: saved.documentId,
    changeId: id,
    backup: saved.reviews[0]!.reviewRef,
  })
  expect(JSON.parse(new TextDecoder().decode(bytes))).toMatchObject({
    status: 'pass',
    notes: 'Original notes',
  })
})
