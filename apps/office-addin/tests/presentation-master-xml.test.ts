import { afterEach, expect, it } from 'vitest'
import { masterXmlFixture, cleanupMasterXmlFixtures } from './helpers/master-xml-fixture.js'
afterEach(cleanupMasterXmlFixtures)
it('backs up all originals and preserves proposed capability before writes', async () => {
  const f = await masterXmlFixture(25)
  const p = await f.propose()
  expect(p.toolName).toBe('edit_slide_master_xml')
  expect(p.preview.qaScope).toEqual({ basis: 'master_xml_savepoint', hostSlideIds: f.order })
  expect(f.adapter.stage).not.toHaveBeenCalled()
  expect(f.adapter.readPage).toHaveBeenCalledTimes(25)
  expect(f.request.mock.calls.some(([v]: any) => v.operation === 'package_backup_finish')).toBe(
    true,
  )
})
it('durably probes unused layouts and updates every dependent before deleting source', async () => {
  const f = await masterXmlFixture(25),
    p = await f.propose()
  await f.confirm()
  const r = f.data.get(String(p.preview.changeId))!
  expect(r.state).toBe('applied')
  expect(r.receiptCount).toBeGreaterThan(25)
  expect(f.order).toHaveLength(25)
  expect(f.adapter.applyLayout.mock.calls.filter(([v]) => v.slideId.startsWith('s'))).toHaveLength(
    24,
  )
  expect(r.inventoryCleanupVerified).toBe(false)
}, 120000)
it('restores all dependent content, actual associations and order after reopen', async () => {
  const f = await masterXmlFixture(3),
    p = await f.propose()
  await f.confirm()
  const id = String(p.preview.changeId)
  f.reopen()
  expect((await f.tool('undo', { change_id: id })).isError).not.toBe(true)
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('undone')
  expect(f.order).toHaveLength(3)
  expect(f.order.slice(1)).toEqual(['s1', 's2'])
  expect(f.deps.get('s1')!.masterId).toBe('m1')
}, 120000)
it.each(['stage_before', 'stage_after', 'callback_after'])(
  'requires explicit reconciliation for %s with no native retry',
  async (mode) => {
    const f = await masterXmlFixture(2),
      p = await f.propose()
    f.mode(mode)
    await expect(f.confirm()).rejects.toThrow()
    const id = String(p.preview.changeId)
    expect(f.data.get(id)!.pending?.action).toBe('original_probe_stage')
    expect(f.adapter.stage).toHaveBeenCalledTimes(1)
    f.mode('')
    f.reopen()
    expect((await f.tool('resume', { change_id: id })).isError).toBe(true)
    expect((await f.tool('reconcile', { change_id: id })).isError).not.toBe(true)
    await f.confirm()
    expect(f.adapter.stage).toHaveBeenCalledTimes(1)
    expect(f.data.get(id)!.state).toBe(mode === 'stage_before' ? 'discarded' : 'probing_original')
  },
  120000,
)
it('explicitly discards prepared intent metadata with no native operation', async () => {
  const f = await masterXmlFixture(),
    p = await f.propose()
  const spy = f.options.writeMasterXmlChange
  f.options.writeMasterXmlChange = async (next, before) => {
    if (next.pending) throw Error('settings_failed')
    return spy(next, before)
  }
  await expect(f.confirm()).rejects.toThrow()
  f.options.writeMasterXmlChange = spy
  expect((await f.tool('discard', { change_id: String(p.preview.changeId) })).isError).not.toBe(
    true,
  )
  expect(f.proposals.pending()!.impact.host).toBe('local_checkpoint')
  await f.confirm()
  expect(f.data.get(String(p.preview.changeId))!.state).toBe('discarded')
  expect(f.adapter.stage).not.toHaveBeenCalled()
})
it('keeps stress telemetry bounded without retaining full SDK and PC call payloads', async () => {
  const f = await masterXmlFixture(3),
    telemetry = f.useBoundedTelemetry(),
    p = await f.propose()
  await f.confirm()
  expect(telemetry.layoutSlideIds.filter((id) => id.startsWith('s'))).toHaveLength(2)
  f.reopen()
  expect((await f.tool('undo', { change_id: String(p.preview.changeId) })).isError).not.toBe(true)
  await f.confirm()
  expect(f.data.get(String(p.preview.changeId))!.state).toBe('undone')
  expect(telemetry.layoutSlideIds.filter((id) => id.startsWith('s'))).toHaveLength(4)
  expect(f.request.mock.calls).toHaveLength(0)
  for (const method of Object.values(f.adapter)) expect(method.mock.calls).toHaveLength(0)
}, 30000)

it('preserves all600 dependent pages in one forward and undo transaction', async () => {
  const f = await masterXmlFixture(600),
    telemetry = f.useBoundedTelemetry(),
    started = Date.now(),
    p = await f.propose()
  console.info('master XML600 prepared ms', Date.now() - started)
  await f.confirm()
  console.info('master XML600 applied ms', Date.now() - started)
  const id = String(p.preview.changeId),
    applied = f.data.get(id)!
  expect(applied.state).toBe('applied')
  expect(applied.scope.affectedPageCount).toBe(600)
  expect(telemetry.layoutSlideIds.filter((id) => id.startsWith('s'))).toHaveLength(599)
  expect(f.order).toHaveLength(600)
  expect(f.order.slice(1)).toEqual(Array.from({ length: 599 }, (_, i) => `s${i + 1}`))
  f.reopen()
  expect((await f.tool('undo', { change_id: id })).isError).not.toBe(true)
  await f.confirm()
  console.info('master XML600 undone ms', Date.now() - started)
  expect(f.data.get(id)!.state).toBe('undone')
  expect(f.order).toHaveLength(600)
  expect(f.order.slice(1)).toEqual(Array.from({ length: 599 }, (_, i) => `s${i + 1}`))
  expect(Array.from(f.deps.values()).every((d) => d.masterId === 'm1')).toBe(true)
}, 1800000)

it('records three reviews of the same page and retains its key sequence through undo', async () => {
  const f = await masterXmlFixture(2),
    p = await f.propose()
  await f.confirm()
  const id = String(p.preview.changeId),
    page = f.order[0]!,
    keys: string[] = []
  for (let i = 0; i < 3; i++) {
    const capture = await f.tool('capture', { change_id: id, slide_id: page })
    expect(capture.isError).not.toBe(true)
    const body = JSON.parse(capture.output)
    expect(
      (
        await f.tool('review', {
          change_id: id,
          slide_id: page,
          screenshot_digest: body.screenshotDigest,
          status: 'pass',
          notes: `review ${i}`,
        })
      ).isError,
    ).not.toBe(true)
    keys.push(f.data.get(id)!.reviews[0]!.reviewRef.key)
  }
  expect(new Set(keys).size).toBe(3)
  expect(f.data.get(id)!.reviewSequence).toBe(3)
  await f.tool('undo', { change_id: id })
  await f.confirm()
  const restored = f.order[0]!,
    capture = JSON.parse((await f.tool('capture', { change_id: id, slide_id: restored })).output)
  expect(
    (
      await f.tool('review', {
        change_id: id,
        slide_id: restored,
        screenshot_digest: capture.screenshotDigest,
        status: 'pass',
        notes: 'after undo',
      })
    ).isError,
  ).not.toBe(true)
  expect(f.data.get(id)!.reviewSequence).toBe(4)
  expect(keys).not.toContain(f.data.get(id)!.reviews[0]!.reviewRef.key)
}, 120000)

it('updates secondary masters and restores their actual dependencies after original inventory disappears', async () => {
  const f = await masterXmlFixture(4, { multipleMasters: true, disappearOriginalOnDelete: true }),
    p = await f.propose([
      {
        path: 'ppt/slideMasters/slideMaster1.xml',
        xml: f.originalMaster.replace('original', 'edited'),
      },
      { path: 'ppt/theme/theme2.xml', xml: '<a:theme xmlns:a="urn:a" name="secondary-edited"/>' },
    ])
  await f.confirm()
  const id = String(p.preview.changeId)
  expect(f.data.get(id)!.scope.affectedMasterCount).toBe(2)
  expect(f.data.get(id)!.scope.affectedPageCount).toBe(4)
  expect(f.masters.some((m) => m.masterId === 'm1')).toBe(false)
  expect(f.masters.some((m) => m.masterId === 'm2')).toBe(false)
  expect((await f.tool('undo', { change_id: id })).isError).not.toBe(true)
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('undone')
  expect(new Set(Array.from(f.deps.values()).map((v) => v.masterId)).size).toBe(2)
  expect(f.order).toHaveLength(4)
}, 120000)
it.each([
  'original_probe_stage',
  'original_probe',
  'original_probe_delete',
  'stage',
  'stage_master',
  'delete_stage_master',
  'restore_stage_master',
  'import_probe_stage',
  'import_probe',
  'import_probe_delete',
  'forward_page',
  'delete_source',
  'restore_stage',
  'restore_probe_stage',
  'restore_probe',
  'restore_probe_delete',
  'restore_page',
  'delete_applied_source',
])(
  'requires explicit metadata reconciliation after %s ACK loss without retry or inverse',
  async (phase) => {
    const f = await masterXmlFixture(2, {
        multipleMasters: ['stage_master', 'delete_stage_master', 'restore_stage_master'].includes(
          phase,
        ),
      }),
      p = await f.propose(),
      id = String(p.preview.changeId)
    let lost = false
    for (const key of ['stage', 'applyLayout', 'remove'] as const) {
      const original = f.adapter[key].getMockImplementation()!
      f.adapter[key].mockImplementation(async (...args: any[]) => {
        const result = await (original as any)(...args)
        if (!lost && f.data.get(id)?.pending?.action === phase) {
          lost = true
          throw Error('lost_ack')
        }
        return result
      })
    }
    if (phase.startsWith('restore') || phase === 'delete_applied_source') {
      await f.confirm()
      await f.tool('undo', { change_id: id })
    }
    await expect(f.confirm()).rejects.toThrow('lost_ack')
    expect(lost).toBe(true)
    expect(f.data.get(id)!.pending?.action).toBe(phase)
    const before =
      f.adapter.stage.mock.calls.length +
      f.adapter.applyLayout.mock.calls.length +
      f.adapter.remove.mock.calls.length
    f.reopen()
    expect((await f.tool('resume', { change_id: id })).isError).toBe(true)
    expect((await f.tool('reconcile', { change_id: id })).isError).not.toBe(true)
    expect(f.proposals.pending()!.impact.host).toBe('local_checkpoint')
    await f.confirm()
    expect(f.data.get(id)!.pending).toBeUndefined()
    expect(
      f.adapter.stage.mock.calls.length +
        f.adapter.applyLayout.mock.calls.length +
        f.adapter.remove.mock.calls.length,
    ).toBe(before)
  },
  120000,
)
it('preserves known after-proof when settings acknowledgement is lost and closes only explicitly', async () => {
  const f = await masterXmlFixture(),
    p = await f.propose(),
    id = String(p.preview.changeId),
    write = f.options.writeMasterXmlChange
  let lost = false
  f.options.writeMasterXmlChange = async (next, before) => {
    await write(next, before)
    if (next.pending?.afterProofRef && !lost) {
      lost = true
      throw Error('settings_ack_lost')
    }
  }
  await expect(f.confirm()).rejects.toThrow('settings_ack_lost')
  expect(f.data.get(id)!.pending?.afterProofRef).toBeDefined()
  f.options.writeMasterXmlChange = write
  f.reopen()
  await f.tool('reconcile', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.pending).toBeUndefined()
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
})

it.each(['document', 'capability', 'clear', 'cancel', 'CAS'])(
  'rejects %s change after proposal before the first native write',
  async (kind) => {
    const f = await masterXmlFixture(2),
      abort = new AbortController()
    await f.propose(undefined, abort.signal)
    if (kind === 'document') f.doc('save-as-other')
    if (kind === 'capability') f.available(false)
    if (kind === 'clear') f.clear()
    if (kind === 'cancel') abort.abort()
    if (kind === 'CAS') {
      const write = f.options.writeMasterXmlChange
      f.options.writeMasterXmlChange = (next, expected) =>
        write({ ...next, intent: 'changed elsewhere' }, expected)
    }
    await expect(f.confirm()).rejects.toThrow()
    expect(f.adapter.stage).not.toHaveBeenCalled()
    expect(f.adapter.applyLayout).not.toHaveBeenCalled()
    expect(f.adapter.remove).not.toHaveBeenCalled()
  },
  120000,
)
it('retains all32 replacement paths in the immutable program before any host write', async () => {
  const f = await masterXmlFixture(2)
  const JSZip = (await import('jszip')).default,
    zip = await JSZip.loadAsync(f.original, { base64: true })
  const replacements = []
  for (let i = 0; i < 32; i++) {
    const path = `ppt/theme/theme${i + 20}.xml`
    zip.file(path, `<a:theme xmlns:a="urn:a" name="original-${i}"/>`)
    replacements.push({ path, xml: `<a:theme xmlns:a="urn:a" name="edited-${i}"/>` })
  }
  const base64 = await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' })
  for (const id of f.order) f.packages.set(id, base64)
  const p = await f.propose(replacements)
  expect(p.toolName).toBe('edit_slide_master_xml')
  expect(f.adapter.stage).not.toHaveBeenCalled()
}, 120000)

it('uses each restored original package own numeric source identity', async () => {
  const f = await masterXmlFixture(3),
    JSZip = (await import('jszip')).default
  const zip = await JSZip.loadAsync(f.packages.get('s1')!, { base64: true })
  const path = 'ppt/presentation.xml',
    xml = await zip.file(path)!.async('string')
  zip.file(path, xml.replace('id="256"', 'id="513"'))
  f.packages.set('s1', await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }))
  const stage = f.adapter.stage.getMockImplementation()!,
    seen: string[] = []
  f.adapter.stage.mockImplementation(async (...args) => {
    const inv = await import('../src/skills/powerpoint/presentation-master-xml-package.js').then(
      (m) => m.inspectMasterXmlPackage(args[0].base64),
    )
    expect(args[0].packageSourceSlideId).toBe(inv.sourceSlideId)
    seen.push(inv.sourceSlideId)
    return stage(...args)
  })
  const apply = f.adapter.applyLayout.getMockImplementation()!
  let lost = false
  f.adapter.applyLayout.mockImplementation(async (...args) => {
    await apply(...args)
    if (args[0].slideId === 's1' && !lost) {
      lost = true
      const z = await JSZip.loadAsync(f.packages.get('s1')!, { base64: true })
      z.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"><p:cSld name="LOST"/></p:sld>')
      f.packages.set('s1', await z.generateAsync({ type: 'base64', compression: 'DEFLATE' }))
      throw Error('lost_ack')
    }
  })
  const p = await f.propose()
  await expect(f.confirm()).rejects.toThrow()
  const id = String(p.preview.changeId)
  f.reopen()
  expect((await f.tool('reconcile', { change_id: id })).isError).not.toBe(true)
  await f.confirm()
  expect((await f.tool('undo', { change_id: id })).isError).not.toBe(true)
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('undone')
  expect(seen).toContain('513#')
}, 120000)

it('propagates shared-theme XML changes to both independently staged native masters', async () => {
  const f = await masterXmlFixture(4, { multipleMasters: true }),
    JSZip = (await import('jszip')).default
  const shared = async (base64: string) => {
    const z = await JSZip.loadAsync(base64, { base64: true }),
      path = 'ppt/slideMasters/_rels/slideMaster2.xml.rels'
    z.file(
      path,
      (await z.file(path)!.async('string')).replace('../theme/theme2.xml', '../theme/theme1.xml'),
    )
    return z.generateAsync({ type: 'base64', compression: 'DEFLATE' })
  }
  for (const id of f.order) f.packages.set(id, await shared(f.packages.get(id)!))
  for (const m of f.masters) m.base64 = await shared(m.base64)
  const z = await JSZip.loadAsync(f.packages.get('s0')!, { base64: true }),
    theme = await z.file('ppt/theme/theme1.xml')!.async('string')
  const p = await f.propose([
    {
      path: 'ppt/theme/theme1.xml',
      xml: theme.replace('name="original-theme"', 'name="shared-edited"'),
    },
  ])
  await f.confirm()
  const id = String(p.preview.changeId)
  expect(f.data.get(id)!.scope.affectedMasterCount).toBe(2)
  expect(f.data.get(id)!.scope.affectedPageCount).toBe(4)
  expect(new Set(f.order.map((id) => f.deps.get(id)!.masterId)).size).toBe(2)
  expect((await f.tool('undo', { change_id: id })).isError).not.toBe(true)
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('undone')
}, 120000)
it('closes baseline original-probe uncertainty through explicit cleanup without replay', async () => {
  const f = await masterXmlFixture(2),
    p = await f.propose(),
    id = String(p.preview.changeId),
    apply = f.adapter.applyLayout.getMockImplementation()!
  let failed = false
  f.adapter.applyLayout.mockImplementation(async (...args) => {
    if (!failed) {
      failed = true
      throw Error('before_write_ack_loss')
    }
    return apply(...args)
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.data.get(id)!.pending?.action).toBe('original_probe')
  f.reopen()
  const count = f.adapter.applyLayout.mock.calls.length
  expect((await f.tool('reconcile', { change_id: id })).isError).not.toBe(true)
  await f.confirm()
  expect(f.adapter.applyLayout).toHaveBeenCalledTimes(count)
  expect((await f.tool('undo', { change_id: id })).isError).not.toBe(true)
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('discarded')
  expect(f.order).toEqual(['s0', 's1'])
}, 120000)
it('freezes review input before awaited document guards mutate caller aliases', async () => {
  const f = await masterXmlFixture(2),
    p = await f.propose()
  await f.confirm()
  const id = String(p.preview.changeId),
    page = f.order[0]!
  const capture = JSON.parse((await f.tool('capture', { change_id: id, slide_id: page })).output)
  const input = {
    change_id: id,
    slide_id: page,
    screenshot_digest: capture.screenshotDigest,
    status: 'pass',
    notes: 'accepted original notes',
  }
  f.readHook(() => {
    input.status = 'fail'
    input.notes = 'late caller alias'
    input.screenshot_digest = 'bad'
  })
  expect((await f.tool('review', input)).isError).not.toBe(true)
  expect(f.data.get(id)!.reviewSequence).toBe(1)
  f.readHook()
}, 120000)
it('returns a finite safe error for arbitrary transport text after reopening the durable change', async () => {
  const f = await masterXmlFixture(2),
    p = await f.propose()
  await f.confirm()
  f.reopen()
  f.request.mockRejectedValueOnce(new Error('private-transport-token=secret'))
  const result = await f.tool('inspect', { change_id: String(p.preview.changeId) })
  expect(result.isError).toBe(true)
  expect(JSON.parse(result.output)).toEqual({ error: 'presentation_master_xml_failed' })
  expect(result.output).not.toContain('secret')
})

it.each([
  'cancelled',
  'presentation_document_changed',
  'presentation_master_xml_stale',
  'presentation_package_backup_capacity',
])('preserves finite known recovery error %s without transport details', async (code) => {
  const f = await masterXmlFixture(2),
    p = await f.propose()
  await f.confirm()
  f.reopen()
  f.request.mockRejectedValueOnce(new Error(code))
  expect(
    JSON.parse((await f.tool('inspect', { change_id: String(p.preview.changeId) })).output),
  ).toEqual({ error: code })
})
it.each(['changing', 'throwing'])(
  'contains an actual Error.message %s accessor at the master tool boundary',
  async (mode) => {
    const f = await masterXmlFixture(2),
      p = await f.propose()
    await f.confirm()
    const error = new Error()
    let reads = 0
    Object.defineProperty(error, 'message', {
      get() {
        reads++
        if (mode === 'throwing') throw Error('private-accessor-token=secret')
        return reads === 1 ? 'cancelled' : 'private-accessor-token=secret'
      },
    })
    f.options.readMasterXmlChange = () => {
      throw error
    }
    const result = await f.tool('inspect', { change_id: String(p.preview.changeId) })
    expect(result.isError).toBe(true)
    expect(JSON.parse(result.output)).toEqual({
      error: mode === 'changing' ? 'cancelled' : 'presentation_master_xml_failed',
    })
    expect(reads).toBe(1)
    expect(result.output).not.toContain('secret')
  },
)
