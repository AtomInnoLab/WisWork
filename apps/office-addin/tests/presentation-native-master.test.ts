import { readMasterBackup } from '../src/skills/powerpoint/presentation-master-backup.js'
import { afterEach, expect, it } from 'vitest'
import type { PowerPointMasterOperation } from '../src/skills/powerpoint/browser-powerpoint-adapter.js'
import {
  nativeMasterFixture as fixture,
  cleanupNativeMasterFixtures,
} from './helpers/native-master-fixture.js'
afterEach(cleanupNativeMasterFixtures)
it('saves and verifies original PC packages before a native write and undoes after reopen', async () => {
  const f = await fixture()
  const p = await f.propose()
  await f.confirm()
  const id = String(p.preview.changeId),
    saved = f.data.get(id)!
  expect(saved.state).toBe('applied')
  expect(saved.receipts).toHaveLength(1)
  expect(
    f.request.mock.calls.some(([body]) => (body as any).operation === 'master_backup_read'),
  ).toBe(true)
  f.reopen()
  await f.tool('undo_slide_master_change', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('undone')
  expect(f.native().masters[0]!.themeColors.Accent1).toBe('#FFFFFF')
})
it('supports32 unique operations on multiple master/layout targets and all600 dependency pages', async () => {
  const started = Date.now()
  const f = await fixture(600)
  const slots = [
    'Accent1',
    'Accent2',
    'Accent3',
    'Accent4',
    'Accent5',
    'Accent6',
    'Dark1',
    'Dark2',
    'Light1',
    'Light2',
    'Hyperlink',
    'FollowedHyperlink',
  ]
  const ops: PowerPointMasterOperation[] = [0, 1].flatMap((i) =>
    slots.map((theme_color) => ({
      op: 'set_master_theme_color' as const,
      master_id: `m${i}`,
      theme_color,
      color: '#000000',
    })),
  )
  // A separate layout/master target remains editable alongside every theme slot.
  ops.push(
    {
      op: 'set_master_background',
      master_id: 'm0',
      fill: { type: 'solid', color: '#000000', transparency: 0 },
    },
    {
      op: 'set_master_background',
      master_id: 'm1',
      fill: { type: 'solid', color: '#000000', transparency: 0 },
    },
  )
  for (let i = 0; i < 6; i++) {
    const master = f.native().masters[i % 2]!
    master.layouts.push({ ...master.layouts[0]!, id: `extra${i}` })
    ops.push({
      op: 'set_layout_background_following',
      master_id: master.id,
      layout_id: `extra${i}`,
      follow_master: false,
      show_master_graphics: false,
    })
  }
  const p = await f.propose(ops)
  const prepared = Date.now()
  await f.confirm()
  console.info('native-master600x32 timings', {
    prepareMs: prepared - started,
    confirmMs: Date.now() - prepared,
    totalMs: Date.now() - started,
  })
  expect(f.data.get(String(p.preview.changeId))!.scope.affectedPageCount).toBe(600)
  expect(f.data.get(String(p.preview.changeId))!.receipts).toHaveLength(32)
  expect(f.data.get(String(p.preview.changeId))!.nextIndex).toBe(32)
  expect(f.adapter.executeMasterOperations).toHaveBeenCalledTimes(32)
}, 900000)
it.each(['before_failure', 'after_failure'])(
  'retains %s pending without automatic undo/replay after restart',
  async (mode) => {
    const f = await fixture()
    const p = await f.propose()
    f.setHost(mode)
    await expect(f.confirm()).rejects.toThrow()
    const id = String(p.preview.changeId)
    expect(f.data.get(id)!.pending).toBeDefined()
    expect(f.adapter.executeMasterOperations).toHaveBeenCalledTimes(1)
    f.reopen()
    const inspected = await f.tool('inspect_slide_master_change', { change_id: id })
    expect(inspected.isError).not.toBe(true)
    const resumed = await f.tool('resume_slide_master_change', { change_id: id })
    expect(resumed.isError).toBe(true)
    expect(f.adapter.executeMasterOperations).toHaveBeenCalledTimes(1)
  },
)
it('reconciles a durable observed proof after a lost receipt ACK and then undoes', async () => {
  const f = await fixture()
  const p = await f.propose()
  f.setReceipt('lost_ack')
  await expect(f.confirm()).rejects.toThrow()
  const id = String(p.preview.changeId)
  expect(f.data.get(id)!.pending?.afterProofRef).toBeDefined()
  f.reopen()
  f.setReceipt('normal')
  await f.tool('reconcile_slide_master_change', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('applied')
  expect(f.adapter.executeMasterOperations).toHaveBeenCalledTimes(1)
  await f.tool('undo_slide_master_change', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('undone')
})
it('explicit before-state reconciliation closes pending without replay and allows prefix undo', async () => {
  const f = await fixture()
  const p = await f.propose()
  f.setHost('before_failure')
  await expect(f.confirm()).rejects.toThrow()
  const id = String(p.preview.changeId)
  f.reopen()
  await f.tool('reconcile_slide_master_change', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('undone')
  expect(f.adapter.executeMasterOperations).toHaveBeenCalledTimes(1)
})
it.each(['page', 'order', 'dependency', 'other_master'])(
  'refuses %s drift before inverse writes',
  async (kind) => {
    const f = await fixture()
    const p = await f.propose()
    await f.confirm()
    const id = String(p.preview.changeId)
    if (kind === 'page') f.drift.set('s0', 'changed')
    if (kind === 'order') f.order.reverse()
    if (kind === 'dependency') f.dependencies.slides[0]!.layoutId = 'different'
    if (kind === 'other_master') f.native().masters[1]!.themeColors.Accent1 = '#123456'
    const outcome = await f.tool('undo_slide_master_change', { change_id: id })
    expect(outcome.isError).toBe(true)
    expect(f.adapter.executeMasterOperations).toHaveBeenCalledTimes(1)
  },
)
it.each(['doc', 'disconnect', 'clear', 'cancel'])(
  'rejects %s changes after asynchronous snapshot reads without host writes',
  async (kind) => {
    const f = await fixture()
    const controller = new AbortController()
    await f.propose(undefined, controller.signal)
    f.setSdkRead(() => {
      if (kind === 'doc') f.setDocument()
      if (kind === 'disconnect') f.disconnect()
      if (kind === 'clear') f.clear()
      if (kind === 'cancel') controller.abort()
    })
    await expect(f.confirm()).rejects.toThrow()
    expect(f.adapter.executeMasterOperations).not.toHaveBeenCalled()
  },
)
it('captures and stores historical page review without certifying whole deck QA', async () => {
  const f = await fixture()
  const p = await f.propose()
  await f.confirm()
  const id = String(p.preview.changeId)
  const capture = await f.tool('capture_slide_master_page', { change_id: id, slide_id: 's0' })
  const evidence = JSON.parse(capture.output)
  expect(evidence.qaPassed).toBe(false)
  const reviewed = await f.tool('record_slide_master_page_review', {
    change_id: id,
    slide_id: 's0',
    screenshot_digest: evidence.screenshotDigest,
    status: 'pass',
    notes: 'Looks consistent',
  })
  expect(reviewed.isError).not.toBe(true)
  expect(f.data.get(id)!.reviews).toHaveLength(1)
})

it.each(['s0', 'note', 'theme_font', 'master_placeholder'])(
  'retains unknown pending when foreign %s content races inside an acknowledged SDK write',
  async (target) => {
    const f = await fixture()
    const p = await f.propose()
    f.setBeforeWrite(() => f.drift.set(target, 'Foreign user edit'))
    await expect(f.confirm()).rejects.toThrow()
    const saved = f.data.get(String(p.preview.changeId))!
    expect(saved.pending).toBeDefined()
    expect(saved.pending!.afterProofRef).toBeUndefined()
    expect(saved.receipts).toHaveLength(0)
    expect(f.adapter.executeMasterOperations).toHaveBeenCalledTimes(1)
    f.reopen()
    const inspected = await f.tool('inspect_slide_master_change', { change_id: saved.changeId })
    expect(JSON.parse(inspected.output).status).toBe('unknown')
    expect(f.adapter.executeMasterOperations).toHaveBeenCalledTimes(1)
  },
)

it('keeps large image bytes on the PC and proves the specific master background media before receipts', async () => {
  const f = await fixture()
  const bytes = new Uint8Array(256 * 1024)
  bytes.set([0x89, 0x50, 0x4e, 0x47])
  const image_base64 = Buffer.from(bytes).toString('base64')
  const p = await f.propose([
    {
      op: 'set_master_background',
      master_id: 'm0',
      fill: { type: 'picture_or_texture', image_base64, transparency: 0.2 },
    },
  ])
  await f.confirm()
  const saved = f.data.get(String(p.preview.changeId))!
  expect(saved.state).toBe('applied')
  expect(JSON.stringify(saved)).not.toContain('image_base64')
  expect(JSON.stringify(saved)).not.toContain(image_base64)
  expect(saved.operations[0]!.op).toBe('set_master_background')
  await f.tool('undo_slide_master_change', { change_id: saved.changeId })
  await f.confirm()
  expect(f.data.get(saved.changeId)!.state).toBe('undone')
})

it('leaves acknowledged but normalized or wrong picture media unknown without undo or promotion', async () => {
  const f = await fixture()
  const bytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4, 5, 6, 7, 8, 9])
  const p = await f.propose([
    {
      op: 'set_master_background',
      master_id: 'm0',
      fill: {
        type: 'picture_or_texture',
        image_base64: Buffer.from(bytes).toString('base64'),
        transparency: 0,
      },
    },
  ])
  f.setHost('wrong_picture')
  await expect(f.confirm()).rejects.toThrow()
  const saved = f.data.get(String(p.preview.changeId))!
  expect(saved.pending?.afterProofRef).toBeUndefined()
  expect(saved.receipts).toHaveLength(0)
  f.reopen()
  const result = await f.tool('reconcile_slide_master_change', { change_id: saved.changeId })
  expect(result.isError).toBe(true)
  expect(f.adapter.executeMasterOperations).toHaveBeenCalledTimes(1)
})

it('reconciles inverse lost ACK after restart without repeating the inverse write', async () => {
  const f = await fixture(),
    p = await f.propose()
  await f.confirm()
  const id = String(p.preview.changeId)
  await f.tool('undo_slide_master_change', { change_id: id })
  f.setReceipt('lost_ack')
  await expect(f.confirm()).rejects.toThrow()
  expect(f.data.get(id)!.pending?.direction).toBe('undo')
  f.reopen()
  f.setReceipt('normal')
  await f.tool('reconcile_slide_master_change', { change_id: id })
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('undone')
  expect(f.adapter.executeMasterOperations).toHaveBeenCalledTimes(2)
})

it('rejects forged concurrent durable state after an awaited SDK observation', async () => {
  const f = await fixture(),
    p = await f.propose()
  await f.confirm()
  const id = String(p.preview.changeId)
  f.setSdkRead(() => {
    const saved = f.data.get(id)!
    f.data.set(id, { ...saved, intent: 'Concurrent checkpoint edit' })
  })
  const undo = await f.tool('undo_slide_master_change', { change_id: id })
  expect(undo.isError).toBe(true)
  expect(f.adapter.executeMasterOperations).toHaveBeenCalledTimes(1)
})

it.each(['gradient', 'sysClr', 'unused'])(
  'rejects unrepresentable original %s inverse before host writes',
  async (kind) => {
    const f = await fixture(kind === 'unused' ? 1 : 2)
    if (kind === 'gradient') {
      f.native().masters[0]!.background = { type: 'Gradient', gradientType: 'Linear' } as any
      f.drift.set(
        'original_bg',
        '<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="12000"><a:srgbClr val="123456"/></a:gs></a:gsLst><a:lin ang="40000"/></a:gradFill></p:bgPr></p:bg>',
      )
    }
    if (kind === 'sysClr') f.drift.set('original_slot', '<a:sysClr val="window" lastClr="FFFFFF"/>')
    const ops: PowerPointMasterOperation[] =
      kind === 'gradient'
        ? [
            {
              op: 'set_master_background',
              master_id: 'm0',
              fill: { type: 'solid', color: '#000000', transparency: 0 },
            },
          ]
        : [{ ...f.op, master_id: kind === 'unused' ? 'm1' : 'm0' }]
    await expect(f.propose(ops)).rejects.toThrow('presentation_native_master_inverse_unproven')
    expect(f.adapter.executeMasterOperations).not.toHaveBeenCalled()
    expect(f.data.size).toBe(0)
  },
)
it.each(['background', 'theme'])(
  'keeps inverse pending if native %s value matches but original XML does not',
  async (kind) => {
    const f = await fixture()
    const operations: PowerPointMasterOperation[] =
      kind === 'background'
        ? [
            {
              op: 'set_master_background',
              master_id: 'm0',
              fill: { type: 'solid', color: '#000000', transparency: 0 },
            },
          ]
        : [f.op]
    const p = await f.propose(operations)
    await f.confirm()
    const id = String(p.preview.changeId)
    f.setBeforeWrite(() => {
      if (kind === 'background')
        f.drift.set(
          'original_bg',
          '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="99000"/></a:srgbClr></a:solidFill></p:bgPr></p:bg>',
        )
      else f.drift.set('original_slot', '<a:sysClr val="window" lastClr="FFFFFF"/>')
    })
    await f.tool('undo_slide_master_change', { change_id: id })
    await expect(f.confirm()).rejects.toThrow('presentation_native_master_package_unproven')
    expect(f.data.get(id)!.state).toBe('undoing')
    expect(f.data.get(id)!.pending?.direction).toBe('undo')
    expect(f.data.get(id)!.receipts).toHaveLength(1)
  },
)

it('rejects original pattern alpha before any host write', async () => {
  const f = await fixture()
  f.native().masters[0]!.background = {
    type: 'Pattern',
    pattern: 'Percent5',
    foregroundColor: '#FFFFFF',
    backgroundColor: '#000000',
  }
  f.drift.set(
    'original_bg',
    '<p:bg><p:bgPr><a:pattFill prst="pct5"><a:fgClr><a:srgbClr val="FFFFFF"><a:alpha val="50000"/></a:srgbClr></a:fgClr><a:bgClr><a:srgbClr val="000000"/></a:bgClr></a:pattFill></p:bgPr></p:bg>',
  )
  await expect(
    f.propose([
      {
        op: 'set_master_background',
        master_id: 'm0',
        fill: { type: 'solid', color: '#000000', transparency: 0 },
      },
    ]),
  ).rejects.toThrow('presentation_native_master_inverse_unproven')
  expect(f.adapter.executeMasterOperations).not.toHaveBeenCalled()
  expect(f.data.size).toBe(0)
})

it('matches original lower-case RGB hex with upper-case native inverse XML', async () => {
  const f = await fixture()
  f.drift.set('original_slot', '<a:srgbClr val="ffffff"/>')
  const p = await f.propose()
  f.setBeforeWrite(() => f.drift.delete('original_slot'))
  await f.confirm()
  await f.tool('undo_slide_master_change', { change_id: String(p.preview.changeId) })
  await f.confirm()
  expect(f.data.get(String(p.preview.changeId))!.state).toBe('undone')
})
it('rejects a theme native/XML preimage color mismatch before host write', async () => {
  const f = await fixture()
  f.drift.set('original_slot', '<a:srgbClr val="123456"/>')
  await expect(f.propose()).rejects.toThrow('presentation_native_master_inverse_unproven')
  expect(f.adapter.executeMasterOperations).not.toHaveBeenCalled()
  expect(f.data.size).toBe(0)
})
it.each(['gradient', 'pattern'])(
  'preserves %s input from an exactly recoverable original solid',
  async (type) => {
    const f = await fixture()
    const fill: PowerPointMasterOperation =
      type === 'gradient'
        ? {
            op: 'set_master_background',
            master_id: 'm0',
            fill: { type: 'gradient', gradient_type: 'Linear' },
          }
        : {
            op: 'set_master_background',
            master_id: 'm0',
            fill: {
              type: 'pattern',
              pattern: 'Percent5',
              foreground_color: '#000000',
              background_color: '#FFFFFF',
            },
          }
    const p = await f.propose([fill])
    await f.confirm()
    await f.tool('undo_slide_master_change', { change_id: String(p.preview.changeId) })
    await f.confirm()
    expect(f.data.get(String(p.preview.changeId))!.state).toBe('undone')
  },
)
it.each([
  '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
  '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:pattFill prst="pct5"/>',
  '<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="50000"/></a:srgbClr></a:solidFill>',
  '<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="100001"/></a:srgbClr></a:solidFill>',
])(
  'rejects nonrepresentable fill structure or native/XML transparency mismatch %s',
  async (xml) => {
    const f = await fixture()
    f.drift.set('original_bg', `<p:bg><p:bgPr>${xml}</p:bgPr></p:bg>`)
    await expect(
      f.propose([
        {
          op: 'set_master_background',
          master_id: 'm0',
          fill: { type: 'solid', color: '#000000', transparency: 0 },
        },
      ]),
    ).rejects.toThrow('presentation_native_master_inverse_unproven')
    expect(f.adapter.executeMasterOperations).not.toHaveBeenCalled()
    expect(f.data.size).toBe(0)
  },
)

it.each(['100000', '070000'])(
  'normalizes exactly representable solid alpha %s on inverse',
  async (alpha) => {
    const f = await fixture()
    f.native().masters[0]!.background.transparency = 1 - Number(alpha) / 100000
    f.drift.set(
      'original_bg',
      `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="${alpha}"/></a:srgbClr></a:solidFill></p:bgPr></p:bg>`,
    )
    const p = await f.propose([
      {
        op: 'set_master_background',
        master_id: 'm0',
        fill: { type: 'solid', color: '#000000', transparency: 0 },
      },
    ])
    f.setBeforeWrite(() => f.drift.delete('original_bg'))
    await f.confirm()
    f.setBeforeWrite(() => {
      const numeric = String(Number(alpha))
      f.drift.set(
        'original_bg',
        `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF">${numeric === '100000' ? '' : `<a:alpha val="${numeric}"/>`}</a:srgbClr></a:solidFill></p:bgPr></p:bg>`,
      )
    })
    await f.tool('undo_slide_master_change', { change_id: String(p.preview.changeId) })
    await f.confirm()
    expect(f.data.get(String(p.preview.changeId))!.state).toBe('undone')
  },
)

it('rejects a hash-valid persisted snapshot missing derived target keys after reopen', async () => {
  const f = await fixture()
  const p = await f.propose()
  await f.confirm()
  const id = String(p.preview.changeId),
    r = f.data.get(id)!
  const original = await readMasterBackup({
    request: f.request,
    documentId: r.documentId,
    changeId: id,
    backup: r.snapshotRef,
  })
  const snapshot = JSON.parse(new TextDecoder().decode(original))
  snapshot.pages[0].protection.targetDigests = {}
  const bytes = new TextEncoder().encode(JSON.stringify(snapshot))
  const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (b) =>
    b.toString(16).padStart(2, '0'),
  ).join('')
  const implementation = f.request.getMockImplementation()!
  f.data.set(id, {
    ...r,
    snapshotRef: { ...r.snapshotRef, sha256: digest, sizeBytes: bytes.length },
  })
  f.request.mockImplementation(async (body, signal) => {
    const result = await implementation(body, signal)
    const input = body as any
    if (input.key !== 'snapshot') return result
    const changed = { ...(result as any), sha256: digest, sizeBytes: bytes.length }
    if (input.operation === 'master_backup_read')
      changed.base64 = Buffer.from(
        bytes.subarray(input.offset, input.offset + input.length),
      ).toString('base64')
    return changed
  })
  f.reopen()
  expect(await f.tool('resume_slide_master_change', { change_id: id })).toMatchObject({
    isError: true,
    mutated: false,
    summary: 'presentation_master_backup_invalid',
  })
  expect(f.adapter.executeMasterOperations).toHaveBeenCalledTimes(1)
})
