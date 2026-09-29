import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationService } from '../../shell/src/main/presentation-service'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { BrowserPresentationPageReplacementAdapter } from '../src/skills/powerpoint/browser-presentation-page-replacement-adapter'
import { createPresentationSlideDuplicationSkill } from '../src/skills/powerpoint/presentation-slide-duplication'
import {
  validatePresentationExistingBatch,
  validExistingBatchTransition,
  type PresentationSlideDuplicationBatch,
} from '../src/skills/powerpoint/presentation-existing-batch'
const roots: string[] = []
const globals = { Office: (globalThis as any).Office, PowerPoint: (globalThis as any).PowerPoint }
afterEach(() => {
  Object.assign(globalThis, globals)
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
async function fixture(count = 2) {
  const root = mkdtempSync(join(tmpdir(), 'native-duplicate-'))
  roots.push(root)
  const service = createPresentationService({ userDataPath: root })
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const { bytes } = await compilePresentationDeck(deck)
  const original = Buffer.from(bytes).toString('base64')
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('ppt/presentation.xml')!.async('string')
  const packageSourceId = `${xml.match(/<p:sldId\b[^>]*\bid="([0-9]+)"/)![1]}#`
  let order = Array.from({ length: count }, (_, i) => `native-source-${i}`)
  const packages = new Map(order.map((id) => [id, original]))
  let documentName = 'deck',
    available = true,
    failBackup = false,
    failInsertedReceipt = false,
    failAppliedReceipt = false,
    failUndoneReceipt = false,
    failFlightReceipt = false,
    disconnectFinalBackup = false,
    switchFinalBackup = false
  let inserted = 0,
    deleted = 0
  let beforeInsert: () => void = () => {},
    beforeDelete: () => void = () => {}
  let documentPause: { entered: () => void; gate: Promise<void> } | undefined
  let lastInsertion: any
  const page = (id: string) => ({
    id,
    load() {
      if (!order.includes(id)) throw Error('office_read_failed')
    },
    exportAsBase64: () => ({ value: packages.get(id)! }),
    delete: () => {
      beforeDelete()
      deleted++
      order = order.filter((value) => value !== id)
    },
  })
  const slides = {
    load() {},
    get items() {
      return order.map(page)
    },
    getItem: page,
  }
  const presentation = {
    slides,
    insertSlidesFromBase64: (base64: string, input: any) => {
      lastInsertion = structuredClone(input)
      beforeInsert()
      inserted++
      const id = `native-copy-${inserted}`
      order.splice(order.indexOf(input.targetSlideId) + 1, 0, id)
      packages.set(id, base64)
    },
  }
  ;(globalThis as any).Office = {
    context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } },
  }
  ;(globalThis as any).PowerPoint = {
    run: async (callback: any) => callback({ presentation, sync: async () => {} }),
  }
  const values = new Map<string, string>()
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save: async () => {},
    location: () => documentName,
  }
  let binding = createPresentationDocumentBinding(settings, () => documentName)
  const request = vi.fn(async (body: any, signal?: AbortSignal) => {
    if (failBackup) throw Error('backup_failed')
    const response = new Response(
      Buffer.from(await service(body, signal ?? new AbortController().signal)).toString('utf8'),
    )
    if (body.operation === 'existing_page_backup_read') {
      if (disconnectFinalBackup) available = false
      if (switchFinalBackup) documentName = 'other'
    }
    return response
  })
  const png =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII='
  const adapter = {
    snapshotSlide: vi.fn(async (index: number) => ({
      slideId: order[index],
      fingerprint: 'native-source',
    })),
    exportPresentationPagePackage: vi.fn(async (id: string) => ({
      slideId: id,
      slideIds: [...order],
      base64: packages.get(id)!,
    })),
    inspectPresentationPage: vi.fn(async (id: string) => ({
      slideId: id,
      shapes: [],
      shapesTruncated: false,
      screenshot: { mime: 'image/png', base64: png },
    })),
    duplicateSlide: vi.fn(),
    executeDeclarative: vi.fn(),
  }
  const proposals = createStructuredProposalController(),
    pageAdapter = new BrowserPresentationPageReplacementAdapter()
  const write = vi.fn(async (next: any, prior: any) => {
    if (failFlightReceipt && next.inFlightIndex === 0 && prior?.inFlightIndex === undefined)
      throw Error('flight_save_failed')
    if (
      (failInsertedReceipt && next.state === 'applying' && next.insertedSlideId) ||
      (failAppliedReceipt && next.state === 'applied') ||
      (failUndoneReceipt && next.state === 'undone')
    )
      throw Error('ack_lost')
    await binding.writeExistingBatch(next, prior)
  })
  const make = () =>
    createPresentationSlideDuplicationSkill({
      documentId: async () => {
        const id = await binding.documentId()
        const pause = documentPause
        documentPause = undefined
        if (pause) {
          pause.entered()
          await pause.gate
        }
        return id
      },
      available: () => available,
      adapter: adapter as any,
      pageAdapter,
      request,
      proposals,
      readExistingBatch: (id) => binding.readExistingBatch(id),
      writeExistingBatch: write,
    })
  let skill = make()
  return {
    armDocumentPause: (entered: () => void, gate: Promise<void>) => {
      documentPause = { entered, gate }
    },
    adapter,
    pageAdapter,
    proposals,
    request,
    write,
    packageSourceId,
    propose: (index = 0) => skill.propose(index),
    saved: (id: string) => binding.readExistingBatch(id) as PresentationSlideDuplicationBatch,
    reopen: () => {
      binding = createPresentationDocumentBinding(settings, () => documentName)
      skill = make()
    },
    call: (name: string, changeId: string, more: Record<string, unknown> = {}) =>
      skill.executeTool({ id: 'recovery', name, input: { change_id: changeId, ...more } }),
    order: () => [...order],
    inserted: () => inserted,
    deleted: () => deleted,
    lastInsertion: () => lastInsertion,
    setFailBackup: () => (failBackup = true),
    setFailInserted: () => (failInsertedReceipt = true),
    setFailApplied: () => (failAppliedReceipt = true),
    setFailUndone: () => (failUndoneReceipt = true),
    setFailFlight: () => (failFlightReceipt = true),
    insertForeignCopy: () => {
      order.splice(1, 0, 'foreign-copy')
      packages.set('foreign-copy', original)
    },
    clearFailures: () => {
      failInsertedReceipt = false
      failAppliedReceipt = false
      failUndoneReceipt = false
      failFlightReceipt = false
    },
    setDisconnectBackup: () => (disconnectFinalBackup = true),
    setSwitchBackup: () => (switchFinalBackup = true),
    setAvailable: (value: boolean) => (available = value),
    setBeforeInsert: (fn: () => void) => (beforeInsert = fn),
    setBeforeDelete: (fn: () => void) => (beforeDelete = fn),
    driftOrder: () => order.reverse(),
    changeSource: async () => {
      zip.file(
        'ppt/slides/slide1.xml',
        (await zip.file('ppt/slides/slide1.xml')!.async('string')).replace(
          '</p:spTree>',
          '<p:extLst><p:ext uri="source-changed"/></p:extLst></p:spTree>',
        ),
      )
      packages.set('native-source-0', await zip.generateAsync({ type: 'base64' }))
    },
    changeCopy: async () => {
      zip.file(
        'ppt/slides/slide1.xml',
        (await zip.file('ppt/slides/slide1.xml')!.async('string')).replace(
          '</p:spTree>',
          '<p:extLst><p:ext uri="changed"/></p:extLst></p:spTree>',
        ),
      )
      packages.set('native-copy-1', await zip.generateAsync({ type: 'base64' }))
    },
  }
}
it('backs up the actual source package and stages a single adjacent page only after confirmation', async () => {
  const f = await fixture(),
    p = await f.propose()
  expect(f.inserted()).toBe(0)
  await f.proposals.confirm(p.proposalId)
  expect(f.inserted()).toBe(1)
  expect(f.lastInsertion()).toEqual({
    targetSlideId: 'native-source-0',
    sourceSlideIds: [f.packageSourceId],
    formatting: 'KeepSourceFormatting',
  })
  expect(f.adapter.duplicateSlide).not.toHaveBeenCalled()
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  expect(f.saved(p.changeId)).toMatchObject({
    version: 4,
    kind: 'native_slide_duplicate',
    state: 'applied',
    nextIndex: 1,
    insertedSlideId: 'native-copy-1',
    sourceSlideId: f.packageSourceId,
  })
  f.reopen()
  const read = await f.call('inspect_slide_duplication', p.changeId)
  expect(JSON.parse(read.output)).toMatchObject({
    state: 'applied',
    hostStatus: 'staged',
    currentPackageMatches: true,
    qaPassed: false,
  })
})
it('refuses backup failure before saving an intent or inserting', async () => {
  const f = await fixture(),
    p = await f.propose()
  f.setFailBackup()
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow('backup_failed')
  expect(f.inserted()).toBe(0)
  expect(f.saved(p.changeId)).toBeUndefined()
})
it.each(['disconnect', 'switch'])(
  'rechecks %s after the last backup read before insertion',
  async (kind) => {
    const f = await fixture(),
      p = await f.propose()
    if (kind === 'disconnect') f.setDisconnectBackup()
    else f.setSwitchBackup()
    await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow()
    expect(f.inserted()).toBe(0)
  },
)
it.each(['inserted', 'applied'])(
  'retains uncertain %s receipt across reopening and explicitly claims without another insertion',
  async (kind) => {
    const f = await fixture(),
      p = await f.propose()
    if (kind === 'inserted') f.setFailInserted()
    else f.setFailApplied()
    await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow('ack_lost')
    expect(f.inserted()).toBe(1)
    expect(f.saved(p.changeId).state).toBe('applying')
    f.clearFailures()
    f.reopen()
    const recovered = await f.call('reconcile_slide_duplication', p.changeId)
    expect(recovered.isError).toBeUndefined()
    expect(f.saved(p.changeId).state).toBe('applying')
    await f.proposals.confirm(JSON.parse(recovered.output).proposalId)
    expect(f.saved(p.changeId)).toMatchObject({
      state: 'applied',
      nextIndex: 1,
      insertedSlideId: 'native-copy-1',
    })
    expect(f.inserted()).toBe(1)
  },
)
it('explicitly closes a strongly proven unchanged baseline without insertion or deletion replay', async () => {
  const f = await fixture(),
    p = await f.propose()
  f.pageAdapter.stage = vi.fn(async () => {
    throw Error('office_api_unsupported')
  })
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow('office_api_unsupported')
  expect(f.saved(p.changeId)).toMatchObject({ state: 'applying', inFlightIndex: 0, nextIndex: 0 })
  const close = await f.call('reconcile_slide_duplication', p.changeId)
  expect(close.isError).toBeUndefined()
  expect(f.saved(p.changeId).state).toBe('applying')
  await f.proposals.confirm(JSON.parse(close.output).proposalId)
  expect(f.saved(p.changeId)).toMatchObject({ state: 'undone', nextIndex: 0 })
  expect(f.saved(p.changeId).insertedSlideId).toBeUndefined()
  expect(f.saved(p.changeId).inFlightIndex).toBeUndefined()
  expect(f.inserted()).toBe(0)
  expect(f.deleted()).toBe(0)
})

it('undo removes only the unchanged owned copy and preserves the original deck order', async () => {
  const f = await fixture(),
    p = await f.propose()
  await f.proposals.confirm(p.proposalId)
  const undo = await f.call('undo_slide_duplication', p.changeId)
  expect(f.deleted()).toBe(0)
  await f.proposals.confirm(JSON.parse(undo.output).proposalId)
  expect(f.deleted()).toBe(1)
  expect(f.order()).toEqual(['native-source-0', 'native-source-1'])
  expect(f.saved(p.changeId).state).toBe('undone')
  const inspect = await f.call('inspect_slide_duplication', p.changeId)
  expect(JSON.parse(inspect.output).currentPackageMatches).toBe(true)
})
it('never deletes again after removal ACK loss; explicitly closes the proven baseline receipt', async () => {
  const f = await fixture(),
    p = await f.propose()
  await f.proposals.confirm(p.proposalId)
  f.setFailUndone()
  const undo = await f.call('undo_slide_duplication', p.changeId)
  await expect(f.proposals.confirm(JSON.parse(undo.output).proposalId)).rejects.toThrow('ack_lost')
  expect(f.saved(p.changeId).state).toBe('undoing')
  expect(f.deleted()).toBe(1)
  f.clearFailures()
  f.reopen()
  const close = await f.call('undo_slide_duplication', p.changeId)
  await f.proposals.confirm(JSON.parse(close.output).proposalId)
  expect(f.deleted()).toBe(1)
  expect(f.saved(p.changeId).state).toBe('undone')
})
it.each(['content', 'order'])('refuses to delete a copy after %s drift', async (kind) => {
  const f = await fixture(),
    p = await f.propose()
  await f.proposals.confirm(p.proposalId)
  if (kind === 'content') await f.changeCopy()
  else f.driftOrder()
  const undo = await f.call('undo_slide_duplication', p.changeId)
  expect(undo.isError).toBe(true)
  expect(f.deleted()).toBe(0)
})
it('checks write availability in the native adapter final insertion guard', async () => {
  const f = await fixture(),
    p = await f.propose()
  const stage = f.pageAdapter.stage.bind(f.pageAdapter)
  f.pageAdapter.stage = async (record, base64, onInserted, guard, signal) =>
    stage(
      record,
      base64,
      onInserted,
      async () => {
        f.setAvailable(false)
        await guard()
      },
      signal,
    )
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow(
    'presentation_existing_persistence_unavailable',
  )
  expect(f.inserted()).toBe(0)
})
it('supports 512 source pages and verifies insertion plus removal at 513 pages', async () => {
  const f = await fixture(512),
    p = await f.propose(511)
  await f.proposals.confirm(p.proposalId)
  expect(f.order()).toHaveLength(513)
  const undo = await f.call('undo_slide_duplication', p.changeId)
  await f.proposals.confirm(JSON.parse(undo.output).proposalId)
  expect(f.order()).toHaveLength(512)
})
it('captures and records only historical review of the inserted page', async () => {
  const f = await fixture(),
    p = await f.propose()
  await f.proposals.confirm(p.proposalId)
  const capture = await f.call('capture_slide_duplication_page', p.changeId)
  const shot = JSON.parse(capture.output)
  expect(shot.qaPassed).toBe(false)
  const review = await f.call('record_slide_duplication_page_review', p.changeId, {
    screenshot_digest: shot.screenshotDigest,
    status: 'pass',
    notes: 'Readability checked',
  })
  expect(review.isError).toBeUndefined()
  expect(f.saved(p.changeId).reviews?.[0]).toMatchObject({
    hostSlideId: 'native-copy-1',
    status: 'pass',
  })
  f.reopen()
  const stale = await f.call('record_slide_duplication_page_review', p.changeId, {
    screenshot_digest: shot.screenshotDigest,
    status: 'pass',
    notes: 'Reused',
  })
  expect(stale.isError).toBe(true)
})
it('strictly validates immutable v4 operation/source/receipt transitions', async () => {
  const f = await fixture(),
    p = await f.propose()
  await f.proposals.confirm(p.proposalId)
  const applied = f.saved(p.changeId)
  expect(validatePresentationExistingBatch(applied)).toBe(true)
  expect(
    validatePresentationExistingBatch({ ...applied, sourceSlideId: applied.hostSlideId }),
  ).toBe(false)
  expect(
    validatePresentationExistingBatch({ ...applied, insertedSlideId: applied.hostSlideId }),
  ).toBe(false)
  expect(
    validExistingBatchTransition(applied, {
      ...applied,
      operations: [{ op: 'duplicate_slide', slide_index: 1 }],
    }),
  ).toBe(false)
  expect(validatePresentationExistingBatch({ ...applied, nextIndex: 0 })).toBe(false)
})

it('keeps failed absence-closure ACK pending until a separately confirmed exact baseline proof', async () => {
  const f = await fixture(),
    p = await f.propose()
  f.pageAdapter.stage = vi.fn(async () => {
    throw Error('office_api_unsupported')
  })
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow()
  f.setFailUndone()
  const close = await f.call('reconcile_slide_duplication', p.changeId)
  await expect(f.proposals.confirm(JSON.parse(close.output).proposalId)).rejects.toThrow('ack_lost')
  expect(f.saved(p.changeId)).toMatchObject({ state: 'applying', inFlightIndex: 0 })
  f.clearFailures()
  f.reopen()
  const again = await f.call('reconcile_slide_duplication', p.changeId)
  await f.proposals.confirm(JSON.parse(again.output).proposalId)
  expect(f.saved(p.changeId).state).toBe('undone')
  expect(f.inserted()).toBe(0)
  expect(f.deleted()).toBe(0)
})

it('freezes an exported source object before asynchronous identity checks and backup saves', async () => {
  const f = await fixture(),
    mutable = await f.adapter.exportPresentationPagePackage('native-source-0')
  let entered!: () => void, release!: () => void
  const started = new Promise<void>((resolve) => (entered = resolve)),
    gate = new Promise<void>((resolve) => (release = resolve))
  f.adapter.exportPresentationPagePackage.mockImplementationOnce(async () => {
    f.armDocumentPause(entered, gate)
    return mutable
  })
  const pending = f.propose()
  await started
  mutable.base64 = Buffer.from('changed external reference').toString('base64')
  mutable.slideIds.reverse()
  release()
  const p = await pending
  await f.proposals.confirm(p.proposalId)
  expect(f.saved(p.changeId).beforeSlideIds).toEqual(['native-source-0', 'native-source-1'])
  expect(f.lastInsertion().sourceSlideIds).toEqual([f.packageSourceId])
  expect(f.inserted()).toBe(1)
})

it('explicitly closes initial intent after flight persistence failed without a native insertion', async () => {
  const f = await fixture(),
    p = await f.propose()
  f.setFailFlight()
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow('flight_save_failed')
  expect(f.saved(p.changeId)).toMatchObject({ state: 'applying', nextIndex: 0 })
  expect(f.saved(p.changeId).inFlightIndex).toBeUndefined()
  expect(f.inserted()).toBe(0)
  f.clearFailures()
  f.reopen()
  const inspected = await f.call('inspect_slide_duplication', p.changeId)
  expect(JSON.parse(inspected.output)).toMatchObject({
    status: 'baseline',
    currentPackageMatches: true,
  })
  const close = await f.call('reconcile_slide_duplication', p.changeId)
  expect(close.isError).toBeUndefined()
  await f.proposals.confirm(JSON.parse(close.output).proposalId)
  expect(f.saved(p.changeId)).toMatchObject({ state: 'undone', nextIndex: 0 })
  expect(f.inserted()).toBe(0)
  expect(f.deleted()).toBe(0)
})
it('never claims an adjacent foreign copy without a durable native insertion intent', async () => {
  const f = await fixture(),
    p = await f.propose()
  f.setFailFlight()
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow('flight_save_failed')
  f.clearFailures()
  f.insertForeignCopy()
  f.reopen()
  const claim = await f.call('reconcile_slide_duplication', p.changeId)
  expect(claim.isError).toBe(true)
  expect(f.saved(p.changeId)).toMatchObject({ state: 'applying', nextIndex: 0 })
  expect(f.inserted()).toBe(0)
  expect(f.deleted()).toBe(0)
})

it('does not close an absence receipt after the source package drifted before confirmation', async () => {
  const f = await fixture(),
    p = await f.propose()
  f.pageAdapter.stage = vi.fn(async () => {
    throw Error('office_api_unsupported')
  })
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow()
  const close = await f.call('reconcile_slide_duplication', p.changeId)
  expect(close.isError).toBeUndefined()
  await f.changeSource()
  await expect(f.proposals.confirm(JSON.parse(close.output).proposalId)).rejects.toThrow(
    'proposal_stale',
  )
  expect(f.saved(p.changeId)).toMatchObject({ state: 'applying', inFlightIndex: 0, nextIndex: 0 })
  expect(f.inserted()).toBe(0)
  expect(f.deleted()).toBe(0)
})
it('freezes fresh confirmation exports before asynchronous document and digest guards', async () => {
  const f = await fixture(),
    p = await f.propose(),
    mutable = await f.adapter.exportPresentationPagePackage('native-source-0')
  let entered!: () => void, release!: () => void
  const started = new Promise<void>((resolve) => (entered = resolve)),
    gate = new Promise<void>((resolve) => (release = resolve))
  f.adapter.exportPresentationPagePackage.mockImplementationOnce(async () => {
    f.armDocumentPause(entered, gate)
    return mutable
  })
  const confirmation = f.proposals.confirm(p.proposalId)
  await started
  mutable.base64 = Buffer.from('changed fresh export reference').toString('base64')
  mutable.slideIds.reverse()
  release()
  await expect(confirmation).resolves.toBeUndefined()
  expect(f.saved(p.changeId)).toMatchObject({
    state: 'applied',
    nextIndex: 1,
    insertedSlideId: 'native-copy-1',
  })
  expect(f.inserted()).toBe(1)
})
