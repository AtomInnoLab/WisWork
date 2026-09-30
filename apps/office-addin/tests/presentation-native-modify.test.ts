import { afterEach, it, expect, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationService } from '../../shell/src/main/presentation-service'
import { createPresentationNativeModifySkill } from '../src/skills/powerpoint/presentation-native-modify'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  validatePresentationExistingBatch,
  type NativeModifyOperation,
  type PresentationNativeModifyBatch,
} from '../src/skills/powerpoint/presentation-existing-batch'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
async function fixture(pageCount = 2) {
  const dir = mkdtempSync(join(tmpdir(), 'native-modify-'))
  roots.push(dir)
  const service = createPresentationService({ userDataPath: dir })
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const { bytes } = await compilePresentationDeck(deck),
    zip = await JSZip.loadAsync(bytes)
  let documentId = 'doc',
    available = true,
    failBackup = false,
    disconnectRead = false,
    switchRead = false,
    failReceipt = false
  let failSecondIntent = false
  let afterWrite: () => void = () => {}
  const shape = {
    id: 'sdk-id',
    name: 'title',
    type: 'GeometricShape',
    left: 1,
    top: 1,
    width: 100,
    height: 50,
  }
  const shapes = [
    structuredClone(shape),
    { ...shape, id: 'graphic-sdk', type: 'Picture' },
    ...Array.from({ length: 32 }, (_, i) => ({ ...shape, id: `sdk-${i}` })),
  ]
  let text = 'old',
    counter = 0
  const packages = new Map<string, string>()
  const packageFor = async (slideId: string) => {
    zip.file(
      'ppt/slides/slide1.xml',
      (await zip.file('ppt/slides/slide1.xml')!.async('string')).replace(
        '</p:spTree>',
        `<p:extLst><p:ext uri="${counter}-${slideId}"/></p:extLst></p:spTree>`,
      ),
    )
    const base64 = await zip.generateAsync({ type: 'base64' })
    packages.set(slideId, base64)
    return base64
  }
  await packageFor('s1')
  await packageFor('s2')
  for (let i = 2; i < pageCount; i++) packages.set(`s${i + 1}`, packages.get('s1')!)
  const slideIds = Array.from({ length: pageCount }, (_, i) => `s${i + 1}`)
  const originalPackages = new Map(packages)
  const settingsValues = new Map<string, string>()
  const settings = {
    get: (k: string) => settingsValues.get(k),
    set: (k: string, v: string) => settingsValues.set(k, v),
    save: vi.fn(async () => {}),
    location: () => documentId,
  }
  let binding = createPresentationDocumentBinding(settings, () => documentId)
  const request = vi.fn(async (body: any, signal?: AbortSignal) => {
    if (failBackup) throw Error('backup_failed')
    const response = new Response(
      Buffer.from(await service(body, signal ?? new AbortController().signal)).toString('utf8'),
    )
    if (body.operation === 'existing_page_backup_read') {
      if (disconnectRead) available = false
      if (switchRead) documentId = 'other'
    }
    return response
  })
  const png =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII='
  const adapter = {
    inspectPresentationPage: vi.fn(async (slideId: string) => ({
      slideId,
      shapes: [],
      shapesTruncated: false,
      screenshot: { mime: 'image/png', base64: png },
    })),
    snapshotSlide: vi.fn(async (index: number) => ({
      slideId: `s${index + 1}`,
      fingerprint: 'old',
    })),
    exportPresentationPagePackage: vi.fn(async (slideId: string) => ({
      slideId,
      slideIds,
      base64: packages.get(slideId)!,
    })),
    listSlideShapes: vi.fn(async (index: number) => ({
      slideId: `s${index + 1}`,
      slideIndex: index,
      shapes: structuredClone(shapes),
    })),
    readSlideText: vi.fn(async (index: number, id: string) => ({
      slideId: `s${index + 1}`,
      shapeId: id,
      text,
      paragraphs: [text],
    })),
    executeDeclarative: vi.fn(async (ops: NativeModifyOperation[]) => {
      for (const op of ops) {
        if (op.op === 'set_shape_text') text = op.text
        else if (op.op === 'delete_shape')
          shapes.splice(
            shapes.findIndex((s) => s.id === op.shape_id),
            1,
          )
        else
          Object.assign(
            shapes.find((s) => s.id === op.shape_id)!,
            { left: op.left, top: op.top, width: op.width, height: op.height },
          )
        counter++
        await packageFor(`s${op.slide_index + 1}`)
      }
      afterWrite()
      return { createdShapeIds: [] }
    }),
  }
  const proposals = createStructuredProposalController()
  const write = vi.fn(async (next: any, expected: any) => {
    if (failSecondIntent && next.inFlightIndex === 1) throw Error('intent_save_failed')
    if (failReceipt && next.nextIndex > 0) throw Error('ack_lost')
    await binding.writeExistingBatch(next, expected)
  })
  const make = () =>
    createPresentationNativeModifySkill({
      documentId: () => binding.documentId(),
      available: () => available,
      adapter: adapter as any,
      request,
      proposals,
      readExistingBatch: (id) => binding.readExistingBatch(id),
      writeExistingBatch: write,
    })
  let skill = make()
  return {
    restoreReceipt: async (id: string) => {
      const saved = binding.readExistingBatch(id) as PresentationNativeModifyBatch
      const undoing = { ...saved, state: 'undoing' as const }
      await binding.writeExistingBatch(undoing, saved)
      const restoredSlideIds = Object.fromEntries(
        saved.pages.map((page) => [page.hostSlideId, `restored-${page.hostSlideId}`]),
      )
      for (const page of saved.pages) {
        packages.set(restoredSlideIds[page.hostSlideId], originalPackages.get(page.hostSlideId)!)
        slideIds[page.slideIndex] = restoredSlideIds[page.hostSlideId]
      }
      await binding.writeExistingBatch({ ...undoing, state: 'undone', restoredSlideIds }, undoing)
    },
    corruptRestored: () => packages.set('restored-s1', packages.get('s1')!),
    adapter,
    proposals,
    request,
    write,
    propose: (ops: NativeModifyOperation[]) => skill.propose(ops),
    saved: (id: string) => binding.readExistingBatch(id) as PresentationNativeModifyBatch,
    reopen: () => {
      binding = createPresentationDocumentBinding(settings, () => documentId)
      skill = make()
    },
    callInput: (name: string, input: Record<string, unknown>) =>
      skill.executeTool({ id: 'review', name, input }),
    call: (name: string, id: string) =>
      skill.executeTool({ id: 'recovery', name, input: { change_id: id } }),
    setAvailable: (v: boolean) => (available = v),
    setFailBackup: () => (failBackup = true),
    setDisconnectRead: () => (disconnectRead = true),
    setSwitchRead: () => (switchRead = true),
    setFailReceipt: () => (failReceipt = true),
    setFailSecondIntent: () => (failSecondIntent = true),
    clearFailSecondIntent: () => (failSecondIntent = false),
    setAfterWrite: (fn: () => void) => (afterWrite = fn),
  }
}
const textOp: NativeModifyOperation = {
  op: 'set_shape_text',
  slide_index: 0,
  shape_id: 'sdk-id',
  text: 'new',
}
const geometryOp: NativeModifyOperation = {
  op: 'set_shape_geometry',
  slide_index: 1,
  shape_id: 'graphic-sdk',
  left: 4,
  top: 5,
  width: 60,
  height: 70,
}
it('refuses a disconnected durable binding before reads and writes', async () => {
  const f = await fixture()
  f.setAvailable(false)
  await expect(f.propose([textOp])).rejects.toThrow('presentation_existing_persistence_unavailable')
  expect(f.adapter.snapshotSlide).not.toHaveBeenCalled()
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('backs up single native text before journaling and writing; persisted result survives reopening', async () => {
  const f = await fixture()
  const p = await f.propose([textOp])
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  await f.proposals.confirm(p.proposalId)
  expect(f.saved(p.changeId)).toMatchObject({
    version: 3,
    state: 'applied',
    nextIndex: 1,
    operations: [textOp],
  })
  expect(f.saved(p.changeId).backups).toHaveLength(1)
  f.reopen()
  const result = await f.call('inspect_native_modify_batch', p.changeId)
  expect(result.isError).toBeUndefined()
  expect(JSON.parse(result.output)).toMatchObject({ state: 'applied', nextIndex: 1 })
})
it('preserves ordered native IDs, both pages, geometry on graphics, and immutable caller operations', async () => {
  const f = await fixture(),
    ops = [{ ...textOp }, { ...geometryOp }]
  const pending = f.propose(ops)
  ops[0].shape_id = 'mutated'
  const p = await pending
  await f.proposals.confirm(p.proposalId)
  expect(f.adapter.executeDeclarative.mock.calls.map(([ops]) => ops[0])).toEqual([
    textOp,
    geometryOp,
  ])
  expect(f.saved(p.changeId).backups).toHaveLength(2)
  expect(f.saved(p.changeId).state).toBe('applied')
})
it('retains 32-operation bounds and text noops', async () => {
  const f = await fixture()
  const ops = Array.from({ length: 32 }, (_, i) => ({
    ...textOp,
    shape_id: `sdk-${i}`,
    text: i === 0 ? 'old' : String(i),
  }))
  const p = await f.propose(ops)
  await f.proposals.confirm(p.proposalId)
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(32)
  expect(f.saved(p.changeId).nextIndex).toBe(32)
})
it('backup failure prevents durable intent and any host write', async () => {
  const f = await fixture()
  const p = await f.propose([textOp])
  f.setFailBackup()
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow('backup_failed')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  expect(f.saved(p.changeId)).toBeUndefined()
})
it('releases both original page backups when the first intent write fails', async () => {
  const f = await fixture()
  const p = await f.propose([textOp, geometryOp])
  f.write.mockRejectedValueOnce(Error('settings_failed'))
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow('settings_failed')
  expect(f.saved(p.changeId)).toBeUndefined()
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  const releases = f.request.mock.calls.filter(
    ([body]) => body.operation === 'existing_page_backup_release',
  )
  expect(releases).toHaveLength(2)
  for (const [body] of releases)
    expect(
      await (
        await f.request({
          operation: 'existing_page_backup_status',
          documentId: body.documentId,
          backupId: body.backupId,
        })
      ).json(),
    ).toHaveProperty('error')
})
it('releases the first page backup when the second upload fails', async () => {
  const f = await fixture()
  const p = await f.propose([textOp, geometryOp])
  const request = f.request.getMockImplementation()!
  let begins = 0
  f.request.mockImplementation(async (body, signal) => {
    if (body.operation === 'existing_page_backup_begin' && ++begins === 2)
      throw Error('second_upload_failed')
    return request(body, signal)
  })
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow('second_upload_failed')
  expect(f.saved(p.changeId)).toBeUndefined()
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  expect(
    f.request.mock.calls.filter(([body]) => body.operation === 'existing_page_backup_release'),
  ).toHaveLength(1)
})
it('abandons an incomplete original page upload after a chunk failure', async () => {
  const f = await fixture()
  const p = await f.propose([textOp])
  const request = f.request.getMockImplementation()!
  f.request.mockImplementation(async (body, signal) => {
    if (body.operation === 'existing_page_backup_chunk') throw Error('chunk_failed')
    return request(body, signal)
  })
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow('chunk_failed')
  expect(f.saved(p.changeId)).toBeUndefined()
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  const abandon = f.request.mock.calls.find(
    ([body]) => body.operation === 'existing_page_backup_abandon',
  )?.[0]
  expect(abandon).toBeDefined()
  expect(
    await (
      await f.request({
        operation: 'existing_page_backup_status',
        documentId: abandon.documentId,
        backupId: abandon.backupId,
      })
    ).json(),
  ).toHaveProperty('error', 'not_found')
})
it('keeps both backups when the first intent write commits but loses its ACK', async () => {
  const f = await fixture()
  const p = await f.propose([textOp, geometryOp])
  const write = f.write.getMockImplementation()!
  f.write.mockImplementationOnce(async (next, expected) => {
    await write(next, expected)
    throw Error('settings_ack_lost')
  })
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow('settings_ack_lost')
  expect(f.saved(p.changeId).state).toBe('applying')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  expect(
    f.request.mock.calls.filter(([body]) => body.operation === 'existing_page_backup_release'),
  ).toHaveLength(0)
})
it.each(['disconnect', 'switch'])(
  'rechecks %s after final backup read before first write',
  async (kind) => {
    const f = await fixture()
    const p = await f.propose([textOp])
    if (kind === 'disconnect') f.setDisconnectRead()
    else f.setSwitchRead()
    await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow(
      kind === 'disconnect'
        ? 'presentation_existing_persistence_unavailable'
        : 'presentation_document_changed',
    )
    expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  },
)
it('lost ACK retains uncertain flight across reopening and refuses replay', async () => {
  const f = await fixture()
  const p = await f.propose([textOp, geometryOp])
  f.setFailReceipt()
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow('ack_lost')
  expect(f.saved(p.changeId)).toMatchObject({ state: 'applying', nextIndex: 0, inFlightIndex: 0 })
  f.reopen()
  const inspected = await f.call('inspect_native_modify_batch', p.changeId)
  expect(JSON.parse(inspected.output).recovery).toBe('original_page_restoration_required')
  const resumed = await f.call('resume_native_modify_batch', p.changeId)
  expect(resumed.isError).toBe(true)
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(1)
})
it('does not acknowledge a native text edit that also moves an unrelated shape', async () => {
  const f = await fixture()
  const before = f.adapter.listSlideShapes.getMockImplementation()!
  f.adapter.listSlideShapes.mockImplementation(async (index: number) => {
    const listed = await before(index)
    if (f.adapter.executeDeclarative.mock.calls.length)
      listed.shapes.find((shape: { id: string }) => shape.id === 'graphic-sdk')!.left += 1
    return listed
  })
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({
    state: 'applying',
    nextIndex: 0,
    inFlightIndex: 0,
  })
})
it.each([
  geometryOp,
  { op: 'delete_shape', slide_index: 0, shape_id: 'sdk-id' } as NativeModifyOperation,
])('does not acknowledge %s when another shape changes', async (operation) => {
  const f = await fixture()
  const before = f.adapter.listSlideShapes.getMockImplementation()!
  f.adapter.listSlideShapes.mockImplementation(async (index: number) => {
    const listed = await before(index)
    if (f.adapter.executeDeclarative.mock.calls.length)
      listed.shapes.find((shape: { id: string }) => shape.id === 'sdk-0')!.top += 1
    return listed
  })
  const proposed = await f.propose([operation])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId).inFlightIndex).toBe(0)
})
it('disconnect after a native write retains uncertain receipt and does not continue', async () => {
  const f = await fixture()
  const p = await f.propose([textOp, geometryOp])
  f.setAfterWrite(() => f.setAvailable(false))
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow(
    'presentation_existing_persistence_unavailable',
  )
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(1)
  expect(f.saved(p.changeId).inFlightIndex).toBe(0)
})

it('allows offline receipt inspection but refuses disconnected continuation host writes', async () => {
  const f = await fixture()
  const p = await f.propose([textOp, geometryOp])
  f.setFailSecondIntent()
  await expect(f.proposals.confirm(p.proposalId)).rejects.toThrow('intent_save_failed')
  expect(f.saved(p.changeId)).toMatchObject({ nextIndex: 1, state: 'applying' })
  expect(f.saved(p.changeId).inFlightIndex).toBeUndefined()
  f.reopen()
  f.clearFailSecondIntent()
  const resumed = await f.call('resume_native_modify_batch', p.changeId)
  expect(resumed.isError).toBeUndefined()
  f.setAvailable(false)
  await expect(f.proposals.confirm(JSON.parse(resumed.output).proposalId)).rejects.toThrow()
  const inspected = await f.call('inspect_native_modify_batch', p.changeId)
  expect(inspected.isError).toBeUndefined()
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(1)
})

it('persists only fresh session page screenshot assessments after applied readback', async () => {
  const f = await fixture(),
    p = await f.propose([textOp])
  await f.proposals.confirm(p.proposalId)
  const capture = await f.callInput('capture_native_modify_page', {
    change_id: p.changeId,
    slide_id: 's1',
  })
  expect(capture.isError).toBeUndefined()
  const shot = JSON.parse(capture.output)
  expect(shot.qaPassed).toBe(false)
  const review = await f.callInput('record_native_modify_page_review', {
    change_id: p.changeId,
    slide_id: 's1',
    screenshot_digest: shot.screenshotDigest,
    status: 'pass',
    notes: 'Title readable',
  })
  expect(review.isError).toBeUndefined()
  expect(f.saved(p.changeId).reviews).toEqual([
    expect.objectContaining({ hostSlideId: 's1', status: 'pass', notes: 'Title readable' }),
  ])
  f.reopen()
  const stale = await f.callInput('record_native_modify_page_review', {
    change_id: p.changeId,
    slide_id: 's1',
    screenshot_digest: shot.screenshotDigest,
    status: 'pass',
    notes: 'Reused session',
  })
  expect(stale.isError).toBe(true)
})

it('refuses an invalid later target before any earlier host write or backup', async () => {
  const f = await fixture()
  await expect(
    f.propose([textOp, { ...geometryOp, shape_id: 'missing-sdk-target' }]),
  ).rejects.toThrow('office_read_failed')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  expect(f.request).not.toHaveBeenCalled()
})

// Eight real PC backups and per-step package proofs need room under the full-suite worker load.
it('keeps deck indices beyond 31 and eight distinct page scopes', async () => {
  const f = await fixture(50)
  const operations = Array.from({ length: 8 }, (_, i) => ({ ...textOp, slide_index: 40 + i }))
  const p = await f.propose(operations)
  await f.proposals.confirm(p.proposalId)
  expect(f.saved(p.changeId).backups).toHaveLength(8)
  expect(f.saved(p.changeId).beforeSlideIds).toHaveLength(50)
  expect(f.adapter.executeDeclarative.mock.calls.map(([ops]) => ops[0].slide_index)).toEqual([
    40, 41, 42, 43, 44, 45, 46, 47,
  ])
}, 15_000)

it('snapshots capture scope before asynchronous package guards', async () => {
  const f = await fixture(),
    p = await f.propose([textOp])
  await f.proposals.confirm(p.proposalId)
  const original = f.adapter.exportPresentationPagePackage.getMockImplementation()!
  let entered!: () => void, release!: () => void
  const started = new Promise<void>((resolve) => (entered = resolve)),
    gate = new Promise<void>((resolve) => (release = resolve))
  f.adapter.exportPresentationPagePackage.mockImplementationOnce(async (slideId) => {
    entered()
    await gate
    return original(slideId)
  })
  const input = { change_id: p.changeId, slide_id: 's1' }
  const pending = f.callInput('capture_native_modify_page', input)
  await started
  input.slide_id = 'unrelated'
  release()
  const result = await pending
  expect(result.isError).toBeUndefined()
  expect(f.adapter.inspectPresentationPage).toHaveBeenLastCalledWith('s1', undefined)
  expect(JSON.parse(result.output).slideId).toBe('s1')
})

it('preserves the browser export deck-order limit of 512 pages', async () => {
  const f = await fixture(512),
    p = await f.propose([{ ...textOp, slide_index: 511 }])
  await f.proposals.confirm(p.proposalId)
  expect(f.saved(p.changeId).beforeSlideIds).toHaveLength(512)
  expect(f.saved(p.changeId).pages[0].slideIndex).toBe(511)
})
it('waits for delayed native text readback before saving a successful receipt', async () => {
  const f = await fixture(),
    original = f.adapter.readSlideText.getMockImplementation()!
  let afterReads = 0
  f.adapter.readSlideText.mockImplementation(async (index, id) => {
    const value = await original(index, id)
    if (f.adapter.executeDeclarative.mock.calls.length && afterReads++ === 0)
      return { ...value, text: 'old', paragraphs: ['old'] }
    return value
  })
  const p = await f.propose([textOp])
  await expect(f.proposals.confirm(p.proposalId)).resolves.toBeUndefined()
  expect(afterReads).toBeGreaterThanOrEqual(2)
  expect(f.saved(p.changeId)).toMatchObject({ state: 'applied', nextIndex: 1 })
})

it('requires own restored identities for every undone page including prototype-like host IDs', async () => {
  const f = await fixture(),
    p = await f.propose([textOp])
  await f.proposals.confirm(p.proposalId)
  const saved = f.saved(p.changeId)
  const undone = {
    ...saved,
    state: 'undone' as const,
    beforeSlideIds: ['constructor', 's2'],
    scope: { slideIds: ['constructor'] },
    pages: saved.pages.map((page) => ({ ...page, hostSlideId: 'constructor' })),
    backups: saved.backups.map((backup) => ({ ...backup, hostSlideId: 'constructor' })),
    restoredSlideIds: {},
  }
  expect(validatePresentationExistingBatch(undone)).toBe(false)
  expect(
    validatePresentationExistingBatch({
      ...undone,
      restoredSlideIds: Object.fromEntries([['constructor', 'restored-native-id']]),
    }),
  ).toBe(true)
})

it('inspects undone pages through restored SDK identities and original package digests', async () => {
  const f = await fixture(),
    p = await f.propose([textOp])
  await f.proposals.confirm(p.proposalId)
  await f.restoreReceipt(p.changeId)
  f.reopen()
  const inspected = await f.call('inspect_native_modify_batch', p.changeId)
  expect(inspected.isError).toBeUndefined()
  expect(JSON.parse(inspected.output)).toMatchObject({
    state: 'undone',
    currentHostVerified: true,
    qaPassed: false,
  })
  expect(f.adapter.exportPresentationPagePackage).toHaveBeenLastCalledWith('restored-s1', undefined)
  f.corruptRestored()
  const changed = await f.call('inspect_native_modify_batch', p.changeId)
  expect(changed.isError).toBeUndefined()
  expect(JSON.parse(changed.output)).toMatchObject({
    currentPackageMatches: false,
    pages: [expect.objectContaining({ status: 'changed' })],
  })
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(1)
})
