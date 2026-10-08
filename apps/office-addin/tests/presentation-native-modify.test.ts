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
import { presentationPackageDigest } from '../src/skills/powerpoint/powerpoint-package'
import {
  validatePresentationExistingBatch,
  type NativeModifyOperation,
  type PresentationNativeModifyBatch,
} from '../src/skills/powerpoint/presentation-existing-batch'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
async function fixture(pageCount = 2, extraShapeCount = 32) {
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
    { ...shape, id: 'table-sdk', type: 'Table' },
    { ...shape, id: 'image-sdk', type: 'Image' },
    { ...shape, id: 'line-sdk', type: 'Line' },
    { ...shape, id: 'group-sdk', type: 'Group' },
    { ...shape, id: 'callout-sdk', type: 'Callout' },
    { ...shape, id: 'freeform-sdk', type: 'Freeform' },
    { ...shape, id: 'smartart-sdk', type: 'SmartArt' },
    { ...shape, id: 'ole-sdk', type: 'Ole' },
    ...Array.from({ length: extraShapeCount }, (_, i) => ({ ...shape, id: `sdk-${i}` })),
  ]
  const texts = new Map<string, string>()
  let tableCell = '10'
  let imageAlt = 'Original source'
  let imageMedia = 'a'.repeat(64)
  let textFont = 'Arial'
  let tableStyle = 'a'.repeat(64)
  let chartFingerprint = 'd'.repeat(64)
  let textStructure = 'a'.repeat(64)
  let emptyTextStructure = 'e'.repeat(64)
  let ordinaryStructure = '7'.repeat(64)
  let ordinaryDriftId = 'sdk-1'
  let counter = 0
  const textFor = (id: string) => texts.get(id) ?? 'old'
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
      shapes: shapes.map((item) => ({
        ...structuredClone(item),
        text: ['Picture', 'Image', 'Chart'].includes(item.type) ? '' : textFor(item.id),
        ...(item.type === 'Table' ? { tableValues: [['Revenue', tableCell]] } : {}),
        ...(item.type === 'Image'
          ? { rotation: 0, altTextTitle: 'Figure', altTextDescription: imageAlt }
          : {}),
      })),
    })),
    exportPresentationPagePackage: vi.fn(async (slideId: string) => ({
      slideId,
      slideIds,
      base64: packages.get(slideId)!,
    })),
    inspectSlidePictureFingerprints: vi.fn(async (slideId: string, shapeIds: string[]) => ({
      slideId,
      slideIds,
      fingerprints: Object.fromEntries(
        shapeIds.map((id) => [id, id === 'image-sdk' ? imageMedia : 'c'.repeat(64)]),
      ),
      mediaDigests: Object.fromEntries(
        shapeIds.map((id) => [id, id === 'image-sdk' ? imageMedia : 'c'.repeat(64)]),
      ),
    })),
    inspectSlideRichText: vi.fn(async (slideId: string, shapeIds: string[]) => ({
      slideId,
      slideIds,
      fingerprints: Object.fromEntries(
        shapeIds.map((id) => [
          id,
          {
            content:
              id === 'sdk-0'
                ? textStructure
                : id === 'empty-sdk'
                  ? emptyTextStructure
                  : 'b'.repeat(64),
            formatting:
              id === 'sdk-0'
                ? textStructure
                : id === 'empty-sdk'
                  ? emptyTextStructure
                  : 'b'.repeat(64),
          },
        ]),
      ),
      shapes: Object.fromEntries(
        shapeIds.map((id) => [
          id,
          {
            packageShapeId: id,
            name: id,
            paragraphs: [
              {
                runs: [
                  {
                    text: textFor(id),
                    directFont: { typeface: id === 'sdk-0' ? textFont : 'Arial' },
                  },
                ],
              },
            ],
          },
        ]),
      ),
    })),
    inspectSlideTableFingerprints: vi.fn(async (slideId: string, shapeIds: string[]) => ({
      slideId,
      slideIds,
      fingerprints: Object.fromEntries(shapeIds.map((id) => [id, tableStyle])),
    })),
    inspectSlideChartFingerprints: vi.fn(async (slideId: string, shapeIds: string[]) => ({
      slideId,
      slideIds,
      fingerprints: Object.fromEntries(shapeIds.map((id) => [id, chartFingerprint])),
    })),
    listSlideShapes: vi.fn(async (index: number) => ({
      slideId: `s${index + 1}`,
      slideIndex: index,
      shapes: structuredClone(shapes),
    })),
    readSlideText: vi.fn(async (index: number, id: string) => ({
      slideId: `s${index + 1}`,
      shapeId: id,
      text: textFor(id),
      paragraphs: [textFor(id)],
    })),
    executeDeclarative: vi.fn(async (ops: NativeModifyOperation[]) => {
      for (const op of ops) {
        if (op.op === 'set_shape_text') texts.set(op.shape_id, op.text)
        else if (op.op === 'delete_shape')
          shapes.splice(
            shapes.findIndex((s) => s.id === op.shape_id),
            1,
          )
        else if (op.op === 'set_shape_geometry')
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
    enableCombinedProof: () => {
      const picture = adapter.inspectSlidePictureFingerprints.getMockImplementation()!
      const richText = adapter.inspectSlideRichText.getMockImplementation()!
      const table = adapter.inspectSlideTableFingerprints.getMockImplementation()!
      const chart = adapter.inspectSlideChartFingerprints.getMockImplementation()!
      const combined = vi.fn(
        async (
          slideId: string,
          ids: {
            pictures: string[]
            text: string[]
            tables: string[]
            charts: string[]
            ordinary: string[]
          },
        ) => {
          const exported = await adapter.exportPresentationPagePackage(slideId)
          const pictures = ids.pictures.length
            ? await picture(slideId, ids.pictures)
            : { fingerprints: {}, mediaDigests: {} }
          const text = ids.text.length
            ? await richText(slideId, ids.text)
            : { shapes: {}, fingerprints: {} }
          return {
            ...exported,
            pictures: { fingerprints: pictures.fingerprints, mediaDigests: pictures.mediaDigests },
            richText: { shapes: text.shapes, fingerprints: text.fingerprints },
            tables: ids.tables.length ? (await table(slideId, ids.tables)).fingerprints : {},
            charts: ids.charts.length ? (await chart(slideId, ids.charts)).fingerprints : {},
            ordinary: Object.fromEntries(
              ids.ordinary.map((id) => [
                id,
                {
                  exact: id === ordinaryDriftId ? ordinaryStructure : '8'.repeat(64),
                  content: id === ordinaryDriftId ? ordinaryStructure : '8'.repeat(64),
                  formatting: id === ordinaryDriftId ? ordinaryStructure : '8'.repeat(64),
                },
              ]),
            ),
          }
        },
      )
      Object.assign(adapter, { inspectSlideNativePackage: combined })
      adapter.inspectSlidePictureFingerprints.mockRejectedValue(Error('separate_read'))
      adapter.inspectSlideRichText.mockRejectedValue(Error('separate_read'))
      adapter.inspectSlideTableFingerprints.mockRejectedValue(Error('separate_read'))
      adapter.inspectSlideChartFingerprints.mockRejectedValue(Error('separate_read'))
      return combined
    },
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
    setTargetText: (value: string) => texts.set('sdk-id', value),
    setOtherText: (value: string) => texts.set('sdk-0', value),
    setTableCell: (value: string) => (tableCell = value),
    setImageAlt: (value: string) => (imageAlt = value),
    setImageMedia: (value: string) => (imageMedia = value),
    setTextFont: (value: string) => (textFont = value),
    setTextStructure: (value: string) => (textStructure = value),
    addEmptyTextBox: () => {
      shapes.push({ ...shape, id: 'empty-sdk', type: 'TextBox' })
      texts.set('empty-sdk', '')
    },
    setEmptyTextStructure: (value: string) => (emptyTextStructure = value),
    setOrdinaryStructure: (value: string) => (ordinaryStructure = value),
    setOrdinaryDriftId: (value: string) => (ordinaryDriftId = value),
    setTableStyle: (value: string) => (tableStyle = value),
    addChart: () => shapes.push({ ...shape, id: 'chart-sdk', type: 'Chart' }),
    setChartFingerprint: (value: string) => (chartFingerprint = value),
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
const imageGeometryOp: NativeModifyOperation = {
  ...geometryOp,
  shape_id: 'image-sdk',
}
it('refuses a disconnected durable binding before reads and writes', async () => {
  const f = await fixture()
  f.setAvailable(false)
  await expect(f.propose([textOp])).rejects.toThrow('presentation_existing_persistence_unavailable')
  expect(f.adapter.snapshotSlide).not.toHaveBeenCalled()
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('uses one coherent package proof on each side of a native write and acknowledges that exact package', async () => {
  const f = await fixture()
  f.addChart()
  const combined = f.enableCombinedProof()
  const proposed = await f.propose([textOp])
  await f.proposals.confirm(proposed.proposalId)
  expect(combined).toHaveBeenCalledTimes(2)
  expect(combined.mock.calls[0]![1].ordinary).toEqual(
    expect.arrayContaining([
      'line-sdk',
      'group-sdk',
      'callout-sdk',
      'freeform-sdk',
      'smartart-sdk',
      'ole-sdk',
    ]),
  )
  const after = await combined.mock.results[1]!.value
  expect(f.saved(proposed.changeId).pages[0]!.expectedPackageDigest).toBe(
    await presentationPackageDigest(after.base64),
  )
  expect(f.adapter.inspectSlidePictureFingerprints).not.toHaveBeenCalled()
  expect(f.adapter.inspectSlideRichText).not.toHaveBeenCalled()
  expect(f.adapter.inspectSlideTableFingerprints).not.toHaveBeenCalled()
  expect(f.adapter.inspectSlideChartFingerprints).not.toHaveBeenCalled()
})
it('keeps a native modify transaction available on pages with more than 100 shapes', async () => {
  const f = await fixture(2, 120)
  const combined = f.enableCombinedProof()
  const proposed = await f.propose([textOp])
  await f.proposals.confirm(proposed.proposalId)
  expect(combined.mock.calls[0]![1].ordinary.length).toBeGreaterThan(100)
  expect(f.saved(proposed.changeId).state).toBe('applied')
})
it('refuses a native write when a combined package proof omits a required shape', async () => {
  const f = await fixture()
  const combined = f.enableCombinedProof()
  const inspect = combined.getMockImplementation()!
  combined.mockImplementation(async (slideId, ids) => {
    const proof = await inspect(slideId, ids)
    return { ...proof, richText: { ...proof.richText, fingerprints: {} } }
  })
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_read_failed')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  expect(f.saved(proposed.changeId).inFlightIndex).toBeUndefined()
})
it('keeps a native write uncertain when the combined package readback is incomplete', async () => {
  const f = await fixture()
  const combined = f.enableCombinedProof()
  const inspect = combined.getMockImplementation()!
  combined.mockImplementation(async (slideId, ids) => {
    const proof = await inspect(slideId, ids)
    return f.adapter.executeDeclarative.mock.calls.length
      ? { ...proof, richText: { ...proof.richText, fingerprints: {} } }
      : proof
  })
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_read_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('does not acknowledge a hidden geometric shape style change in the combined package', async () => {
  const f = await fixture()
  f.enableCombinedProof()
  f.setAfterWrite(() => f.setOrdinaryStructure('9'.repeat(64)))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('does not acknowledge a hidden SmartArt package change in an unrelated shape', async () => {
  const f = await fixture()
  f.setOrdinaryDriftId('smartart-sdk')
  f.enableCombinedProof()
  f.setAfterWrite(() => f.setOrdinaryStructure('9'.repeat(64)))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
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
it('does not acknowledge a native text edit that also changes unrelated text', async () => {
  const f = await fixture()
  f.setAfterWrite(() => f.setOtherText('changed without a matching operation'))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({
    state: 'applying',
    nextIndex: 0,
    inFlightIndex: 0,
  })
})
it('does not acknowledge a native text edit that also changes an unrelated table cell', async () => {
  const f = await fixture()
  f.setAfterWrite(() => f.setTableCell('11'))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('does not acknowledge a native text edit that also changes unrelated table formatting', async () => {
  const f = await fixture()
  f.setAfterWrite(() => f.setTableStyle('b'.repeat(64)))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('refuses a native write when an existing table cannot be mapped to its package', async () => {
  const f = await fixture()
  f.adapter.inspectSlideTableFingerprints.mockRejectedValue(Error('office_api_unsupported'))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_api_unsupported')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  expect(f.saved(proposed.changeId).inFlightIndex).toBeUndefined()
})
it('keeps a native write uncertain when table formatting cannot be read back', async () => {
  const f = await fixture()
  const inspect = f.adapter.inspectSlideTableFingerprints.getMockImplementation()!
  f.adapter.inspectSlideTableFingerprints.mockImplementation(async (slideId, shapeIds) => {
    if (f.adapter.executeDeclarative.mock.calls.length) throw Error('office_api_unsupported')
    return inspect(slideId, shapeIds)
  })
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_api_unsupported')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('does not acknowledge an unrelated chart package change during a native text edit', async () => {
  const f = await fixture()
  f.addChart()
  f.setAfterWrite(() => f.setChartFingerprint('e'.repeat(64)))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('refuses a native write when an existing chart cannot be mapped to its package', async () => {
  const f = await fixture()
  f.addChart()
  f.adapter.inspectSlideChartFingerprints.mockRejectedValue(Error('office_api_unsupported'))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_api_unsupported')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  expect(f.saved(proposed.changeId).inFlightIndex).toBeUndefined()
})
it('keeps a native write uncertain when chart package readback fails', async () => {
  const f = await fixture()
  f.addChart()
  const inspect = f.adapter.inspectSlideChartFingerprints.getMockImplementation()!
  f.adapter.inspectSlideChartFingerprints.mockImplementation(async (slideId, shapeIds) => {
    if (f.adapter.executeDeclarative.mock.calls.length) throw Error('office_api_unsupported')
    return inspect(slideId, shapeIds)
  })
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_api_unsupported')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('does not acknowledge a native text edit that also changes unrelated image attribution', async () => {
  const f = await fixture()
  f.setAfterWrite(() => f.setImageAlt('Changed source'))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('does not acknowledge an image geometry edit that also changes its attribution', async () => {
  const f = await fixture()
  f.setAfterWrite(() => f.setImageAlt('Changed source'))
  const proposed = await f.propose([imageGeometryOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('does not acknowledge an image geometry edit that also changes its media', async () => {
  const f = await fixture()
  f.setAfterWrite(() => f.setImageMedia('b'.repeat(64)))
  const proposed = await f.propose([imageGeometryOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('does not acknowledge a native text edit that also changes unrelated image media', async () => {
  const f = await fixture()
  f.setAfterWrite(() => f.setImageMedia('b'.repeat(64)))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('does not acknowledge an unrelated rich-text formatting change hidden from Office.js text snapshots', async () => {
  const f = await fixture()
  f.setAfterWrite(() => f.setTextFont('Georgia'))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('does not acknowledge an unrelated text-shape XML change omitted by the parsed font view', async () => {
  const f = await fixture()
  f.setAfterWrite(() => f.setTextStructure('c'.repeat(64)))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('protects an empty text box from formatting drift during another native edit', async () => {
  const f = await fixture()
  f.addEmptyTextBox()
  f.setAfterWrite(() => f.setEmptyTextStructure('f'.repeat(64)))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.adapter.inspectSlideRichText).toHaveBeenCalledWith(
    's1',
    expect.arrayContaining(['empty-sdk']),
    expect.any(AbortSignal),
  )
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('refuses a native write when rich text cannot be mapped to its package', async () => {
  const f = await fixture()
  f.adapter.inspectSlideRichText.mockRejectedValue(Error('office_api_unsupported'))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_api_unsupported')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  expect(f.saved(proposed.changeId).inFlightIndex).toBeUndefined()
})
it('refuses a native write when an existing picture cannot be mapped to its package', async () => {
  const f = await fixture()
  f.adapter.inspectSlidePictureFingerprints.mockRejectedValue(Error('office_api_unsupported'))
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_api_unsupported')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  expect(f.saved(proposed.changeId)).toMatchObject({ nextIndex: 0, state: 'applying' })
  expect(f.saved(proposed.changeId).inFlightIndex).toBeUndefined()
})
it('keeps the write uncertain when a picture cannot be read back after the host write', async () => {
  const f = await fixture()
  const inspect = f.adapter.inspectSlidePictureFingerprints.getMockImplementation()!
  f.adapter.inspectSlidePictureFingerprints.mockImplementation(
    async (slideId: string, shapeIds: string[]) => {
      if (f.adapter.executeDeclarative.mock.calls.length) throw Error('office_api_unsupported')
      return inspect(slideId, shapeIds)
    },
  )
  const proposed = await f.propose([textOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_api_unsupported')
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(1)
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('does not acknowledge a native geometry edit that also changes target text', async () => {
  const f = await fixture()
  f.setAfterWrite(() => f.setTargetText('changed without a matching operation'))
  const proposed = await f.propose([geometryOp])
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
})
it('does not start a native write without a complete semantic shape snapshot', async () => {
  const f = await fixture()
  const proposed = await f.propose([textOp])
  f.adapter.snapshotSlide.mockImplementation(async (index: number) => ({
    slideId: `s${index + 1}`,
    fingerprint: 'old',
    shapes: undefined as never,
  }))
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_read_failed')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('does not start a native write when the shape list and semantic snapshot disagree', async () => {
  const f = await fixture()
  const proposed = await f.propose([textOp])
  const original = f.adapter.snapshotSlide.getMockImplementation()!
  f.adapter.snapshotSlide.mockImplementation(async (index: number) => {
    const snapshot = await original(index)
    snapshot.shapes.find((shape: { id: string }) => shape.id === 'sdk-0')!.left += 1
    return snapshot
  })
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_read_failed')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('does not acknowledge a write when its shape list and semantic snapshot disagree', async () => {
  const f = await fixture()
  const proposed = await f.propose([textOp])
  const original = f.adapter.snapshotSlide.getMockImplementation()!
  f.adapter.snapshotSlide.mockImplementation(async (index: number) => {
    const snapshot = await original(index)
    if (f.adapter.executeDeclarative.mock.calls.length)
      snapshot.shapes.find((shape: { id: string }) => shape.id === 'sdk-0')!.left += 1
    return snapshot
  })
  await expect(f.proposals.confirm(proposed.proposalId)).rejects.toThrow('office_verify_failed')
  expect(f.saved(proposed.changeId)).toMatchObject({ inFlightIndex: 0, nextIndex: 0 })
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
