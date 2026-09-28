import { afterEach, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { createOfficeHostRuntime } from '../src/agent/host-runtime'
import type { StructuredProposalController } from '../src/agent/proposal-controller'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { BrowserPresentationBaselineAdapter } from '../src/skills/powerpoint/browser-presentation-baseline-adapter'
import { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
const png =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII='
const otherPng =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII='
afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})
it('confirms a simple native table cell edit with a durable savepoint and reopens for undo', async () => {
  const f = await fixture()
  f.setShapeType('Table')
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_table_cell', {
    baseline_id,
    slide_id: 'slide',
    shape_id: 'shape',
    row_index: 0,
    column_index: 0,
    text: 'after',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  expect(f.tableText()).toBe('before')
  await f.confirm()
  expect(f.tableText()).toBe('after')
  expect(f.records()[0]!.record).toMatchObject({
    kind: 'table_cell',
    rowIndex: 0,
    columnIndex: 0,
    before: 'before',
    after: 'after',
    state: 'applied',
  })
  f.reopen()
  const change_id = f.records()[0]!.record.changeId
  const inspected = await f.call('inspect_existing_presentation_change', { change_id })
  expect(JSON.parse(inspected.output).status).toBe('not_pending')
  const undo = await f.call('undo_existing_presentation_change', { change_id })
  expect(undo.isError, undo.output).not.toBe(true)
  await f.confirm()
  expect(f.tableText()).toBe('before')
  expect(f.records()[0]!.record.state).toBe('undone')
  expect(f.editTableCell).toHaveBeenCalledTimes(2)
})
it('applies and reverses two ordered native table cells with one durable batch', async () => {
  const f = await fixture()
  f.setShapeType('Table')
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_table_batch', {
    baseline_id,
    intent: 'Update two figures',
    preserved: ['Other cells'],
    validation: ['Readback'],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', row_index: 0, column_index: 0, text: 'after' },
      { slide_id: 'slide', shape_id: 'shape', row_index: 0, column_index: 1, text: 'after-2' },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  expect(f.tableText()).toBe('before')
  expect(f.tableText2()).toBe('before-2')
  await f.confirm()
  expect([f.tableText(), f.tableText2()]).toEqual(['after', 'after-2'])
  const history = f.binding().listChangeHistory()
  const batch = history.find((entry) => entry.kind === 'existing_batch')
  expect(batch?.record).toMatchObject({
    state: 'applied',
    cursor: 2,
    operations: [{ kind: 'table_cell' }, { kind: 'table_cell' }],
  })
  f.reopen()
  await f.getRuntime().changes!.refresh()
  const workbench = f
    .getRuntime()
    .changes!.snapshot()
    .entries.find((entry) => entry.id === `existing_batch:${batch!.record.changeId}`)
  expect(workbench?.before).toContain('shape[0,0]')
  expect(workbench?.before).toContain('shape[0,1]')
  const undo = await f.call('undo_existing_presentation_batch', {
    change_id: batch!.record.changeId,
  })
  expect(undo.isError, undo.output).not.toBe(true)
  await f.confirm()
  expect([f.tableText(), f.tableText2()]).toEqual(['before', 'before-2'])
  expect(f.binding().readExistingBatch(batch!.record.changeId)?.state).toBe('undone')
})
it('resumes a table batch after one cell was durably written and the next write failed', async () => {
  const f = await fixture()
  f.setShapeType('Table')
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_table_batch', {
    baseline_id,
    intent: 'Update two figures',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', row_index: 0, column_index: 0, text: 'after' },
      { slide_id: 'slide', shape_id: 'shape', row_index: 0, column_index: 1, text: 'after-2' },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const change_id = JSON.parse(proposed.output).changeId as string
  const write = f.editTableCell.getMockImplementation()!
  f.editTableCell
    .mockImplementationOnce(write)
    .mockRejectedValueOnce(new Error('office_write_failed'))
  await expect(f.confirm()).rejects.toThrow()
  expect(f.binding().readExistingBatch(change_id)).toMatchObject({ state: 'applying', cursor: 1 })
  expect([f.tableText(), f.tableText2()]).toEqual(['after', 'before-2'])
  f.reopen()
  const inspected = await f.call('inspect_existing_presentation_batch', { change_id })
  expect(JSON.parse(inspected.output).values).toEqual(['after', 'before'])
  const resumed = await f.call('resume_existing_presentation_batch', { change_id })
  expect(resumed.isError, resumed.output).not.toBe(true)
  await f.confirm()
  expect([f.tableText(), f.tableText2()]).toEqual(['after', 'after-2'])
  expect(f.binding().readExistingBatch(change_id)).toMatchObject({ state: 'applied', cursor: 2 })
})

it('stops a confirmed batch after its first durable write when Stop is requested', async () => {
  const f = await fixture()
  f.setShapeType('Table')
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_table_batch', {
    baseline_id,
    intent: 'Update two figures',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', row_index: 0, column_index: 0, text: 'after' },
      { slide_id: 'slide', shape_id: 'shape', row_index: 0, column_index: 1, text: 'after-2' },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  const write = f.editTableCell.getMockImplementation()!
  f.editTableCell.mockImplementationOnce(async (...args) => {
    await write(...args)
    f.getRuntime().proposals.logout()
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.editTableCell).toHaveBeenCalledTimes(1)
  expect([f.tableText(), f.tableText2()]).toEqual(['after', 'before-2'])
  expect(f.binding().readExistingBatch(changeId)).toMatchObject({ state: 'applying', cursor: 1 })
  f.reopen()
  const resumed = await f.call('resume_existing_presentation_batch', { change_id: changeId })
  expect(resumed.isError, resumed.output).not.toBe(true)
  await f.confirm()
  expect([f.tableText(), f.tableText2()]).toEqual(['after', 'after-2'])
  expect(f.editTableCell).toHaveBeenCalledTimes(2)
})

it('stops a batch undo after its first durable reverse write and resumes it', async () => {
  const f = await fixture()
  f.setShapeType('Table')
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_table_batch', {
    baseline_id,
    intent: 'Update two figures',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', row_index: 0, column_index: 0, text: 'after' },
      { slide_id: 'slide', shape_id: 'shape', row_index: 0, column_index: 1, text: 'after-2' },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  await f.confirm()
  f.reopen()
  const undo = await f.call('undo_existing_presentation_batch', { change_id: changeId })
  expect(undo.isError, undo.output).not.toBe(true)
  f.editTableCell.mockClear()
  const write = f.editTableCell.getMockImplementation()!
  f.editTableCell.mockImplementationOnce(async (...args) => {
    await write(...args)
    f.getRuntime().proposals.logout()
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.editTableCell).toHaveBeenCalledTimes(1)
  expect([f.tableText(), f.tableText2()]).toEqual(['after', 'before-2'])
  expect(f.binding().readExistingBatch(changeId)).toMatchObject({ state: 'undoing', cursor: 1 })
  f.reopen()
  const resumed = await f.call('resume_existing_presentation_batch', { change_id: changeId })
  expect(resumed.isError, resumed.output).not.toBe(true)
  await f.confirm()
  expect([f.tableText(), f.tableText2()]).toEqual(['before', 'before-2'])
  expect(f.editTableCell).toHaveBeenCalledTimes(2)
  expect(f.binding().readExistingBatch(changeId)).toMatchObject({ state: 'undone', cursor: 0 })
})
it('finalizes an already written table cell after a lost receipt without replaying it', async () => {
  const f = await fixture()
  f.setShapeType('Table')
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_table_batch', {
    baseline_id,
    intent: 'Update two figures',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', row_index: 0, column_index: 0, text: 'after' },
      { slide_id: 'slide', shape_id: 'shape', row_index: 0, column_index: 1, text: 'after-2' },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const change_id = JSON.parse(proposed.output).changeId as string
  const write = f.editTableCell.getMockImplementation()!
  f.editTableCell.mockImplementationOnce(write).mockImplementationOnce(async (...args) => {
    await write(...args)
    throw new Error('receipt_lost')
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.binding().readExistingBatch(change_id)).toMatchObject({ state: 'applying', cursor: 1 })
  expect([f.tableText(), f.tableText2()]).toEqual(['after', 'after-2'])
  f.reopen()
  const resumed = await f.call('resume_existing_presentation_batch', { change_id })
  expect(resumed.isError, resumed.output).not.toBe(true)
  await f.confirm()
  expect(f.editTableCell).toHaveBeenCalledTimes(2)
  expect(f.binding().readExistingBatch(change_id)).toMatchObject({ state: 'applied', cursor: 2 })
})
it('rejects a table batch if a target cell changes to a third value before confirmation', async () => {
  const f = await fixture()
  f.setShapeType('Table')
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_table_batch', {
    baseline_id,
    intent: 'Update two figures',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', row_index: 0, column_index: 0, text: 'after' },
      { slide_id: 'slide', shape_id: 'shape', row_index: 0, column_index: 1, text: 'after-2' },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  f.setTableText2('manual')
  await expect(f.confirm()).rejects.toThrow()
  expect(f.editTableCell).not.toHaveBeenCalled()
  expect(f.binding().listChangeHistory()).toHaveLength(0)
})
it('rejects a table cell target changed after its baseline package read', async () => {
  const f = await fixture()
  f.setShapeType('Table')
  const baseline_id = await f.baseline()
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'readPresentationTableCell').mockResolvedValue({
    slideId: 'slide',
    shapeId: 'shape',
    rowIndex: 0,
    columnIndex: 0,
    text: 'manual',
    rowCount: 1,
    columnCount: 1,
  })
  const result = await f.call('edit_existing_presentation_table_cell', {
    baseline_id,
    slide_id: 'slide',
    shape_id: 'shape',
    row_index: 0,
    column_index: 0,
    text: 'after',
  })
  expect(result.output).toBe('presentation_baseline_changed')
  expect(f.records()).toHaveLength(0)
})
it('refuses table cell undo when the host cell becomes multi-run with the same text', async () => {
  const f = await fixture()
  f.setShapeType('Table')
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_table_cell', {
    baseline_id,
    slide_id: 'slide',
    shape_id: 'shape',
    row_index: 0,
    column_index: 0,
    text: 'after',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  await f.confirm()
  const zip = new JSZip()
  zip.file(
    'ppt/slides/slide1.xml',
    '<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="shape" name="Table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>af</a:t></a:r><a:r><a:t>ter</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>',
  )
  const base64 = await zip.generateAsync({ type: 'base64' })
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'exportPresentationPagePackage').mockResolvedValue({
    slideId: 'slide',
    slideIds: ['slide', 'other'],
    base64,
  })
  const change_id = f.records()[0]!.record.changeId
  const undo = await f.call('undo_existing_presentation_change', { change_id })
  expect(undo.output).toBe('presentation_existing_target_unsupported')
  expect(f.tableText()).toBe('after')
  expect(f.editTableCell).toHaveBeenCalledTimes(1)
})
it('returns unreviewed post-write evidence for a verified native text change', async () => {
  const f = await fixture()
  await f.propose()
  const proposal = f.getRuntime().proposals as StructuredProposalController
  const decision = proposal.waitForDecision(proposal.pending()!.id)
  await f.confirm()
  expect(await decision).toMatchObject({
    status: 'confirmed',
    postWrite: { status: 'captured', pages: [{ slideId: 'slide', pngBase64: png }] },
  })
})

it('keeps a verified text write successful when the post-write screenshot fails', async () => {
  const f = await fixture()
  await f.propose()
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'inspectPresentationPage').mockRejectedValueOnce(
    new Error('office_read_failed'),
  )
  const proposal = f.getRuntime().proposals as StructuredProposalController
  const decision = proposal.waitForDecision(proposal.pending()!.id)
  await f.confirm()
  expect(await decision).toMatchObject({
    status: 'confirmed',
    postWrite: { status: 'unavailable' },
  })
  expect(f.records()[0]!.record).toMatchObject({ state: 'applied' })
  expect(f.text()).toBe('after')
})

it('captures each exact affected native page after a verified batch', async () => {
  const f = await fixture()
  const all = await f.call('read_presentation_baseline', { scope: 'deck' })
  expect(all.isError, all.output).not.toBe(true)
  const baseline_id = JSON.parse(all.output).baselineId as string
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update two pages',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      { slide_id: 'other', shape_id: 'other-shape', kind: 'text', text: 'other-after' },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const proposal = f.getRuntime().proposals as StructuredProposalController
  const decision = proposal.waitForDecision(proposal.pending()!.id)
  await f.confirm()
  expect(await decision).toMatchObject({
    status: 'confirmed',
    postWrite: {
      status: 'captured',
      pages: [
        { slideId: 'slide', pngBase64: png },
        { slideId: 'other', pngBase64: png },
      ],
    },
  })
})
async function fixture() {
  vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } })
  let location = 'file://existing.pptx',
    text = 'before',
    tableText = 'before',
    tableText2 = 'before-2',
    otherText = 'other-before',
    geometry = { left: 1, top: 2, width: 100, height: 40 },
    shapeType = 'TextBox',
    font = {
      name: 'Arial' as string | null,
      size: 20 as number | null,
      color: '#000000' as string | null,
      bold: false as boolean | null,
      italic: false as boolean | null,
      underline: 'None' as string | null,
    },
    rangeFont = {
      name: 'Arial',
      size: 20,
      color: '#000000',
      bold: false,
      italic: false,
      underline: 'None',
    },
    selection = ['shape'],
    screenshot = png,
    packageRunBold = true,
    packageRunLinked = false,
    changePackageFormatOnRangeEdit = false
  const values = new Map<string, string>()
  const save = vi.fn(async () => {})
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save,
    location: () => location,
  }
  const bind = () => createPresentationDocumentBinding(settings, () => 'doc')
  const binding = bind()
  await binding.documentId()
  const shape = () => ({
    id: 'shape',
    name: 'Title',
    type: shapeType,
    ...geometry,
    text,
    font: { ...font },
  })
  const otherShape = () => ({ ...shape(), id: 'other-shape', text: otherText })
  vi.spyOn(BrowserPresentationBaselineAdapter.prototype, 'readContext').mockImplementation(
    async () => ({
      slideIds: ['slide', 'other'],
      selectedSlideIds: ['slide'],
      selectedShapeIds: selection,
      slideWidth: 960,
      slideHeight: 540,
    }),
  )
  vi.spyOn(BrowserPresentationBaselineAdapter.prototype, 'readPage').mockImplementation(
    async (id) => ({
      slideId: id,
      shapes: id === 'slide' ? [shape()] : id === 'other' ? [otherShape()] : [],
    }),
  )
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'inspectSlideMasters').mockRejectedValue(
    new Error('office_api_unsupported'),
  )
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'readPresentationPageText').mockImplementation(
    async (slideId, shapeId) => ({ slideId, shapeId, text, paragraphs: [text] }),
  )
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'readPresentationPageGeometry').mockImplementation(
    async (slideId, shapeId) => ({ slideId, shapeId, geometry: { ...geometry } }),
  )
  const editText = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'editPresentationPageText')
    .mockImplementation(async (slide, _shape, next, expected) => {
      if (slide === 'other') {
        if (otherText !== expected) throw new Error('office_concurrent_change')
        otherText = next
        return
      }
      if (text !== expected) throw new Error('office_concurrent_change')
      text = next
    })
  let textRangePackageMode = false
  const readTextRange = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'readPresentationPageTextRange')
    .mockImplementation(async (slideId, shapeId, start, length) => {
      textRangePackageMode = true
      return {
        slideId,
        shapeId,
        start,
        length,
        fullText: text,
        text: text.slice(start, start + length),
        font: { ...rangeFont },
      }
    })
  const editTextRange = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'editPresentationPageTextRange')
    .mockImplementation(async (expected, next) => {
      if (text !== expected.fullText || JSON.stringify(rangeFont) !== JSON.stringify(expected.font))
        throw new Error('office_concurrent_change')
      text = text.slice(0, expected.start) + next + text.slice(expected.start + expected.length)
      if (changePackageFormatOnRangeEdit) packageRunBold = !packageRunBold
    })
  const editGeometry = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'editPresentationPageGeometry')
    .mockImplementation(async (_slide, _shape, next, expected) => {
      if (JSON.stringify(geometry) !== JSON.stringify(expected))
        throw new Error('office_concurrent_change')
      geometry = { ...next }
    })
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'readPresentationTableCell').mockImplementation(
    async (slideId, shapeId, rowIndex, columnIndex) => ({
      slideId,
      shapeId,
      rowIndex,
      columnIndex,
      text: columnIndex === 0 ? tableText : tableText2,
      rowCount: 1,
      columnCount: 2,
    }),
  )
  const editTableCell = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'editPresentationTableCell')
    .mockImplementation(async (_slideId, _shapeId, _rowIndex, _columnIndex, next, expected) => {
      const actual = _columnIndex === 0 ? tableText : tableText2
      if (actual !== expected) throw new Error('office_concurrent_change')
      if (_columnIndex === 0) tableText = next
      else tableText2 = next
    })
  const zip = new JSZip()
  let packageComment = ''
  const tableXml = () =>
    `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="shape" name="Table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>${tableText}</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>${tableText2}</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>`
  const escapeXml = (value: string) => value.replaceAll('&', '&amp;').replaceAll('<', '&lt;')
  const textXml = () =>
    `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="shape" name="Title"/></p:nvSpPr><p:txBody><a:p><a:r><a:rPr b="${packageRunBold ? 1 : 0}">${packageRunLinked ? '<a:hlinkClick r:id="rId1"/>' : ''}</a:rPr><a:t>${escapeXml(text.slice(0, 3))}</a:t></a:r><a:r><a:rPr i="1"/><a:t>${escapeXml(text.slice(3))}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'exportPresentationPagePackage').mockImplementation(
    async (slideId) => {
      zip.file('ppt/slides/slide1.xml', textRangePackageMode ? textXml() : tableXml())
      return {
        slideId,
        slideIds: ['slide', 'other'],
        base64: await zip.generateAsync({ type: 'base64', comment: packageComment }),
      }
    },
  )
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'inspectPresentationPage').mockImplementation(
    async (slideId) => ({
      slideId,
      slideWidth: 960,
      slideHeight: 540,
      shapes: slideId === 'other' ? [otherShape()] : [shape()],
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: screenshot },
    }),
  )
  const invalidateQa = vi.fn(async (_ids?: readonly string[]) => {})
  const backups = new Map<
    string,
    { meta: Record<string, unknown>; bytes: Uint8Array; ready: boolean }
  >()
  const releasedBackups = new Map<string, Record<string, unknown>>()
  let backupOffline = false
  let backupQuota = 8
  let afterBackupRelease: (() => void) | undefined
  const request = async (body: unknown) => {
    const input = body as Record<string, unknown>
    if (backupOffline) throw new Error('offline')
    const id = input.backupId as string
    let stored = backups.get(id)
    if (input.operation === 'existing_page_backup_release') {
      const receipt =
        releasedBackups.get(id) ??
        (stored?.ready ? { ...stored.meta, status: 'released' } : undefined)
      if (!receipt) throw new Error('backup_missing')
      releasedBackups.set(id, receipt)
      backups.delete(id)
      afterBackupRelease?.()
      return new Response(JSON.stringify(receipt), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      })
    }
    if (input.operation === 'existing_page_backup_begin') {
      if (!stored) {
        if (backups.size >= backupQuota) throw new Error('quota_exceeded')
        stored = {
          meta: Object.fromEntries(
            ['backupId', 'documentId', 'hostSlideId', 'slideIds', 'sha256', 'sizeBytes'].map(
              (key) => [key, input[key]],
            ),
          ),
          bytes: new Uint8Array(),
          ready: false,
        }
        backups.set(id, stored)
      }
    } else if (!stored) {
      if (input.operation === 'existing_page_backup_status')
        return new Response(JSON.stringify({ error: 'not_found' }), { status: 404 })
      throw new Error('backup_missing')
    }
    if (input.operation === 'existing_page_backup_chunk') {
      const chunk = Uint8Array.from(atob(input.base64 as string), (character) =>
        character.charCodeAt(0),
      )
      if (input.offset !== stored!.bytes.length) throw new Error('offset_mismatch')
      stored!.bytes = Uint8Array.from([...stored!.bytes, ...chunk])
    }
    if (input.operation === 'existing_page_backup_finish') stored!.ready = true
    const result =
      input.operation === 'existing_page_backup_read'
        ? {
            backupId: id,
            offset: input.offset,
            sizeBytes: stored!.meta.sizeBytes,
            sha256: stored!.meta.sha256,
            base64: btoa(
              Array.from(
                stored!.bytes.slice(
                  input.offset as number,
                  (input.offset as number) + (input.length as number),
                ),
                (byte) => String.fromCharCode(byte),
              ).join(''),
            ),
          }
        : {
            ...stored!.meta,
            status: stored!.ready ? 'ready' : 'uploading',
            receivedBytes: stored!.bytes.length,
          }
    return new Response(JSON.stringify(result), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    })
  }
  const create = () =>
    createOfficeHostRuntime('powerpoint', {
      presentation: {
        ...bind(),
        available: () => false,
        request,
        invalidateQa,
      },
    })
  let runtime = create()
  const call = (name: string, input: Record<string, unknown> = {}) =>
    runtime.skill.executeTool({ id: 'call', name, input })
  const baseline = async () => {
    const r = await call('read_presentation_baseline', { scope: 'selected' })
    expect(r.isError, r.output).not.toBe(true)
    return JSON.parse(r.output).baselineId as string
  }
  const propose = async (kind: 'text' | 'geometry' = 'text') => {
    const baseline_id = await baseline()
    const r = await call(`edit_existing_presentation_${kind}`, {
      baseline_id,
      slide_id: 'slide',
      shape_id: 'shape',
      ...(kind === 'text' ? { text: 'after' } : { geometry: { ...geometry, left: 30 } }),
    })
    expect(r.isError, r.output).not.toBe(true)
    return r
  }
  const confirm = () => {
    const proposals = runtime.proposals as StructuredProposalController
    return proposals.confirm(proposals.pending()!.id)
  }
  const records = () =>
    bind()
      .listChangeHistory()
      .filter((e) => e.kind === 'existing')
  return {
    call,
    baseline,
    propose,
    confirm,
    records,
    values,
    save,
    editText,
    readTextRange,
    editTextRange,
    editGeometry,
    editTableCell,
    tableText: () => tableText,
    tableText2: () => tableText2,
    setTableText2: (value: string) => {
      tableText2 = value
    },
    invalidateQa,
    binding: bind,
    getRuntime: () => runtime,
    reopen: () => {
      runtime.dispose()
      runtime = create()
    },
    setText: (v: string) => {
      text = v
    },
    text: () => text,
    otherText: () => otherText,
    setLocation: (v: string) => {
      location = v
    },
    setSelection: (v: string[]) => {
      selection = v
    },
    setShapeType: (v: string) => {
      shapeType = v
    },
    setFont: (v: typeof font) => {
      font = v
    },
    setRangeFont: (v: typeof rangeFont) => {
      rangeFont = v
    },
    setPackageRunBold: (value: boolean) => {
      packageRunBold = value
    },
    setPackageRunLinked: (value: boolean) => {
      packageRunLinked = value
    },
    setChangePackageFormatOnRangeEdit: (value: boolean) => {
      changePackageFormatOnRangeEdit = value
    },
    setScreenshot: (v: string) => {
      screenshot = v
    },
    setBackupOffline: (value: boolean) => {
      backupOffline = value
    },
    setBackupQuota: (value: number) => {
      backupQuota = value
    },
    setAfterBackupRelease: (callback: (() => void) | undefined) => {
      afterBackupRelease = callback
    },
    releasedBackups: () => releasedBackups.size,
    setPackageComment: (value: string) => {
      packageComment = value
    },
    readyBackups: () => [...backups.values()].filter((backup) => backup.ready).length,
    geometry: () => geometry,
  }
}
it.each(['Chart', 'Table', 'Group', 'SmartArt', 'Placeholder'])(
  'does not offer native geometry for complex %s shapes',
  async (type) => {
    const f = await fixture()
    f.setShapeType(type)
    const baseline_id = await f.baseline()
    const single = await f.call('edit_existing_presentation_geometry', {
      baseline_id,
      slide_id: 'slide',
      shape_id: 'shape',
      geometry: { ...f.geometry(), left: 30 },
    })
    expect(single.isError).toBe(true)
    expect(single.output).toContain('presentation_existing_target_unsupported')
    const batch = await f.call('edit_existing_presentation_batch', {
      baseline_id,
      intent: 'Move chart and update caption',
      preserved: [],
      validation: [],
      risk: 'medium',
      operations: [
        {
          slide_id: 'slide',
          shape_id: 'shape',
          kind: 'geometry',
          geometry: { ...f.geometry(), left: 30 },
        },
        { slide_id: 'other', shape_id: 'other-shape', kind: 'text', text: 'other-after' },
      ],
    })
    expect(batch.isError).toBe(true)
    expect(batch.output).toContain('presentation_existing_target_unsupported')
    expect(f.editGeometry).not.toHaveBeenCalled()
    expect(f.getRuntime().proposals.pending()).toBeUndefined()
  },
)
it.each(['name', 'size', 'color', 'bold', 'italic', 'underline'] as const)(
  'rejects whole-range text edits when %s formatting is indeterminate',
  async (field) => {
    const f = await fixture()
    f.setFont({
      name: 'Arial',
      size: 20,
      color: '#000000',
      bold: false,
      italic: false,
      underline: 'None',
      [field]: null,
    })
    const baseline_id = await f.baseline()
    const single = await f.call('edit_existing_presentation_text', {
      baseline_id,
      slide_id: 'slide',
      shape_id: 'shape',
      text: 'after',
    })
    expect(single.output).toContain('presentation_existing_target_unsupported')
    const batch = await f.call('edit_existing_presentation_batch', {
      baseline_id,
      intent: 'Update two labels',
      preserved: [],
      validation: [],
      risk: 'medium',
      operations: [
        { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
        { slide_id: 'other', shape_id: 'other-shape', kind: 'text', text: 'other-after' },
      ],
    })
    expect(batch.isError).toBe(true)
    expect(f.editText).not.toHaveBeenCalled()
    expect(f.getRuntime().proposals.pending()).toBeUndefined()
  },
)
it('rejects text edits when aggregate font data is unavailable', async () => {
  const f = await fixture()
  vi.spyOn(BrowserPresentationBaselineAdapter.prototype, 'readPage').mockImplementation(
    async (slideId) => ({
      slideId,
      shapes: [{ id: 'shape', name: 'Title', type: 'TextBox', ...f.geometry(), text: 'before' }],
    }),
  )
  const baseline_id = await f.baseline()
  const result = await f.call('edit_existing_presentation_text', {
    baseline_id,
    slide_id: 'slide',
    shape_id: 'shape',
    text: 'after',
  })
  expect(result.output).toContain('presentation_existing_target_unsupported')
  expect(f.editText).not.toHaveBeenCalled()
})
it('confirms and undoes a bounded text range in a mixed-font shape', async () => {
  const f = await fixture()
  f.setFont({
    name: null,
    size: 20,
    color: '#000000',
    bold: false,
    italic: false,
    underline: 'None',
  })
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_text_range', {
    baseline_id,
    slide_id: 'slide',
    shape_id: 'shape',
    range_start: 0,
    range_length: 3,
    text: 'new',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  expect(f.editTextRange).not.toHaveBeenCalled()
  await f.confirm()
  expect(f.text()).toBe('newore')
  expect(f.editText).not.toHaveBeenCalled()
  expect(f.readyBackups()).toBe(1)
  const record = f.records()[0]!.record
  expect(record).toMatchObject({
    kind: 'text_range',
    state: 'applied',
    start: 0,
    length: 3,
    before: 'before',
    after: 'newore',
  })
  if (record.kind !== 'text_range') throw new Error('unexpected_change_kind')
  expect(record.runStructureDigest).toMatch(/^[a-f0-9]{64}$/)
  f.reopen()
  const undo = await f.call('undo_existing_presentation_change', { change_id: record.changeId })
  expect(undo.isError, undo.output).not.toBe(true)
  await f.confirm()
  expect(f.text()).toBe('before')
  expect(f.records()[0]!.record.state).toBe('undone')
  expect(f.editTextRange).toHaveBeenCalledTimes(2)
})
it('rejects an invalid range and blocks a stale range font before writing', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  for (const input of [
    { range_start: 0, range_length: 3, text: 'long' },
    { range_start: 0, range_length: 1, text: '\n' },
    { range_start: 0, range_length: 1, text: '\ud800' },
  ]) {
    const result = await f.call('edit_existing_presentation_text_range', {
      baseline_id,
      slide_id: 'slide',
      shape_id: 'shape',
      ...input,
    })
    expect(result.output).toBe('invalid_tool_input')
  }
  const proposed = await f.call('edit_existing_presentation_text_range', {
    baseline_id,
    slide_id: 'slide',
    shape_id: 'shape',
    range_start: 0,
    range_length: 3,
    text: 'new',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  f.setRangeFont({
    name: 'Arial',
    size: 22,
    color: '#000000',
    bold: false,
    italic: false,
    underline: 'None',
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.text()).toBe('before')
  expect(f.editTextRange).not.toHaveBeenCalled()
})
it('rejects a text range crossing package runs or carrying a hyperlink', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  const crossing = await f.call('edit_existing_presentation_text_range', {
    baseline_id,
    slide_id: 'slide',
    shape_id: 'shape',
    range_start: 2,
    range_length: 3,
    text: 'new',
  })
  expect(crossing.output).toBe('presentation_existing_target_unsupported')
  f.setPackageRunLinked(true)
  const linked = await f.call('edit_existing_presentation_text_range', {
    baseline_id,
    slide_id: 'slide',
    shape_id: 'shape',
    range_start: 0,
    range_length: 3,
    text: 'new',
  })
  expect(linked.output).toBe('presentation_existing_target_unsupported')
  expect(f.editTextRange).not.toHaveBeenCalled()
})
it('blocks a pending text range when package run formatting changes', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_text_range', {
    baseline_id,
    slide_id: 'slide',
    shape_id: 'shape',
    range_start: 0,
    range_length: 3,
    text: 'new',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  f.setPackageRunBold(false)
  await expect(f.confirm()).rejects.toThrow()
  expect(f.text()).toBe('before')
  expect(f.editTextRange).not.toHaveBeenCalled()
  expect(f.readyBackups()).toBe(0)
})
it('keeps a text range pending if PowerPoint changes run formatting during the write', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_text_range', {
    baseline_id,
    slide_id: 'slide',
    shape_id: 'shape',
    range_start: 0,
    range_length: 3,
    text: 'new',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  f.setChangePackageFormatOnRangeEdit(true)
  await expect(f.confirm()).rejects.toThrow()
  expect(f.text()).toBe('newore')
  expect(f.readyBackups()).toBe(1)
  expect(f.records()[0]!.record.state).toBe('pending')
  expect(f.editTextRange).toHaveBeenCalledTimes(1)
  f.reopen()
  const resume = await f.call('resume_existing_presentation_change', {
    change_id: f.records()[0]!.record.changeId,
  })
  expect(resume.isError).toBe(true)
  expect(f.editTextRange).toHaveBeenCalledTimes(1)
})
it.each(['text', 'geometry'] as const)(
  'persists a confirmed existing %s change and undoes it after reopening while offline',
  async (kind) => {
    const f = await fixture()
    await f.propose(kind)
    expect(f.records()).toHaveLength(0)
    expect(f.editText).not.toHaveBeenCalled()
    expect(f.editGeometry).not.toHaveBeenCalled()
    await f.confirm()
    const record = f.records()[0]!.record
    expect(record).toMatchObject({ state: 'applied', kind, hostSlideId: 'slide', shapeId: 'shape' })
    expect(f.readyBackups()).toBe(1)
    expect(record).toHaveProperty('backup.sha256')
    expect(record).not.toHaveProperty('projectId')
    expect(f.invalidateQa).toHaveBeenCalledWith(['slide'])
    f.reopen()
    const listed = await f.call('list_existing_presentation_changes')
    expect(JSON.parse(listed.output).changes).toHaveLength(1)
    expect(
      (await f.call('undo_existing_presentation_change', { change_id: record.changeId })).isError,
    ).not.toBe(true)
    await f.confirm()
    expect(f.records()[0]!.record.state).toBe('undone')
    expect(kind === 'text' ? f.text() : f.geometry().left).toBe(kind === 'text' ? 'before' : 1)
    f.setTableText2('unrelated content changed')
    const unsafeRelease = await f.call('release_existing_presentation_change', {
      change_id: record.changeId,
    })
    expect(unsafeRelease.isError).toBe(true)
    expect(f.readyBackups()).toBe(1)
    f.setTableText2('before-2')
    const release = await f.call('release_existing_presentation_change', {
      change_id: record.changeId,
    })
    expect(release.isError, release.output).not.toBe(true)
    await f.confirm()
    expect(f.readyBackups()).toBe(0)
    expect(f.binding().readExistingChange(record.changeId)?.backupReleasedAt).toBeTruthy()
  },
)
it('keeps a single existing-page edit unwritten when its PC package backup is unavailable and resumes it later', async () => {
  const f = await fixture()
  await f.propose()
  f.setBackupOffline(true)
  await expect(f.confirm()).rejects.toThrow()
  expect(f.text()).toBe('before')
  expect(f.editText).not.toHaveBeenCalled()
  const record = f.records()[0]?.record
  expect(record).toMatchObject({ state: 'pending', backup: { hostSlideId: 'slide' } })
  f.setBackupOffline(false)
  f.reopen()
  const resumed = await f.call('resume_existing_presentation_change', {
    change_id: record!.changeId,
  })
  expect(resumed.isError, resumed.output).not.toBe(true)
  await f.confirm()
  expect(f.text()).toBe('after')
  expect(f.readyBackups()).toBe(1)
})
it('does not write a single existing-page edit when the PC backup quota is full', async () => {
  const f = await fixture()
  await f.propose('geometry')
  f.setBackupQuota(0)
  await expect(f.confirm()).rejects.toThrow()
  expect(f.editGeometry).not.toHaveBeenCalled()
  expect(f.geometry().left).toBe(1)
  expect(f.records()[0]?.record.state).toBe('pending')
  const previousBackup = f.records()[0]!.record.backup!
  f.setBackupQuota(8)
  f.reopen()
  f.setPackageComment('equivalent ZIP with changed metadata')
  const resumed = await f.call('resume_existing_presentation_change', {
    change_id: f.records()[0]!.record.changeId,
  })
  expect(resumed.isError, resumed.output).not.toBe(true)
  await f.confirm()
  expect(f.geometry().left).toBe(30)
  expect(f.readyBackups()).toBe(1)
  const finalBackup = f.records()[0]!.record.backup!
  expect(finalBackup.packageDigest).toBe(previousBackup.packageDigest)
  expect(finalBackup.sha256).not.toBe(previousBackup.sha256)
})
it('rejects a single edit if the original page package changes after proposal', async () => {
  const f = await fixture()
  await f.propose()
  f.setTableText2('concurrent unrelated change')
  await expect(f.confirm()).rejects.toThrow()
  expect(f.editText).not.toHaveBeenCalled()
  expect(f.readyBackups()).toBe(0)
  expect(f.records()[0]?.record.state).toBe('pending')
})
it('reuses a ready original page backup after an interrupted single host write', async () => {
  const f = await fixture()
  await f.propose()
  f.editText.mockRejectedValueOnce(new Error('host_unavailable'))
  await expect(f.confirm()).rejects.toThrow()
  expect(f.text()).toBe('before')
  expect(f.readyBackups()).toBe(1)
  f.reopen()
  const resumed = await f.call('resume_existing_presentation_change', {
    change_id: f.records()[0]!.record.changeId,
  })
  expect(resumed.isError, resumed.output).not.toBe(true)
  await f.confirm()
  expect(f.text()).toBe('after')
  expect(f.readyBackups()).toBe(1)
})
it('applies and reverses an ordered text plus geometry batch through one durable savepoint', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update title and placement',
    preserved: ['Other objects'],
    validation: ['Readback', 'Page screenshot'],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      {
        slide_id: 'slide',
        shape_id: 'shape',
        kind: 'geometry',
        geometry: { ...f.geometry(), left: 30 },
      },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  expect(f.binding().listChangeHistory()).toHaveLength(0)
  await f.confirm()
  const record = f.binding().listChangeHistory()[0]
  expect(record).toMatchObject({ kind: 'existing_batch', record: { state: 'applied', cursor: 2 } })
  expect(f.text()).toBe('after')
  expect(f.geometry().left).toBe(30)
  expect(f.invalidateQa).toHaveBeenCalledWith(['slide'])
  const listed = await f.call('list_existing_presentation_changes')
  expect(JSON.parse(listed.output).changes).toMatchObject([{ kind: 'batch', operationCount: 2 }])
  const captured = await f.call('capture_existing_presentation_batch_page', {
    change_id: changeId,
    slide_id: 'slide',
  })
  expect(captured.isError, captured.output).not.toBe(true)
  expect(JSON.parse(captured.output)).toMatchObject({ hostSlideId: 'slide', qaPassed: false })
  const digest = JSON.parse(captured.output).screenshotDigest as string
  const reviewed = await f.call('record_existing_presentation_batch_page_review', {
    change_id: changeId,
    slide_id: 'slide',
    screenshot_digest: digest,
    status: 'pass',
    notes: 'Title and position checked',
  })
  expect(reviewed.isError, reviewed.output).not.toBe(true)
  expect(f.binding().readExistingBatch(changeId)?.reviews).toMatchObject([
    { hostSlideId: 'slide', status: 'pass' },
  ])
  f.reopen()
  const workbench = f.getRuntime().changes!
  await workbench.refresh()
  const row = workbench
    .snapshot()
    .entries.find((entry) => entry.id === `existing_batch:${changeId}`)
  expect(row).toMatchObject({
    source: 'existing_batch',
    state: 'applied',
    actions: ['inspect', 'undo'],
    reviews: [{ hostSlideId: 'slide', status: 'pass' }],
  })
  await workbench.run(row!.id, 'undo')
  expect(f.getRuntime().proposals.pending()).toBeDefined()
  await f.confirm()
  expect(f.text()).toBe('before')
  expect(f.geometry().left).toBe(1)
  expect(f.binding().listChangeHistory()[0].record).toMatchObject({ state: 'undone', cursor: 0 })
  expect(f.binding().readExistingBatch(changeId)?.reviews).toBeUndefined()
  expect(f.readyBackups()).toBe(1)
  await workbench.refresh()
  const undoneRow = workbench
    .snapshot()
    .entries.find((entry) => entry.id === `existing_batch:${changeId}`)
  expect(undoneRow?.actions).toEqual(['inspect', 'release'])
  f.setTableText2('unrelated page content changed')
  const blockedRelease = await f.call('release_existing_presentation_batch', {
    change_id: changeId,
  })
  expect(blockedRelease.isError).toBe(true)
  expect(f.readyBackups()).toBe(1)
  f.setTableText2('before-2')
  await workbench.run(undoneRow!.id, 'release')
  await f.confirm()
  expect(f.readyBackups()).toBe(0)
  expect(f.binding().readExistingBatch(changeId)?.backupReleasedAt).toBeTruthy()
  await workbench.refresh()
  expect(workbench.snapshot().entries.find((entry) => entry.id === undoneRow!.id)?.actions).toEqual(
    ['inspect'],
  )
})
it('stops releasing batch backups after the first receipt and safely retries release', async () => {
  const f = await fixture()
  const baseline = await f.call('read_presentation_baseline', { scope: 'deck' })
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id: JSON.parse(baseline.output).baselineId,
    intent: 'Update two pages',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      { slide_id: 'other', shape_id: 'other-shape', kind: 'text', text: 'other-after' },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  await f.confirm()
  expect(f.readyBackups()).toBe(2)
  f.reopen()
  const undo = await f.call('undo_existing_presentation_batch', { change_id: changeId })
  expect(undo.isError, undo.output).not.toBe(true)
  await f.confirm()
  f.reopen()
  const release = await f.call('release_existing_presentation_batch', { change_id: changeId })
  expect(release.isError, release.output).not.toBe(true)
  f.setAfterBackupRelease(() => f.getRuntime().proposals.logout())
  await expect(f.confirm()).rejects.toThrow()
  expect(f.readyBackups()).toBe(1)
  expect(f.releasedBackups()).toBe(1)
  expect(f.binding().readExistingBatch(changeId)?.backupReleasedAt).toBeUndefined()
  f.setAfterBackupRelease(undefined)
  f.reopen()
  const retry = await f.call('release_existing_presentation_batch', { change_id: changeId })
  expect(retry.isError, retry.output).not.toBe(true)
  await f.confirm()
  expect(f.readyBackups()).toBe(0)
  expect(f.releasedBackups()).toBe(2)
  expect(f.binding().readExistingBatch(changeId)?.backupReleasedAt).toBeTruthy()
})

it('completes all affected page package backups before the first batch host write', async () => {
  const f = await fixture()
  const baseline = await f.call('read_presentation_baseline', { scope: 'deck' })
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id: JSON.parse(baseline.output).baselineId,
    intent: 'Update two titles',
    preserved: ['Other objects'],
    validation: ['Read back both pages'],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      { slide_id: 'other', shape_id: 'other-shape', kind: 'text', text: 'other-after' },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  f.editText.mockImplementationOnce(async (_slide, _shape, next) => {
    expect(f.readyBackups()).toBe(2)
    f.setText(next)
  })
  await f.confirm()
  const record = f.binding().readExistingBatch(JSON.parse(proposed.output).changeId)
  expect(record?.backups).toHaveLength(2)
  expect(record?.backups?.map((backup) => backup.hostSlideId)).toEqual(['slide', 'other'])
})
it('keeps PowerPoint unchanged when PC backup is unavailable and resumes after reconnect', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update title and position',
    preserved: ['Other objects'],
    validation: ['Readback'],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      {
        slide_id: 'slide',
        shape_id: 'shape',
        kind: 'geometry',
        geometry: { ...f.geometry(), left: 30 },
      },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  f.setBackupOffline(true)
  await expect(f.confirm()).rejects.toThrow('offline')
  expect(f.text()).toBe('before')
  expect(f.geometry().left).toBe(1)
  expect(f.binding().readExistingBatch(changeId)).toMatchObject({ state: 'applying', cursor: 0 })
  f.setBackupOffline(false)
  f.reopen()
  const resumed = await f.call('resume_existing_presentation_batch', { change_id: changeId })
  expect(resumed.isError, resumed.output).not.toBe(true)
  await f.confirm()
  expect(f.binding().readExistingBatch(changeId)).toMatchObject({ state: 'applied', cursor: 2 })
  expect(f.readyBackups()).toBe(1)
})
it('does not begin a multi-page batch when PC backup quota fills partway', async () => {
  const f = await fixture()
  const baseline = await f.call('read_presentation_baseline', { scope: 'deck' })
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id: JSON.parse(baseline.output).baselineId,
    intent: 'Update two titles',
    preserved: ['Other objects'],
    validation: ['Readback'],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      { slide_id: 'other', shape_id: 'other-shape', kind: 'text', text: 'other-after' },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  f.setBackupQuota(1)
  await expect(f.confirm()).rejects.toThrow('quota_exceeded')
  expect(f.readyBackups()).toBe(1)
  expect(f.text()).toBe('before')
  expect(f.otherText()).toBe('other-before')
  const changeId = JSON.parse(proposed.output).changeId as string
  const missingBackup = f.binding().readExistingBatch(changeId)!.backups![1]!
  f.setBackupQuota(2)
  f.reopen()
  f.setPackageComment('regenerated ZIP metadata')
  const resumed = await f.call('resume_existing_presentation_batch', {
    change_id: changeId,
  })
  expect(resumed.isError, resumed.output).not.toBe(true)
  await f.confirm()
  expect(f.readyBackups()).toBe(2)
  expect(f.otherText()).toBe('other-after')
  expect(f.binding().readExistingBatch(changeId)!.backups![1]!.sha256).not.toBe(
    missingBackup.sha256,
  )
})
it('stops a batch when a non-target shape changes during a confirmed write', async () => {
  const f = await fixture()
  const original = vi
    .mocked(BrowserPresentationBaselineAdapter.prototype.readPage)
    .getMockImplementation()!
  let untouched = 'keep'
  vi.spyOn(BrowserPresentationBaselineAdapter.prototype, 'readPage').mockImplementation(
    async (slideId, signal) => {
      const page = await original(slideId, signal)
      if (slideId === 'slide')
        page.shapes.push({ ...page.shapes[0]!, id: 'untouched', name: 'Keep', text: untouched })
      return page
    },
  )
  const baseline = await f.call('read_presentation_baseline', { scope: 'deck' })
  const baseline_id = JSON.parse(baseline.output).baselineId as string
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update both titles',
    preserved: ['Other objects'],
    validation: ['Read back both pages'],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      { slide_id: 'other', shape_id: 'other-shape', kind: 'text', text: 'other-after' },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  f.editText.mockImplementationOnce(async () => {
    f.setText('after')
    untouched = 'changed'
  })
  await expect(f.confirm()).rejects.toThrow('presentation_existing_preserved_changed')
  expect(f.otherText()).toBe('other-before')
  const saved = f.binding().readExistingBatch(JSON.parse(proposed.output).changeId)
  expect(saved).toMatchObject({ state: 'applying', cursor: 0 })
  expect(saved?.preservedPageDigests).toHaveProperty('slide')
  expect(saved?.preservedPageDigests).toHaveProperty('other')
  f.reopen()
  expect(
    (
      await f.call('inspect_existing_presentation_batch', {
        change_id: JSON.parse(proposed.output).changeId,
      })
    ).output,
  ).toBe('presentation_existing_preserved_changed')
})
it('stops a batch when a text edit also changes an unplanned target font', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update title and position',
    preserved: ['Title font'],
    validation: ['Read back title and position'],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      {
        slide_id: 'slide',
        shape_id: 'shape',
        kind: 'geometry',
        geometry: { ...f.geometry(), left: 30 },
      },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  f.editText.mockImplementationOnce(async () => {
    f.setText('after')
    f.setFont({
      name: 'Other',
      size: 20,
      color: '#000000',
      bold: false,
      italic: false,
      underline: 'None',
    })
  })
  await expect(f.confirm()).rejects.toThrow('presentation_existing_preserved_changed')
  expect(f.geometry().left).toBe(1)
  expect(
    f.binding().readExistingBatch(JSON.parse(proposed.output).changeId)?.preservedTargetDigests,
  ).toHaveProperty(JSON.stringify(['slide', 'shape']))
})
it('edits two native objects on two pages and invalidates both QA page scopes', async () => {
  const f = await fixture()
  const baseline = await f.call('read_presentation_baseline', { scope: 'deck' })
  expect(baseline.isError, baseline.output).not.toBe(true)
  const baseline_id = JSON.parse(baseline.output).baselineId as string
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update both titles',
    preserved: ['Other objects'],
    validation: ['Review both pages'],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      { slide_id: 'other', shape_id: 'other-shape', kind: 'text', text: 'other-after' },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  await f.confirm()
  expect(f.text()).toBe('after')
  expect(f.otherText()).toBe('other-after')
  expect(f.invalidateQa).toHaveBeenCalledWith(['slide', 'other'])
  for (const slide_id of ['slide', 'other']) {
    const shot = await f.call('capture_existing_presentation_batch_page', {
      change_id: changeId,
      slide_id,
    })
    expect(shot.isError, shot.output).not.toBe(true)
    const reviewed = await f.call('record_existing_presentation_batch_page_review', {
      change_id: changeId,
      slide_id,
      screenshot_digest: JSON.parse(shot.output).screenshotDigest,
      status: 'pass',
      notes: 'Checked',
    })
    expect(reviewed.isError, reviewed.output).not.toBe(true)
  }
  expect(f.binding().readExistingBatch(changeId)?.reviews).toHaveLength(2)
  f.reopen()
  const undo = await f.call('undo_existing_presentation_batch', { change_id: changeId })
  expect(undo.isError, undo.output).not.toBe(true)
  await f.confirm()
  expect(f.text()).toBe('before')
  expect(f.otherText()).toBe('other-before')
})
it('rejects a changed or cross-session batch screenshot review', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update title and position',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      {
        slide_id: 'slide',
        shape_id: 'shape',
        kind: 'geometry',
        geometry: { ...f.geometry(), left: 30 },
      },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  await f.confirm()
  const capture = async () =>
    f.call('capture_existing_presentation_batch_page', { change_id: changeId, slide_id: 'slide' })
  const first = await capture()
  expect(first.isError, first.output).not.toBe(true)
  const review = {
    change_id: changeId,
    slide_id: 'slide',
    screenshot_digest: JSON.parse(first.output).screenshotDigest,
    status: 'pass',
    notes: 'Checked',
  }
  f.setScreenshot(otherPng)
  const changed = await f.call('record_existing_presentation_batch_page_review', review)
  expect(changed).toMatchObject({ isError: true, output: 'presentation_existing_batch_qa_stale' })
  f.setScreenshot(png)
  await capture()
  f.reopen()
  const reopened = await f.call('record_existing_presentation_batch_page_review', review)
  expect(reopened).toMatchObject({ isError: true, output: 'presentation_existing_batch_qa_stale' })
  expect(f.binding().readExistingBatch(changeId)?.reviews).toBeUndefined()
})
it('invalidates a captured batch review when another confirmed host edit begins', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update title and position',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      {
        slide_id: 'slide',
        shape_id: 'shape',
        kind: 'geometry',
        geometry: { ...f.geometry(), left: 30 },
      },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  await f.confirm()
  const captured = await f.call('capture_existing_presentation_batch_page', {
    change_id: changeId,
    slide_id: 'slide',
  })
  expect(captured.isError, captured.output).not.toBe(true)
  const nextBaseline = await f.call('read_presentation_baseline', { scope: 'deck' })
  expect(nextBaseline.isError, nextBaseline.output).not.toBe(true)
  const next = await f.call('edit_existing_presentation_text', {
    baseline_id: JSON.parse(nextBaseline.output).baselineId,
    slide_id: 'other',
    shape_id: 'other-shape',
    text: 'other-after',
  })
  expect(next.isError, next.output).not.toBe(true)
  await f.confirm()
  const stale = await f.call('record_existing_presentation_batch_page_review', {
    change_id: changeId,
    slide_id: 'slide',
    screenshot_digest: JSON.parse(captured.output).screenshotDigest,
    status: 'pass',
    notes: 'Old capture',
  })
  expect(stale).toMatchObject({ isError: true, output: 'presentation_existing_batch_qa_stale' })
  expect(f.binding().readExistingBatch(changeId)?.reviews).toBeUndefined()
})
it('resumes a batch after the first step was saved and the second host write failed', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update title and placement',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      {
        slide_id: 'slide',
        shape_id: 'shape',
        kind: 'geometry',
        geometry: { ...f.geometry(), left: 30 },
      },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  f.editGeometry.mockRejectedValueOnce(new Error('office_write_failed'))
  await expect(f.confirm()).rejects.toThrow()
  const entry = f.binding().listChangeHistory()[0]
  expect(entry).toMatchObject({ kind: 'existing_batch', record: { state: 'applying', cursor: 1 } })
  expect(f.text()).toBe('after')
  expect(f.geometry().left).toBe(1)
  f.reopen()
  const inspected = await f.call('inspect_existing_presentation_batch', { change_id: changeId })
  expect(JSON.parse(inspected.output).values).toEqual(['after', 'before'])
  const resumed = await f.call('resume_existing_presentation_batch', { change_id: changeId })
  expect(resumed.isError, resumed.output).not.toBe(true)
  await f.confirm()
  expect(f.editText).toHaveBeenCalledTimes(1)
  expect(f.binding().listChangeHistory()[0].record).toMatchObject({ state: 'applied', cursor: 2 })
})
it('finalizes an interrupted first-step receipt without replaying the host write', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update title and placement',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      {
        slide_id: 'slide',
        shape_id: 'shape',
        kind: 'geometry',
        geometry: { ...f.geometry(), left: 30 },
      },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  let failed = false
  f.save.mockImplementation(async () => {
    if (!failed && f.editText.mock.calls.length === 1) {
      failed = true
      throw new Error('receipt_failed')
    }
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.binding().readExistingBatch(changeId)).toMatchObject({ state: 'applying', cursor: 0 })
  expect(f.text()).toBe('after')
  f.reopen()
  const inspected = await f.call('inspect_existing_presentation_batch', { change_id: changeId })
  expect(JSON.parse(inspected.output).values).toEqual(['after', 'before'])
  const resumed = await f.call('resume_existing_presentation_batch', { change_id: changeId })
  expect(resumed.isError, resumed.output).not.toBe(true)
  await f.confirm()
  expect(f.editText).toHaveBeenCalledTimes(1)
  expect(f.binding().readExistingBatch(changeId)).toMatchObject({ state: 'applied', cursor: 2 })
})
it.each(['text', 'selection', 'document'] as const)(
  'rejects %s drift before confirmation without a host write',
  async (drift) => {
    const f = await fixture()
    await f.propose()
    if (drift === 'text') f.setText('manual')
    else if (drift === 'selection') f.setSelection([])
    else f.setLocation('file://save-as.pptx')
    await expect(f.confirm()).rejects.toThrow()
    expect(f.editText).not.toHaveBeenCalled()
    expect(f.records()).toHaveLength(0)
  },
)
it('saves pending before writing and only completes the receipt after a failed completion save', async () => {
  const f = await fixture()
  await f.propose()
  let failed = false
  f.save.mockImplementation(async () => {
    const raw = f.values.get('wiswork.presentation.existing-change.v1')
    if (!failed && raw && JSON.parse(raw).state === 'applied') {
      failed = true
      throw new Error('save_failed')
    }
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.text()).toBe('after')
  expect(f.records()[0]!.record.state).toBe('pending')
  const id = f.records()[0]!.record.changeId
  f.reopen()
  const inspection = JSON.parse(
    (await f.call('inspect_existing_presentation_change', { change_id: id })).output,
  )
  expect(inspection.status).toBe('already_applied')
  await f.call('resume_existing_presentation_change', { change_id: id })
  await f.confirm()
  expect(f.records()[0]!.record.state).toBe('applied')
  expect(f.editText).toHaveBeenCalledTimes(1)
})
it('does not write the host when the pending save fails and refuses undo after a manual target edit', async () => {
  const f = await fixture()
  await f.propose()
  f.save.mockRejectedValueOnce(new Error('save_failed'))
  await expect(f.confirm()).rejects.toThrow()
  expect(f.editText).not.toHaveBeenCalled()
  await f.propose()
  await f.confirm()
  f.setText('manual')
  expect(
    (
      await f.call('undo_existing_presentation_change', {
        change_id: f.records()[0]!.record.changeId,
      })
    ).isError,
  ).toBe(true)
  expect(f.text()).toBe('manual')
})
it('binds visual review to a freshly captured unchanged page and preserves it as historical evidence', async () => {
  const f = await fixture()
  await f.propose()
  await f.confirm()
  const change_id = f.records()[0]!.record.changeId
  const capture = await f.call('capture_existing_presentation_change', { change_id })
  expect(capture.isError, capture.output).not.toBe(true)
  expect(capture.modelContent?.[0]?.type).toBe('image')
  const screenshot_digest = JSON.parse(capture.output).screenshotDigest
  const r = await f.call('record_existing_presentation_change_review', {
    change_id,
    screenshot_digest,
    status: 'pass',
    notes: '文字可读且位置合适',
  })
  expect(r.isError, r.output).not.toBe(true)
  expect(f.records()[0]!.record.review?.status).toBe('pass')
  f.setScreenshot(otherPng)
  expect(
    (
      await f.call('record_existing_presentation_change_review', {
        change_id,
        screenshot_digest,
        status: 'pass',
        notes: 'old',
      })
    ).isError,
  ).toBe(true)
  f.reopen()
  expect(
    (
      await f.call('record_existing_presentation_change_review', {
        change_id,
        screenshot_digest,
        status: 'pass',
        notes: 'old',
      })
    ).isError,
  ).toBe(true)
})
it('keeps proposals stale after clearing the session and rejects out-of-scope shapes', async () => {
  const f = await fixture(),
    baseline_id = await f.baseline()
  expect(
    (
      await f.call('edit_existing_presentation_text', {
        baseline_id,
        slide_id: 'other',
        shape_id: 'shape',
        text: 'no',
      })
    ).isError,
  ).toBe(true)
  await f.propose()
  const proposals = f.getRuntime().proposals as StructuredProposalController
  const id = proposals.pending()!.id
  f.getRuntime().clearSession()
  await expect(proposals.confirm(id)).rejects.toThrow('proposal_missing')
  expect(f.editText).not.toHaveBeenCalled()
})
it('recovers an interrupted undo after reopening without writing the host twice', async () => {
  const f = await fixture()
  await f.propose()
  await f.confirm()
  const change_id = f.records()[0]!.record.changeId
  await f.call('undo_existing_presentation_change', { change_id })
  let failed = false
  f.save.mockImplementation(async () => {
    const raw = f.values.get('wiswork.presentation.existing-change.v1')
    if (!failed && raw && JSON.parse(raw).state === 'undone') {
      failed = true
      throw new Error('save_failed')
    }
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.text()).toBe('before')
  expect(f.records()[0]!.record.state).toBe('undo_pending')
  f.reopen()
  await f.call('resume_existing_presentation_change', { change_id })
  await f.confirm()
  expect(f.records()[0]!.record.state).toBe('undone')
  expect(f.editText).toHaveBeenCalledTimes(2)
})
it('revalidates after QA invalidation waits and after pending settings save waits', async () => {
  const f = await fixture()
  await f.propose()
  f.invalidateQa.mockImplementationOnce(async () => {
    f.setText('manual during hook')
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.editText).not.toHaveBeenCalled()
  f.setText('before')
  await f.propose()
  f.save.mockImplementationOnce(async () => {
    f.setSelection([])
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.editText).not.toHaveBeenCalled()
  expect(f.records()[0]!.record.state).toBe('pending')
})
it('blocks a new existing edit while a generated text transaction is unresolved', async () => {
  const f = await fixture(),
    b = f.binding(),
    documentId = await b.documentId()
  await b.writeTextChange(
    {
      version: 1,
      changeId: 'generated',
      documentId,
      projectId: 'project',
      requestId: 'request',
      artifactDigest: 'a'.repeat(64),
      pageId: 'page',
      hostSlideId: 'slide',
      shapeId: 'shape',
      before: 'before',
      after: 'generated',
      state: 'pending',
    },
    undefined,
  )
  await f.propose()
  await expect(f.confirm()).rejects.toThrow('presentation_change_history_pending')
  expect(f.editText).not.toHaveBeenCalled()
})
it('reopens two existing changes and undoes through the offline workbench in dependency order', async () => {
  const f = await fixture()
  await f.propose()
  await f.confirm()
  await f.propose('geometry')
  await f.confirm()
  f.reopen()
  const workbench = f.getRuntime().changes!
  await workbench.refresh()
  expect(workbench.snapshot().entries).toHaveLength(2)
  for (const entry of workbench.snapshot().entries) {
    await workbench.run(entry.id, 'undo')
    expect(f.getRuntime().proposals.pending(), JSON.stringify(workbench.snapshot())).toBeDefined()
    await f.confirm()
    await workbench.refresh()
  }
  expect(f.records().map((e) => e.record.state)).toEqual(['undone', 'undone'])
  expect(f.text()).toBe('before')
  expect(f.geometry().left).toBe(1)
})
it('rejects a changed screenshot before storing a review and invalidates a capture when another change is confirmed', async () => {
  const f = await fixture()
  await f.propose()
  await f.confirm()
  const change_id = f.records()[0]!.record.changeId
  const r = await f.call('capture_existing_presentation_change', { change_id }),
    screenshot_digest = JSON.parse(r.output).screenshotDigest
  f.setScreenshot(otherPng)
  expect(
    (
      await f.call('record_existing_presentation_change_review', {
        change_id,
        screenshot_digest,
        status: 'pass',
        notes: 'mismatch',
      })
    ).output,
  ).toBe('presentation_existing_qa_stale')
  expect(f.records()[0]!.record.review).toBeUndefined()
  f.setScreenshot(png)
  await f.propose('geometry')
  await f.confirm()
  expect(
    (
      await f.call('record_existing_presentation_change_review', {
        change_id,
        screenshot_digest,
        status: 'pass',
        notes: 'stale',
      })
    ).output,
  ).toBe('presentation_existing_qa_stale')
})
it('does not reuse a baseline after a successful edit and validates malformed tool input before proposing', async () => {
  const f = await fixture(),
    baseline_id = await f.baseline()
  for (const input of [
    { baseline_id, slide_id: 'slide', shape_id: 'shape', text: 'after', extra: true },
    { baseline_id, slide_id: 'slide', shape_id: 'shape', text: 'a'.repeat(12001) },
  ])
    expect((await f.call('edit_existing_presentation_text', input)).output).toBe(
      'invalid_tool_input',
    )
  await f.call('edit_existing_presentation_text', {
    baseline_id,
    slide_id: 'slide',
    shape_id: 'shape',
    text: 'after',
  })
  await f.confirm()
  expect(
    (
      await f.call('edit_existing_presentation_text', {
        baseline_id,
        slide_id: 'slide',
        shape_id: 'shape',
        text: 'again',
      })
    ).output,
  ).toBe('presentation_baseline_changed')
  expect(f.editText).toHaveBeenCalledTimes(1)
})
