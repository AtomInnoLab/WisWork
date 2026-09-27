import JSZip from 'jszip'
import { expect, it, vi } from 'vitest'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationBaselineSkill } from '../src/skills/powerpoint/presentation-baseline'
import { inspectPowerPointPageNotes } from '../src/skills/powerpoint/presentation-notes-package'
const page = (slideId = 's1', text = 'original') => ({
  slideId,
  masterId: 'm',
  layoutId: 'l',
  shapes: [
    {
      id: 'shape',
      name: 'Title',
      type: 'TextBox',
      left: 10,
      top: 10,
      width: 100,
      height: 40,
      text,
      font: { name: 'Arial', size: 20, color: '#000000' },
    },
  ],
})
function fixture() {
  let documentId = 'doc',
    context = {
      slideIds: ['s1', 's2'],
      selectedSlideIds: ['s2'],
      selectedShapeIds: ['shape'],
      slideWidth: 960,
      slideHeight: 540,
    }
  const pages = new Map([
    ['s1', page()],
    ['s2', page('s2')],
  ])
  const adapter = {
    readContext: vi.fn(async () => structuredClone(context)),
    readPage: vi.fn(async (id: string) => structuredClone(pages.get(id)!)),
  }
  const inspectPage = vi.fn(async (slideId: string) => ({
    slideId,
    slideWidth: 960,
    slideHeight: 540,
    shapes: pages.get(slideId)!.shapes.map(({ text: _text, font: _font, ...shape }) => shape),
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: {
      mime: 'image/png' as const,
      base64:
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII=',
    },
  }))
  let packageBase64 = ''
  const exportPagePackage = vi.fn(async (slideId: string) => ({
    slideId,
    slideIds: [...context.slideIds],
    base64: packageBase64,
  }))
  const skill = createPresentationBaselineSkill({
    adapter,
    documentId: async () => documentId,
    inspectPage,
    exportPagePackage,
  })
  const call = (name: string, input: Record<string, unknown> = {}, signal?: AbortSignal) =>
    skill.executeTool({ id: 'call', name, input }, signal)
  const read = (scope = 'current') => call('read_presentation_baseline', { scope })
  return {
    skill,
    adapter,
    pages,
    inspectPage,
    exportPagePackage,
    call,
    read,
    setDocument: (id: string) => {
      documentId = id
    },
    setContext: (next: typeof context) => {
      context = next
    },
    getContext: () => context,
    setPackage: (base64: string) => {
      packageBase64 = base64
    },
  }
}
async function complexPagePackage(cell = 'North') {
  const zip = new JSZip()
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>${cell}</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>`,
  )
  return zip.generateAsync({ type: 'base64' })
}
async function chartPagePackage() {
  const zip = new JSZip()
  zip.file(
    'ppt/slides/slide1.xml',
    '<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:c="urn:c" xmlns:r="urn:r"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Chart"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rId5"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>',
  )
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    '<Relationships><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>',
  )
  zip.file(
    'ppt/charts/chart1.xml',
    '<c:chartSpace xmlns:c="urn:c"><c:chart><c:plotArea><c:barChart><c:ser><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>12</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>',
  )
  return zip.generateAsync({ type: 'base64' })
}
async function notesPagePackage(
  note = 'First &amp; second',
  target = '../notesSlides/notesSlide1.xml',
) {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', '<p:sld><p:cSld><p:spTree/></p:cSld></p:sld>')
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    `<Relationships><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="${target}"/></Relationships>`,
  )
  zip.file(
    'ppt/notesSlides/notesSlide1.xml',
    `<p:notes><p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:nvPr><p:ph type="ftr"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Footer ignored</a:t></a:r></a:p></p:txBody></p:sp>
    <p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>${note}</a:t></a:r></a:p><a:p><a:r><a:t>Second paragraph</a:t></a:r></a:p></p:txBody></p:sp>
  </p:spTree></p:cSld></p:notes>`,
  )
  return zip.generateAsync({ type: 'base64' })
}
it('captures arbitrary current/selected/deck scopes without a generation artifact or a host write', async () => {
  const f = fixture()
  const result = await f.read()
  expect(result.isError, result.output).not.toBe(true)
  const b = JSON.parse(result.output)
  expect(b.scope).toEqual({ kind: 'current', slideIds: ['s2'] })
  expect(b.pages[0].shapes[0].text).toBe('original')
  expect(b.coverage.notes).toBe('not_read')
  expect(b.coverage.sources).toBe('not_read')
  expect(b.qaPassed).toBe(false)
  expect(b.contentDigest).toMatch(/^[a-f0-9]{64}$/)
  expect(JSON.parse((await f.read('selected')).output).scope).toEqual({
    kind: 'selected',
    slideIds: ['s2'],
    shapeIds: ['shape'],
  })
  expect(JSON.parse((await f.read('deck')).output).scope.slideIds).toEqual(['s1', 's2'])
  expect(f.inspectPage).not.toHaveBeenCalled()
})
it('reports manual edits and selection/order drift without silently replacing the baseline', async () => {
  const f = fixture(),
    b = JSON.parse((await f.read('deck')).output)
  expect(
    JSON.parse((await f.call('check_presentation_baseline', { baseline_id: b.baselineId })).output)
      .unchanged,
  ).toBe(true)
  f.pages.set('s1', page('s1', 'manual change'))
  f.setContext({ ...f.getContext(), slideIds: ['s2', 's1'], selectedShapeIds: [] })
  const r = JSON.parse(
    (await f.call('check_presentation_baseline', { baseline_id: b.baselineId })).output,
  )
  expect(r).toMatchObject({
    unchanged: false,
    changedSlideIds: ['s1'],
    orderChanged: true,
    selectionChanged: true,
  })
  expect(
    JSON.parse((await f.call('check_presentation_baseline', { baseline_id: b.baselineId })).output)
      .unchanged,
  ).toBe(false)
})
it('rejects empty current selection, overlarge scope, forged IDs and unsupported input', async () => {
  const f = fixture()
  f.setContext({ ...f.getContext(), selectedSlideIds: [], selectedShapeIds: [] })
  expect((await f.read()).output).toBe('presentation_selection_empty')
  f.setContext({ ...f.getContext(), slideIds: Array.from({ length: 21 }, (_, i) => `s${i}`) })
  expect((await f.read('deck')).output).toBe('presentation_baseline_scope_limit')
  expect(
    (await f.call('read_presentation_baseline', { scope: 'current', extra: true })).output,
  ).toBe('invalid_tool_input')
  expect((await f.call('check_presentation_baseline', { baseline_id: 'forged' })).output).toBe(
    'presentation_baseline_missing',
  )
})
it('discards a torn read and a read across Save As', async () => {
  const f = fixture()
  f.adapter.readPage.mockImplementation(async (id) => {
    const value = structuredClone(f.pages.get(id)!)
    f.pages.set(id, page(id, value.shapes[0]!.text + '!'))
    return value
  })
  expect((await f.read()).output).toBe('presentation_baseline_changed')
  f.adapter.readPage.mockImplementation(async (id) => {
    f.setDocument('other')
    return page(id)
  })
  expect((await f.read()).output).toBe('presentation_document_changed')
})
it('clear and cancellation suppress late reads and invalidate saved baselines', async () => {
  const f = fixture(),
    b = JSON.parse((await f.read()).output)
  f.skill.clear()
  expect((await f.call('check_presentation_baseline', { baseline_id: b.baselineId })).isError).toBe(
    true,
  )
  let release!: (v: ReturnType<typeof f.getContext>) => void
  f.adapter.readContext.mockImplementationOnce(
    () =>
      new Promise((resolve) => {
        release = resolve
      }),
  )
  const pending = f.read()
  await vi.waitFor(() => expect(release).toBeDefined())
  f.skill.clear()
  release(f.getContext())
  expect((await pending).output).toBe('cancelled')
  const controller = new AbortController()
  controller.abort()
  expect((await f.call('read_presentation_baseline', {}, controller.signal)).output).toBe(
    'cancelled',
  )
})
it('captures a screenshot by exact baseline page ID and refuses stale or out-of-scope content', async () => {
  const f = fixture(),
    b = JSON.parse((await f.read()).output)
  const r = await f.call('read_presentation_baseline_page', {
    baseline_id: b.baselineId,
    slide_id: 's2',
  })
  expect(r.isError, r.output).not.toBe(true)
  expect(r.modelContent?.[0]?.type).toBe('image')
  expect(f.inspectPage).toHaveBeenCalledWith('s2', undefined)
  expect(
    (await f.call('read_presentation_baseline_page', { baseline_id: b.baselineId, slide_id: 's1' }))
      .isError,
  ).toBe(true)
  f.pages.set('s2', page('s2', 'manual'))
  expect(
    (await f.call('read_presentation_baseline_page', { baseline_id: b.baselineId, slide_id: 's2' }))
      .output,
  ).toBe('presentation_baseline_changed')
})
it.each(['png', 'geometry'])(
  'rejects an invalid or mismatched screenshot response (%s)',
  async (kind) => {
    const f = fixture(),
      b = JSON.parse((await f.read()).output)
    const original = await f.inspectPage('s2')
    if (kind === 'png') original.screenshot.base64 = 'iVBORw0KGgo_invalid'
    else original.shapes[0]!.left++
    f.inspectPage.mockResolvedValue(original)
    expect(
      (
        await f.call('read_presentation_baseline_page', {
          baseline_id: b.baselineId,
          slide_id: 's2',
        })
      ).isError,
    ).toBe(true)
  },
)
it('keeps the latest completed request and drops a superseded concurrent read', async () => {
  const f = fixture()
  let release!: (v: ReturnType<typeof f.getContext>) => void
  f.adapter.readContext.mockImplementationOnce(
    () =>
      new Promise((resolve) => {
        release = resolve
      }),
  )
  const pending = f.read()
  await vi.waitFor(() => expect(release).toBeDefined())
  const latest = JSON.parse((await f.read('deck')).output)
  release(f.getContext())
  expect((await pending).output).toBe('cancelled')
  expect(
    JSON.parse(
      (await f.call('check_presentation_baseline', { baseline_id: latest.baselineId })).output,
    ).unchanged,
  ).toBe(true)
})
it('bounds combined page output and refuses stale document IDs', async () => {
  const f = fixture()
  const b = JSON.parse((await f.read()).output)
  f.setDocument('save-as')
  expect((await f.call('check_presentation_baseline', { baseline_id: b.baselineId })).output).toBe(
    'presentation_document_changed',
  )
  f.pages.set('s2', page('s2', '中'.repeat(100_000)))
  expect((await f.read()).output).toBe('presentation_baseline_size_limit')
})
it('reports deleted baseline pages instead of retargeting by slide index', async () => {
  const f = fixture(),
    b = JSON.parse((await f.read()).output)
  f.setContext({
    ...f.getContext(),
    slideIds: ['s1'],
    selectedSlideIds: ['s1'],
    selectedShapeIds: [],
  })
  expect(
    JSON.parse((await f.call('check_presentation_baseline', { baseline_id: b.baselineId })).output),
  ).toMatchObject({ unchanged: false, changedSlideIds: ['s2'], orderChanged: true })
})
it('preserves unsupported theme status and detects a changed master snapshot', async () => {
  const f = fixture()
  let theme = 'black'
  const skill = createPresentationBaselineSkill({
    adapter: f.adapter,
    documentId: async () => 'doc',
    readMasters: async () => ({
      masters: [
        {
          id: 'm',
          name: 'master',
          background: { type: 'solid', color: theme },
          themeColors: { Accent1: theme },
          layouts: [],
        },
      ],
    }),
  })
  const b = JSON.parse(
    (await skill.executeTool({ id: 'read', name: 'read_presentation_baseline', input: {} })).output,
  )
  expect(b.coverage.theme).toBe('read')
  theme = 'white'
  const r = await skill.executeTool({
    id: 'check',
    name: 'check_presentation_baseline',
    input: { baseline_id: b.baselineId },
  })
  expect(JSON.parse(r.output)).toMatchObject({ unchanged: false, stylesChanged: true })
})
it('reads bounded native complex objects only from the exact scoped baseline slide', async () => {
  const f = fixture()
  f.setPackage(await complexPagePackage())
  const baseline = JSON.parse((await f.read()).output)
  const result = await f.call('read_presentation_baseline_complex_page', {
    baseline_id: baseline.baselineId,
    slide_id: 's2',
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({
    baselineId: baseline.baselineId,
    slideId: 's2',
    tables: [{ shapeId: '7', rows: [['North']] }],
    charts: [],
    cacheOnly: true,
    qaPassed: false,
  })
  expect(result.mutated).toBe(false)
  expect(f.exportPagePackage).toHaveBeenCalledTimes(2)
  expect(f.exportPagePackage).toHaveBeenCalledWith('s2', undefined)
  expect(
    (
      await f.call('read_presentation_baseline_complex_page', {
        baseline_id: baseline.baselineId,
        slide_id: 's1',
      })
    ).output,
  ).toBe('presentation_baseline_scope_mismatch')
})
it('rejects complex page reads when the package or host baseline changes during export', async () => {
  const f = fixture()
  const first = await complexPagePackage('North')
  const changed = await complexPagePackage('South')
  f.setPackage(first)
  const baseline = JSON.parse((await f.read()).output)
  f.exportPagePackage
    .mockImplementationOnce(async (slideId) => ({
      slideId,
      slideIds: [...f.getContext().slideIds],
      base64: first,
    }))
    .mockImplementationOnce(async (slideId) => ({
      slideId,
      slideIds: [...f.getContext().slideIds],
      base64: changed,
    }))
  expect(
    (
      await f.call('read_presentation_baseline_complex_page', {
        baseline_id: baseline.baselineId,
        slide_id: 's2',
      })
    ).output,
  ).toBe('presentation_baseline_changed')
  f.exportPagePackage.mockReset()
  f.exportPagePackage.mockImplementation(async (slideId) => {
    f.pages.set('s2', page('s2', 'manual change'))
    return { slideId, slideIds: [...f.getContext().slideIds], base64: first }
  })
  expect(
    (
      await f.call('read_presentation_baseline_complex_page', {
        baseline_id: baseline.baselineId,
        slide_id: 's2',
      })
    ).output,
  ).toBe('presentation_baseline_changed')
})
it('reads one exact scoped chart source without treating cache as verified data', async () => {
  const f = fixture()
  f.pages.get('s2')!.shapes[0]!.type = 'Chart'
  f.pages.get('s2')!.shapes[0]!.id = '7'
  f.setContext({ ...f.getContext(), selectedShapeIds: ['7'] })
  f.setPackage(await chartPagePackage())
  const baseline = JSON.parse((await f.read()).output)
  const result = await f.call('read_presentation_baseline_chart_source', {
    baseline_id: baseline.baselineId,
    slide_id: 's2',
    shape_id: '7',
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({
    baselineId: baseline.baselineId,
    slideId: 's2',
    shapeId: '7',
    sourceKind: 'cache_only',
    verification: 'not_verified',
    qaPassed: false,
    writeAuthorized: false,
  })
  expect(result.mutated).toBe(false)
  expect(f.exportPagePackage).toHaveBeenCalledTimes(2)
  expect(
    (
      await f.call('read_presentation_baseline_chart_source', {
        baseline_id: baseline.baselineId,
        slide_id: 's1',
        shape_id: '7',
      })
    ).output,
  ).toBe('presentation_baseline_scope_mismatch')
})
it('refuses chart source evidence if the chart package changes during the read', async () => {
  const f = fixture()
  f.pages.get('s2')!.shapes[0]!.type = 'Chart'
  f.pages.get('s2')!.shapes[0]!.id = '7'
  f.setContext({ ...f.getContext(), selectedShapeIds: ['7'] })
  const before = await chartPagePackage()
  const edited = await JSZip.loadAsync(before, { base64: true })
  edited.file(
    'ppt/charts/chart1.xml',
    (await edited.file('ppt/charts/chart1.xml')!.async('string')).replace('12', '13'),
  )
  const after = await edited.generateAsync({ type: 'base64' })
  f.setPackage(before)
  const baseline = JSON.parse((await f.read()).output)
  f.exportPagePackage
    .mockImplementationOnce(async (slideId) => ({
      slideId,
      slideIds: ['s1', 's2'],
      base64: before,
    }))
    .mockImplementationOnce(async (slideId) => ({ slideId, slideIds: ['s1', 's2'], base64: after }))
  expect(
    (
      await f.call('read_presentation_baseline_chart_source', {
        baseline_id: baseline.baselineId,
        slide_id: 's2',
        shape_id: '7',
      })
    ).output,
  ).toBe('presentation_baseline_changed')
})
it('reads bounded notes from an exact stable baseline page without treating them as a source', async () => {
  const f = fixture()
  f.setPackage(await notesPagePackage())
  const baseline = JSON.parse((await f.read()).output)
  const result = await f.call('read_presentation_baseline_notes', {
    baseline_id: baseline.baselineId,
    slide_id: 's2',
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({
    baselineId: baseline.baselineId,
    slideId: 's2',
    status: 'read',
    text: 'First & second\nSecond paragraph',
    sourceVerified: false,
    qaPassed: false,
    writeAuthorized: false,
  })
  expect(result.mutated).toBe(false)
  expect(f.exportPagePackage).toHaveBeenCalledTimes(2)
  expect(
    (
      await f.call('read_presentation_baseline_notes', {
        baseline_id: baseline.baselineId,
        slide_id: 's1',
      })
    ).output,
  ).toBe('presentation_baseline_scope_mismatch')
})
it('reads notes from a real compiled one-page PowerPoint package', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const packageBase64 = Buffer.from((await compilePresentationDeck(deck)).bytes).toString('base64')
  const f = fixture()
  f.setPackage(packageBase64)
  const baseline = JSON.parse((await f.read()).output)
  const result = await f.call('read_presentation_baseline_notes', {
    baseline_id: baseline.baselineId,
    slide_id: 's2',
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({ status: 'read', sourceVerified: false })
  expect(JSON.parse(result.output).text).toContain('演讲备注')
})
it('distinguishes an absent notes relationship from an empty body', async () => {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', '<p:sld><p:cSld/></p:sld>')
  expect(await inspectPowerPointPageNotes(await zip.generateAsync({ type: 'base64' }))).toEqual({
    status: 'not_present',
    text: '',
  })
})
it('rejects changed and unsafe notes packages instead of returning stale text', async () => {
  const f = fixture()
  const before = await notesPagePackage()
  const after = await notesPagePackage('Changed')
  f.setPackage(before)
  const baseline = JSON.parse((await f.read()).output)
  f.exportPagePackage
    .mockImplementationOnce(async (slideId) => ({
      slideId,
      slideIds: ['s1', 's2'],
      base64: before,
    }))
    .mockImplementationOnce(async (slideId) => ({ slideId, slideIds: ['s1', 's2'], base64: after }))
  expect(
    (
      await f.call('read_presentation_baseline_notes', {
        baseline_id: baseline.baselineId,
        slide_id: 's2',
      })
    ).output,
  ).toBe('presentation_baseline_changed')
  f.setPackage(await notesPagePackage('Text', '../../outside.xml'))
  expect(
    (
      await f.call('read_presentation_baseline_notes', {
        baseline_id: baseline.baselineId,
        slide_id: 's2',
      })
    ).output,
  ).toBe('office_api_unsupported')
})
