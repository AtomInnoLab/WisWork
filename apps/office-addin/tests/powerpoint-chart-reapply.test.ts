import { expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { type PowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter.js'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill.js'
const png = 'iVBORw0KGgoAAAA='

function adapter(overrides: Partial<PowerPointAdapter> = {}): PowerPointAdapter {
  return {
    inspectSlideMasters: vi.fn().mockResolvedValue({
      masters: [
        {
          id: 'master-1',
          name: 'Main',
          background: { type: 'Solid', color: '#FFFFFF', transparency: 0 },
          themeColors: { Light1: '#FFFFFF', Dark1: '#000000' },
          layouts: [
            {
              id: 'layout-1',
              name: 'Title',
              isMasterBackgroundFollowed: true,
              areBackgroundGraphicsHidden: false,
              background: { type: 'Solid' },
            },
          ],
        },
      ],
    }),
    executeMasterOperations: vi.fn().mockResolvedValue(undefined),
    screenshotSlide: vi
      .fn()
      .mockResolvedValue({ slideId: 'host-slide-1', mime: 'image/png', base64: png }),
    listSlideShapes: vi.fn().mockResolvedValue({
      slideId: 'slide-1',
      slideIndex: 0,
      shapes: [
        { id: '2', name: 'Title', type: 'TextBox', left: 10, top: 20, width: 200, height: 40 },
      ],
    }),
    readSlideText: vi.fn().mockResolvedValue({
      slideId: 'slide-1',
      shapeId: '2',
      text: 'Hello',
      paragraphs: ['Hello'],
    }),
    readSlideTable: vi.fn().mockResolvedValue([
      ['方案', '结果'],
      ['甲', '120'],
    ]),
    verifySlides: vi.fn().mockResolvedValue({
      slideWidth: 960,
      slideHeight: 540,
      slides: [],
    }),
    snapshotSlide: vi.fn().mockResolvedValue({ slideId: 'slide-1', fingerprint: 'slide-1:1' }),
    editSlideText: vi.fn().mockResolvedValue(undefined),
    duplicateSlide: vi.fn().mockResolvedValue({ slideId: 'slide-copy' }),
    exportSlidePackage: vi.fn().mockRejectedValue(new Error('office_api_unsupported')),
    replaceSlidePackage: vi.fn().mockRejectedValue(new Error('office_api_unsupported')),
    executeDeclarative: vi.fn().mockRejectedValue(new Error('office_api_unsupported')),
    ...overrides,
  }
}

const call = (name: string, input: Record<string, unknown> = {}) => ({ id: 'call-1', name, input })

async function fixture() {
  const book = new JSZip()
  book.file(
    'xl/workbook.xml',
    '<workbook><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>',
  )
  book.file(
    'xl/_rels/workbook.xml.rels',
    '<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
  )
  book.file(
    'xl/worksheets/sheet1.xml',
    '<worksheet><sheetData><row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><v>1</v></c></row></sheetData></worksheet>',
  )
  const zip = new JSZip()
  zip.file(
    'ppt/slides/slide1.xml',
    '<p:sld><p:graphicFrame><p:cNvPr id="8"/><c:chart r:id="rId5"/></p:graphicFrame></p:sld>',
  )
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    '<Relationships><Relationship Id="rId5" Type="x/chart" Target="../charts/chart1.xml"/></Relationships>',
  )
  zip.file(
    'ppt/charts/chart1.xml',
    '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:ser><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$2</c:f><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$2</c:f><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart><c:externalData r:id="rId9"/></c:chartSpace>',
  )
  zip.file(
    'ppt/charts/_rels/chart1.xml.rels',
    '<Relationships><Relationship Id="rId9" Type="x/package" Target="../embeddings/Book1.xlsx"/></Relationships>',
  )
  zip.file('ppt/embeddings/Book1.xlsx', await book.generateAsync({ type: 'uint8array' }))
  const original = await zip.generateAsync({ type: 'base64' })
  let current = original
  let activeSlideId = 's1'
  let replacementCount = 0
  let otherSlideId = 's2'
  const backups = new Map<string, { bytes: Uint8Array; meta: Record<string, unknown> }>()
  const records = new Map<
    string,
    import('../src/skills/powerpoint/presentation-existing-chart.js').PresentationExistingChartChange
  >()
  const chartSavepoint = {
    documentId: async () => 'doc-1',
    readExistingChartChange: (id: string) => records.get(id),
    writeExistingChartChange: vi.fn(
      async (
        record: import('../src/skills/powerpoint/presentation-existing-chart.js').PresentationExistingChartChange,
      ) => {
        records.set(record.changeId, structuredClone(record))
      },
    ),
    request: async (body: unknown) => {
      const input = body as Record<string, unknown>
      const id = input.backupId as string
      if (input.operation === 'existing_page_backup_begin' && !backups.has(id))
        backups.set(id, {
          bytes: new Uint8Array(),
          meta: { ...input, status: 'uploading', receivedBytes: 0 },
        })
      const backup = backups.get(id)
      if (!backup) return new Response(JSON.stringify({ error: 'not_found' }), { status: 404 })
      if (input.operation === 'existing_page_backup_chunk') {
        const bytes = Uint8Array.from(atob(input.base64 as string), (char) => char.charCodeAt(0))
        backup.bytes = Uint8Array.from([...backup.bytes, ...bytes])
        backup.meta.receivedBytes = backup.bytes.length
      }
      if (input.operation === 'existing_page_backup_finish') backup.meta.status = 'ready'
      if (input.operation === 'existing_page_backup_release') {
        backups.delete(id)
        return new Response(JSON.stringify({ ...input, status: 'released' }))
      }
      if (input.operation === 'existing_page_backup_read') {
        const part = backup.bytes.subarray(
          input.offset as number,
          (input.offset as number) + (input.length as number),
        )
        return new Response(
          JSON.stringify({
            backupId: id,
            offset: input.offset,
            sizeBytes: backup.meta.sizeBytes,
            sha256: backup.meta.sha256,
            base64: btoa(String.fromCharCode(...part)),
          }),
        )
      }
      return new Response(JSON.stringify(backup.meta))
    },
  }
  const fake = adapter({
    verifySlides: vi.fn().mockImplementation(() =>
      Promise.resolve({
        slideWidth: 960,
        slideHeight: 540,
        slides: [{ slideId: activeSlideId }, { slideId: otherSlideId }],
      }),
    ),
    exportSlidePackage: vi
      .fn()
      .mockImplementation(() =>
        Promise.resolve({ slideId: activeSlideId, base64: current, fingerprint: activeSlideId }),
      ),
    replaceSlidePackage: vi.fn().mockImplementation((_index, base64) => {
      current = base64
      activeSlideId = `replacement-${++replacementCount}`
      return Promise.resolve({ slideId: activeSlideId })
    }),
  })
  const proposals = createStructuredProposalController()
  const create = () => createPowerPointSkill({ adapter: fake, proposals, chartSavepoint })
  let skill = create()
  const run = (name: string, changeId: string) =>
    skill.executeTool(call(name, { change_id: changeId }))
  const confirm = () => proposals.confirm(proposals.pending()!.id)
  await skill.executeTool(
    call('update_slide_chart_values', { slide_index: 0, shape_id: '8', values: [['5']] }),
  )
  await confirm()
  const changeId = [...records.keys()][0]!
  await run('undo_slide_chart_values_change', changeId)
  await confirm()
  return {
    records,
    changeId,
    fake,
    proposals,
    run,
    confirm,
    backups,
    chartSavepoint,
    setOtherSlideId: (value: string) => {
      otherSlideId = value
    },
    reopen: () => {
      skill = create()
    },
    setSlideId: (value: string) => {
      activeSlideId = value
    },
    setPackage: (value: string) => {
      current = value
    },
    original,
  }
}
it('reapplies chart values after reopen into independent history and can undo again', async () => {
  const f = await fixture()
  const originalRecord = structuredClone(f.records.get(f.changeId)!)
  f.reopen()
  const result = await f.run('reapply_slide_chart_values_change', f.changeId)
  expect(result.isError, result.output).not.toBe(true)
  await f.confirm()
  expect(f.records.get(f.changeId)).toEqual(originalRecord)
  const replay = [...f.records.values()].find((r) => r.changeId !== f.changeId)!
  expect(replay).toMatchObject({
    state: 'applied',
    values: [['5']],
    reapplies: f.changeId,
    afterPackageDigest: originalRecord.afterPackageDigest,
  })
  f.reopen()
  const undo = await f.run('undo_slide_chart_values_change', replay.changeId)
  expect(undo.isError, undo.output).not.toBe(true)
  await f.confirm()
  expect(f.records.get(replay.changeId)?.state).toBe('undone')
})

it.each(['legacy', 'released', 'missing', 'tampered'] as const)(
  'refuses chart reapply with %s source evidence',
  async (kind) => {
    const f = await fixture()
    const record = f.records.get(f.changeId)!
    if (kind === 'legacy') delete record.values
    if (kind === 'released') record.backupReleasedAt = '2026-09-29T00:00:00.000Z'
    if (kind === 'missing') f.backups.delete(record.backup.backupId)
    if (kind === 'tampered') f.backups.get(record.backup.backupId)!.bytes[0] ^= 1
    const result = await f.run('reapply_slide_chart_values_change', f.changeId)
    expect(result.isError).toBe(true)
    expect(f.records.size).toBe(1)
    expect(f.fake.replaceSlidePackage).toHaveBeenCalledTimes(2)
  },
)
it.each(['identity', 'page', 'order'] as const)(
  'rejects a restored chart %s conflict without creating history',
  async (kind) => {
    const f = await fixture()
    if (kind === 'identity') f.setSlideId('manual-copy')
    if (kind === 'page') {
      const zip = await JSZip.loadAsync(f.original, { base64: true })
      zip.file('manual.txt', 'manual page change')
      f.setPackage(await zip.generateAsync({ type: 'base64' }))
    }
    if (kind === 'order') f.setOtherSlideId('manual-other')
    const result = await f.run('reapply_slide_chart_values_change', f.changeId)
    expect(result.isError).toBe(true)
    expect(f.records.size).toBe(1)
  },
)
it.each(['history', 'released', 'missing', 'identity'] as const)(
  'rejects chart %s changes after proposing reapply',
  async (kind) => {
    const f = await fixture()
    const result = await f.run('reapply_slide_chart_values_change', f.changeId)
    expect(result.isError, result.output).not.toBe(true)
    const record = f.records.get(f.changeId)!
    if (kind === 'history') record.values = [['6']]
    if (kind === 'released') record.backupReleasedAt = '2026-09-29T00:00:00.000Z'
    if (kind === 'missing') f.backups.delete(record.backup.backupId)
    if (kind === 'identity') f.setSlideId('manual-copy')
    await expect(f.confirm()).rejects.toThrow()
    expect(f.records.size).toBe(1)
    expect(f.fake.replaceSlidePackage).toHaveBeenCalledTimes(2)
  },
)
it('rejects stale retained values whose recomputed after digest differs', async () => {
  const f = await fixture()
  f.records.get(f.changeId)!.values = [['6']]
  const result = await f.run('reapply_slide_chart_values_change', f.changeId)
  expect(result.isError).toBe(true)
  expect(f.records.size).toBe(1)
})
it('recovers a lost chart reapply receipt after reopen without replacing the host again', async () => {
  const f = await fixture()
  const originalRecord = structuredClone(f.records.get(f.changeId)!)
  await f.run('reapply_slide_chart_values_change', f.changeId)
  f.chartSavepoint.writeExistingChartChange.mockImplementation(async (record) => {
    if (record.state === 'applied') throw new Error('receipt_lost')
    f.records.set(record.changeId, structuredClone(record))
  })
  await expect(f.confirm()).rejects.toThrow('receipt_lost')
  const replay = [...f.records.values()].find((r) => r.changeId !== f.changeId)!
  expect(replay.state).toBe('write_pending')
  f.chartSavepoint.writeExistingChartChange.mockImplementation(async (record) => {
    f.records.set(record.changeId, structuredClone(record))
  })
  f.reopen()
  const resume = await f.run('resume_slide_chart_values_change', replay.changeId)
  expect(resume.isError, resume.output).not.toBe(true)
  await f.confirm()
  expect(f.records.get(replay.changeId)?.state).toBe('applied')
  expect(f.fake.replaceSlidePackage).toHaveBeenCalledTimes(3)
  expect(f.records.get(f.changeId)).toEqual(originalRecord)
  await f.run('undo_slide_chart_values_change', replay.changeId)
  await f.confirm()
  expect(f.records.get(replay.changeId)?.state).toBe('undone')
})

it('exposes chart reapply only when durable chart savepoints are available', () => {
  const plain = createPowerPointSkill({
    adapter: adapter(),
    proposals: createStructuredProposalController(),
  })
  expect(plain.tools.some((tool) => tool.name === 'reapply_slide_chart_values_change')).toBe(false)
})
