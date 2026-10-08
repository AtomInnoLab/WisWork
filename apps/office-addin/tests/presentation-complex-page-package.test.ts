import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  inspectPowerPointComplexPagePackage,
  inspectPowerPointTableCellPackage,
  inspectPowerPointTableCellsPackage,
} from '../src/skills/powerpoint/presentation-complex-page-package.js'

async function packageWith(slide: string, rels = '', chart = ''): Promise<string> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', slide)
  if (rels) zip.file('ppt/slides/_rels/slide1.xml.rels', rels)
  if (chart) zip.file('ppt/charts/chart1.xml', chart)
  return zip.generateAsync({ type: 'base64' })
}

const table =
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>North</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>42</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
const chartFrame =
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8" name="Chart"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rId5"/></a:graphicData></a:graphic></p:graphicFrame>'
const slide = (content: string) =>
  `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:c="urn:c" xmlns:r="urn:r"><p:cSld><p:spTree>${content}</p:spTree></p:cSld></p:sld>`
const rels =
  '<Relationships><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>'
const chart =
  '<c:chartSpace xmlns:c="urn:c"><c:chart><c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>34</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'

describe('complex PowerPoint page package read', () => {
  it('reads native table cells and chart cached series without claiming source data', async () => {
    const result = await inspectPowerPointComplexPagePackage(
      await packageWith(slide(table + chartFrame), rels, chart),
    )
    expect(result).toMatchObject({
      tables: [{ shapeId: '7', rows: [['North', '42']], simpleCells: [[true, true]] }],
      charts: [
        {
          shapeId: '8',
          plotTypes: ['barChart'],
          cacheOnly: true,
          series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: ['12', '34'] }],
        },
      ],
      truncated: false,
    })
  })

  it('reads bounded explicit chart direction, grouping, legend and value-label options', async () => {
    const styled = chart
      .replace(
        '<c:barChart>',
        '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:dLbls><c:showVal val="1"/></c:dLbls>',
      )
      .replace('</c:chart>', '<c:legend><c:legendPos val="b"/></c:legend></c:chart>')
    const result = await inspectPowerPointComplexPagePackage(
      await packageWith(slide(chartFrame), rels, styled),
    )
    expect(result.charts[0]?.visualOptions).toEqual({
      barDirections: ['col'],
      groupings: ['clustered'],
      legendPositions: ['b'],
      valueLabels: ['1'],
    })
  })

  it('rejects external, traversing, and non-chart relationships', async () => {
    for (const replacement of [
      'TargetMode="External" Target="https://example.com/chart.xml"',
      'Target="../../evil.xml"',
      'Type="other" Target="../charts/chart1.xml"',
    ]) {
      const altered = rels.replace(/Type="[^"]+" Target="[^"]+"/, replacement)
      await expect(
        inspectPowerPointComplexPagePackage(await packageWith(slide(chartFrame), altered, chart)),
      ).rejects.toThrow('office_api_unsupported')
    }
  })

  it('rejects malformed XML and bounds returned cell text', async () => {
    await expect(inspectPowerPointComplexPagePackage(await packageWith('<p:sld>'))).rejects.toThrow(
      'office_api_unsupported',
    )
    const long = table.replace('North', 'x'.repeat(500))
    const result = await inspectPowerPointComplexPagePackage(await packageWith(slide(long)))
    expect(result.tables[0]?.rows[0]?.[0].length).toBeLessThanOrEqual(128)
    expect(result.truncated).toBe(true)
  })
  it('marks table cells in a merged table ineligible for coordinate edits', async () => {
    const merged = table.replace('<a:tc>', '<a:tc gridSpan="2">')
    const result = await inspectPowerPointComplexPagePackage(await packageWith(slide(merged)))
    expect(result.tables[0]?.simpleCells).toEqual([[false, false]])
    await expect(
      inspectPowerPointTableCellPackage(await packageWith(slide(merged)), '7', 0, 0),
    ).rejects.toThrow('presentation_existing_target_unsupported')
  })
  it('binds a reversible cell to its entire table layout and other cell contents', async () => {
    const original = await inspectPowerPointTableCellPackage(
      await packageWith(slide(table)),
      '7',
      0,
      0,
    )
    const changedTarget = table.replace('North', 'South')
    const changed = await inspectPowerPointTableCellPackage(
      await packageWith(slide(changedTarget)),
      '7',
      0,
      0,
    )
    expect(changed.structureDigest).toBe(original.structureDigest)
    const inserted = table.replace(
      '<a:tbl>',
      '<a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>North</a:t></a:r></a:p></a:txBody></a:tc></a:tr>',
    )
    const shifted = await inspectPowerPointTableCellPackage(
      await packageWith(slide(inserted)),
      '7',
      0,
      0,
    )
    expect(shifted.text).toBe(original.text)
    expect(shifted.structureDigest).not.toBe(original.structureDigest)
  })

  it('binds a batch to unchanged table content outside all selected cells', async () => {
    const threeCells = table.replace(
      '</a:tr>',
      '<a:tc><a:txBody><a:p><a:r><a:t>Other</a:t></a:r></a:p></a:txBody></a:tc></a:tr>',
    )
    const cells = [
      { rowIndex: 0, columnIndex: 0 },
      { rowIndex: 0, columnIndex: 1 },
    ]
    const original = await inspectPowerPointTableCellsPackage(
      await packageWith(slide(threeCells)),
      '7',
      cells,
    )
    expect(original.cells).toEqual([
      { ...cells[0], text: 'North' },
      { ...cells[1], text: '42' },
    ])
    const changedSelected = threeCells.replace('North', 'South').replace('42', '43')
    expect(
      (
        await inspectPowerPointTableCellsPackage(
          await packageWith(slide(changedSelected)),
          '7',
          cells,
        )
      ).structureDigest,
    ).toBe(original.structureDigest)
    const changedOther = threeCells.replace('Other', 'Else')
    expect(
      (await inspectPowerPointTableCellsPackage(await packageWith(slide(changedOther)), '7', cells))
        .structureDigest,
    ).not.toBe(original.structureDigest)
    const changedTopology = threeCells.replace(
      '</a:tbl>',
      '<a:tr><a:tc><a:txBody><a:p/></a:txBody></a:tc></a:tr></a:tbl>',
    )
    expect(
      (
        await inspectPowerPointTableCellsPackage(
          await packageWith(slide(changedTopology)),
          '7',
          cells,
        )
      ).structureDigest,
    ).not.toBe(original.structureDigest)
  })

  it('rejects repeated coordinates and an invalid batch size', async () => {
    const base64 = await packageWith(slide(table))
    const cell = { rowIndex: 0, columnIndex: 0 }
    await expect(inspectPowerPointTableCellsPackage(base64, '7', [cell, cell])).rejects.toThrow(
      'invalid_tool_input',
    )
    await expect(inspectPowerPointTableCellsPackage(base64, '7', [])).rejects.toThrow(
      'invalid_tool_input',
    )
    await expect(
      inspectPowerPointTableCellsPackage(base64, '7', Array(9).fill(cell)),
    ).rejects.toThrow('invalid_tool_input')
  })

  it('bounds the entire serialized summary even with many large cells', async () => {
    const cell = `<a:tc><a:txBody><a:p><a:r><a:t>${'x'.repeat(128)}</a:t></a:r></a:p></a:txBody></a:tc>`
    const row = `<a:tr>${cell.repeat(12)}</a:tr>`
    const frame = (id: number) =>
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl>${row.repeat(20)}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
    const result = await inspectPowerPointComplexPagePackage(
      await packageWith(slide(Array.from({ length: 8 }, (_, index) => frame(index + 1)).join(''))),
    )
    expect(new TextEncoder().encode(JSON.stringify(result)).byteLength).toBeLessThanOrEqual(
      128 * 1024,
    )
    expect(result.truncated).toBe(true)
  })

  it('marks unsupported multilevel and sparse chart caches as incomplete', async () => {
    const single = chart.replace(
      '<c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache>',
      '<c:multiLvlStrCache><c:lvl><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:lvl></c:multiLvlStrCache>',
    )
    const supported = await inspectPowerPointComplexPagePackage(
      await packageWith(slide(chartFrame), rels, single),
    )
    expect(supported.truncated).toBe(false)
    expect(supported.charts[0]?.series[0]?.categories).toEqual(['Q1', 'Q2'])
    const multi = single.replace(
      '</c:multiLvlStrCache>',
      '<c:lvl><c:pt idx="0"><c:v>2026</c:v></c:pt></c:lvl></c:multiLvlStrCache>',
    )
    const sparse = chart.replace(
      '<c:pt idx="1"><c:v>34</c:v></c:pt>',
      '<c:pt idx="3"><c:v>34</c:v></c:pt>',
    )
    for (const xml of [multi, sparse]) {
      const result = await inspectPowerPointComplexPagePackage(
        await packageWith(slide(chartFrame), rels, xml),
      )
      expect(result.truncated).toBe(true)
      expect(result.charts[0]?.truncated).toBe(true)
    }
  })
})
