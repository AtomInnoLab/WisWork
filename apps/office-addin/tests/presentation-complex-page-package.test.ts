import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { inspectPowerPointComplexPagePackage } from '../src/skills/powerpoint/presentation-complex-page-package.js'

async function packageWith(slide: string, rels = '', chart = ''): Promise<string> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', slide)
  if (rels) zip.file('ppt/slides/_rels/slide1.xml.rels', rels)
  if (chart) zip.file('ppt/charts/chart1.xml', chart)
  return zip.generateAsync({ type: 'base64' })
}

const table = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>North</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>42</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
const chartFrame = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8" name="Chart"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rId5"/></a:graphicData></a:graphic></p:graphicFrame>'
const slide = (content: string) => `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:c="urn:c" xmlns:r="urn:r"><p:cSld><p:spTree>${content}</p:spTree></p:cSld></p:sld>`
const rels = '<Relationships><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>'
const chart = '<c:chartSpace xmlns:c="urn:c"><c:chart><c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>34</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'

describe('complex PowerPoint page package read', () => {
  it('reads native table cells and chart cached series without claiming source data', async () => {
    const result = await inspectPowerPointComplexPagePackage(await packageWith(slide(table + chartFrame), rels, chart))
    expect(result).toMatchObject({
      tables: [{ shapeId: '7', rows: [['North', '42']] }],
      charts: [{ shapeId: '8', cacheOnly: true, series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: ['12', '34'] }] }],
      truncated: false,
    })
  })

  it('rejects external, traversing, and non-chart relationships', async () => {
    for (const replacement of [
      'TargetMode="External" Target="https://example.com/chart.xml"',
      'Target="../../evil.xml"',
      'Type="other" Target="../charts/chart1.xml"',
    ]) {
      const altered = rels.replace(/Type="[^"]+" Target="[^"]+"/, replacement)
      await expect(inspectPowerPointComplexPagePackage(await packageWith(slide(chartFrame), altered, chart))).rejects.toThrow('office_api_unsupported')
    }
  })

  it('rejects malformed XML and bounds returned cell text', async () => {
    await expect(inspectPowerPointComplexPagePackage(await packageWith('<p:sld>'))).rejects.toThrow('office_api_unsupported')
    const long = table.replace('North', 'x'.repeat(500))
    const result = await inspectPowerPointComplexPagePackage(await packageWith(slide(long)))
    expect(result.tables[0]?.rows[0]?.[0].length).toBeLessThanOrEqual(128)
    expect(result.truncated).toBe(true)
  })
})
