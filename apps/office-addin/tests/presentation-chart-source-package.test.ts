import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { inspectPowerPointChartSourcePackage, updatePowerPointChartDataPackage } from '../src/skills/powerpoint/presentation-chart-source-package.js'

const slide = '<p:sld><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8" name="Chart"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rId5"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>'
const slideRels = '<Relationships><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>'
const chartRels = (target: string, mode = '') => `<Relationships><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="${target}" ${mode}/></Relationships>`
const chart = (formula = 'Sheet1!$A$2:$A$3', cached = '2', linked = true) => `<c:chartSpace><c:chart><c:plotArea><c:barChart><c:ser><c:cat><c:strRef><c:f>${formula}</c:f><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>${cached}</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart>${linked ? '<c:externalData r:id="rId9"/>' : ''}</c:chartSpace>`
async function workbook(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('xl/workbook.xml', '<workbook><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>')
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
  zip.file('xl/sharedStrings.xml', '<sst><si><t>Q1</t></si><si><t>Q2</t></si></sst>')
  zip.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2"><v>1</v></c></row><row r="3"><c r="A3" t="s"><v>1</v></c><c r="B3"><v>2</v></c></row></sheetData></worksheet>')
  return zip.generateAsync({type:'uint8array'})
}
async function pptx(options: { chart?: string; target?: string; mode?: string; slideRels?: string; includeWorkbook?: boolean } = {}): Promise<string> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', slide)
  zip.file('ppt/slides/_rels/slide1.xml.rels', options.slideRels ?? slideRels)
  zip.file('ppt/charts/chart1.xml', options.chart ?? chart())
  zip.file('ppt/charts/_rels/chart1.xml.rels', chartRels(options.target ?? '../embeddings/Book1.xlsx', options.mode))
  if (options.includeWorkbook !== false) zip.file('ppt/embeddings/Book1.xlsx', await workbook())
  return zip.generateAsync({ type:'base64' })
}
describe('PowerPoint chart source inspection', () => {
  it('verifies matching embedded workbook ranges and returns a digest', async () => {
    const result = await inspectPowerPointChartSourcePackage(await pptx(), '8')
    expect(result).toMatchObject({ shapeId:'8', sourceKind:'embedded_xlsx', verification:'matches', series:[{categories:['Q1','Q2'], values:['1','2']}] })
    expect(result.sourceDigest).toMatch(/^[a-f0-9]{64}$/)
  })
  it('identifies source/cache mismatch', async () => {
    expect(await inspectPowerPointChartSourcePackage(await pptx({chart:chart(undefined,'9')}), '8')).toMatchObject({sourceKind:'embedded_xlsx', verification:'mismatch'})
  })
  it('does not verify a cache whose declared point count differs from its points', async () => {
    const malformed = chart().replace('<c:strCache>', '<c:strCache><c:ptCount val="3"/>')
    expect(await inspectPowerPointChartSourcePackage(await pptx({ chart: malformed }), '8')).toMatchObject({ sourceKind:'embedded_xlsx', verification:'not_verified', reason:'unsupported_formula_or_cache' })
  })
  it('classifies external links without fetching them', async () => {
    expect(await inspectPowerPointChartSourcePackage(await pptx({target:'https://example.com/Book.xlsx',mode:'TargetMode="External"',includeWorkbook:false}), '8')).toMatchObject({sourceKind:'external_link',verification:'not_verified',reason:'external_source_not_fetched'})
  })
  it('rejects traversal and malformed slide chart relationships', async () => {
    await expect(inspectPowerPointChartSourcePackage(await pptx({target:'../../evil.xlsx'}),'8')).rejects.toThrow('office_api_unsupported')
    await expect(inspectPowerPointChartSourcePackage(await pptx({slideRels:'<Relationships><Relationship Id="rId5" Type="x/chart" Target="../../evil.xml"/></Relationships>'}),'8')).rejects.toThrow('office_api_unsupported')
  })
  it('keeps unsupported formulas unverified and cache-only charts separate', async () => {
    expect(await inspectPowerPointChartSourcePackage(await pptx({chart:chart('Other!A2:A3')}),'8')).toMatchObject({sourceKind:'embedded_xlsx',verification:'not_verified',reason:'unsupported_formula_or_cache'})
    expect(await inspectPowerPointChartSourcePackage(await pptx({chart:chart(undefined,undefined,false)}),'8')).toMatchObject({sourceKind:'cache_only',verification:'not_verified'})
  })
})

describe('PowerPoint chart embedded data update', () => {
  it('updates the workbook and chart cache together', async () => {
    const result = await updatePowerPointChartDataPackage(await pptx(), '8', [['3', '4']])
    expect(result.changedPaths).toEqual(['ppt/charts/chart1.xml', 'ppt/embeddings/Book1.xlsx'])
    expect(await inspectPowerPointChartSourcePackage(result.base64, '8')).toMatchObject({ verification:'matches', series:[{categories:['Q1','Q2'], values:['3','4']}] })
  })
  it('refuses stale cache, external sources, and unsupported formulas', async () => {
    await expect(updatePowerPointChartDataPackage(await pptx({chart:chart(undefined,'9')}), '8', [['3','4']])).rejects.toThrow('office_api_unsupported')
    await expect(updatePowerPointChartDataPackage(await pptx({target:'https://example.com/Book.xlsx',mode:'TargetMode="External"',includeWorkbook:false}), '8', [['3','4']])).rejects.toThrow('office_api_unsupported')
    await expect(updatePowerPointChartDataPackage(await pptx({chart:chart('Other!A2:A3')}), '8', [['3','4']])).rejects.toThrow('office_api_unsupported')
  })
  it('refuses shape mismatches, count changes and nonnumeric values', async () => {
    const source = await pptx()
    await expect(updatePowerPointChartDataPackage(source, '9', [['3','4']])).rejects.toThrow('office_api_unsupported')
    await expect(updatePowerPointChartDataPackage(source, '8', [['3']])).rejects.toThrow('invalid_tool_input')
    await expect(updatePowerPointChartDataPackage(source, '8', [['3','NaN']])).rejects.toThrow('invalid_tool_input')
  })
  it('refuses an embedded workbook shared by another chart', async () => {
    const zip = await JSZip.loadAsync(await pptx(), { base64: true })
    zip.file('ppt/charts/_rels/chart2.xml.rels', chartRels('../embeddings/Book1.xlsx'))
    await expect(updatePowerPointChartDataPackage(await zip.generateAsync({ type:'base64' }), '8', [['3','4']])).rejects.toThrow('office_api_unsupported')
    zip.file('ppt/charts/_rels/chart2.xml.rels', chartRels('../embeddings/./Book1.xlsx'))
    await expect(updatePowerPointChartDataPackage(await zip.generateAsync({ type:'base64' }), '8', [['3','4']])).rejects.toThrow('office_api_unsupported')
  })
  it('refuses a chart part shared by two shapes on the same slide', async () => {
    const zip = await JSZip.loadAsync(await pptx(), { base64: true })
    zip.file('ppt/slides/slide1.xml', slide.replace('</p:spTree>', '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Alias"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rId5"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree>'))
    await expect(updatePowerPointChartDataPackage(await zip.generateAsync({ type:'base64' }), '8', [['3','4']])).rejects.toThrow('office_api_unsupported')
    zip.file('ppt/slides/slide1.xml', slide.replace('</p:spTree>', '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Alias"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rId6"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree>'))
    zip.file('ppt/slides/_rels/slide1.xml.rels', slideRels.replace('</Relationships>', '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>'))
    await expect(updatePowerPointChartDataPackage(await zip.generateAsync({ type:'base64' }), '8', [['3','4']])).rejects.toThrow('office_api_unsupported')
    zip.file('ppt/slides/_rels/slide1.xml.rels', slideRels.replace('</Relationships>', '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/./chart1.xml"/></Relationships>'))
    await expect(updatePowerPointChartDataPackage(await zip.generateAsync({ type:'base64' }), '8', [['3','4']])).rejects.toThrow('office_api_unsupported')
  })
})
