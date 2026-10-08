import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { checkPresentationChartData } from '@wiswork/pptx-engine/presentation-chart-data'
import { reproducePresentationCalculation } from '@wiswork/pptx-engine/presentation-calculation'
import {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'
import { pdfToPages } from '../../../../packages/file-parse/src/pdf.ts'

const root = dirname(fileURLToPath(import.meta.url))
const files = [
  'apple-fy2024-q4-financial-statements.pdf',
  'model-inputs.json',
  'scenario-results.csv',
  'reference-plan.json',
  'reference-deck.json',
  'p0-09-reference.pptx',
]
const lines = readFileSync(join(root, 'SHA256SUMS'), 'utf8').trim().split('\n')
assert.equal(lines.length, files.length)
for (const [index, line] of lines.entries()) {
  const match = /^([a-f0-9]{64}) {2}([A-Za-z0-9.-]+)$/.exec(line)
  assert.equal(match?.[2], files[index])
  assert.equal(
    createHash('sha256')
      .update(readFileSync(join(root, files[index])))
      .digest('hex'),
    match[1],
  )
}
const model = JSON.parse(readFileSync(join(root, 'model-inputs.json'), 'utf8'))
assert.equal(model.caseId, 'PPT-P0-09')
assert.equal(model.asOf, '2024-10-31')
assert.equal(model.source.auditStatus, 'unaudited')
assert.equal(model.source.page, 1)
assert.deepEqual([model.historical.netSales, model.historical.operatingIncome], [391035, 123216])
assert.equal(model.historical.currency, 'USD')
assert.equal(model.historical.unit, 'millions')
assert.equal(model.requiredSlides.length, 8)
assert.equal(model.model.scenarios.length, 3)
const pdf = join(root, model.source.file)
const pages = execFileSync('pdfinfo', [pdf], { encoding: 'utf8' })
assert.match(pages, /^Pages:\s+4\s*$/m)
const firstPage = execFileSync('pdftotext', ['-f', '1', '-l', '1', '-layout', pdf, '-'], {
  encoding: 'utf8',
})
assert.match(firstPage, /CONDENSED CONSOLIDATED STATEMENTS OF OPERATIONS \(Unaudited\)/)
assert.match(firstPage, /Twelve Months Ended[\s\S]*?September 28,[\s\S]*?2024/)
assert.match(firstPage, /Total net sales[^\n]*391,035[^\n]*383,285/)
assert.match(firstPage, /Operating income[^\n]*123,216[^\n]*114,301/)
const csv = readFileSync(join(root, 'scenario-results.csv'), 'utf8').trim().split('\n')
assert.equal(csv.length, 4)
assert.equal(
  csv[0],
  'scenario,label,sales_growth_bps,operating_margin_bps,net_sales_usd_millions,operating_income_usd_millions,evidence_type',
)
const round = (numerator) => Math.floor((numerator + 5000) / 10000)
for (const [index, scenario] of model.model.scenarios.entries()) {
  const columns = csv[index + 1].split(',')
  assert.equal(columns.length, 7)
  assert.deepEqual(columns.slice(0, 4), [
    scenario.id,
    scenario.label,
    String(scenario.salesGrowthBps),
    String(scenario.operatingMarginBps),
  ])
  const sales = round(model.historical.netSales * (10000 + scenario.salesGrowthBps))
  const income = round(sales * scenario.operatingMarginBps)
  assert.deepEqual(columns.slice(4), [String(sales), String(income), 'calculated_hypothesis'])
}
const plan = parsePresentationPlan(
  JSON.parse(readFileSync(join(root, 'reference-plan.json'), 'utf8')),
)
const deck = JSON.parse(readFileSync(join(root, 'reference-deck.json'), 'utf8'))
assertDeckMatchesPresentationPlan(deck, plan)
assert.equal(plan.domain, 'finance')
assert.equal(plan.slides.length, 8)
assert.equal(plan.claims.length, 8)
assert.equal(plan.sources.length, 6)
assert.ok(plan.claims.every((claim) => claim.reviewStatus === 'needs_review'))
assert.ok(plan.claims.every((claim) => claim.professionalContext?.domain === 'finance'))
assert.equal(plan.claims.filter((claim) => claim.type === 'calculation').length, 6)
const sourceFiles = new Map([
  [model.source.url, model.source.file],
  ['local:model-inputs.json', 'model-inputs.json'],
  ['local:scenario-results.csv', 'scenario-results.csv'],
])
const pdfText = (await pdfToPages(readFileSync(pdf)))[0]
for (const source of plan.sources) {
  const file = sourceFiles.get(source.uri)
  assert.ok(file)
  const bytes = readFileSync(join(root, file))
  assert.equal(source.snapshotAttachmentId, createHash('sha256').update(bytes).digest('hex'))
  assert.ok((file.endsWith('.pdf') ? pdfText : bytes.toString('utf8')).includes(source.excerpt))
  if (file.endsWith('.pdf')) assert.match(source.locator, /PDF 1 FY2024 12 months/)
}
const historical = plan.claims.find((claim) => claim.id === 'historical')
assert.deepEqual(historical.sourceIds, ['historical-sales', 'historical-income'])
assert.match(historical.statement, /未经审计/)
for (const scenario of model.model.scenarios) {
  for (const [kind, index] of [
    ['sales', 4],
    ['income', 5],
  ]) {
    const claim = plan.claims.find((item) => item.id === `${scenario.id}-${kind}`)
    assert.ok(claim)
    assert.equal(reproducePresentationCalculation(claim).status, 'reproduced')
    assert.ok(claim.sourceIds.includes(`${scenario.id}-csv`))
    assert.ok(plan.slides[index].claimIds.includes(claim.id))
    assert.match(claim.statement, /假设/)
  }
}
for (const index of [4, 5]) {
  const slide = deck.slides[index]
  const chart = slide.elements.find((element) => element.kind === 'chart')
  assert.ok(chart)
  const checked = checkPresentationChartData(plan, slide.id, [chart])
  assert.ok(checked.charts.every((item) => item.findings.length === 0))
}
const reference = await JSZip.loadAsync(readFileSync(join(root, 'p0-09-reference.pptx')))
assert.equal(
  Object.keys(reference.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length,
  8,
)
assert.equal(
  Object.keys(reference.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name)).length,
  2,
)
for (let index = 0; index < 8; index++) {
  const slide = await reference.file(`ppt/slides/slide${index + 1}.xml`)?.async('string')
  assert.ok(slide?.includes(`<a:t>${model.requiredSlides[index]}</a:t>`))
  if (index === 4 || index === 5) assert.ok(slide?.includes('非 Apple 指引或预测'))
}
for (const [chartIndex, columnIndex] of [
  [1, 4],
  [2, 5],
]) {
  const xml = await reference.file(`ppt/charts/chart${chartIndex}.xml`)?.async('string')
  assert.match(xml, /<c:valAx>[\s\S]*?<c:scaling>[\s\S]*?<c:min val="0"\/>/)
  const values = [...xml.matchAll(/<c:v>([^<]+)<\/c:v>/g)].map((match) => match[1])
  assert.deepEqual(
    values.slice(1, 4),
    model.model.scenarios.map((scenario) => scenario.label),
  )
  assert.deepEqual(
    values.slice(4),
    csv.slice(1).map((line) => line.split(',')[columnIndex]),
  )
}
console.log(
  'PPT-P0-09: official four-page disclosure, three scenario calculations and two native charts verified',
)
