import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { pdfToPages } from '../../../../packages/file-parse/src/pdf.ts'
import { parsePresentationDeck } from '@wiswork/pptx-engine/presentation'
import { checkPresentationChartData } from '@wiswork/pptx-engine/presentation-chart-data'
import { reproducePresentationCalculation } from '@wiswork/pptx-engine/presentation-calculation'
import {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'

const root = dirname(fileURLToPath(import.meta.url))
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')
const article = readFileSync(join(root, 'hess-peterson-2015-article.pdf'))
const dictionary = readFileSync(join(root, 'hess-peterson-2015-dictionary.pdf'))
const csv = readFileSync(join(root, 'treatment-outcomes.csv'))
const plan = parsePresentationPlan(JSON.parse(readFileSync(join(root, 'reference-plan.json'))))
const deck = parsePresentationDeck(JSON.parse(readFileSync(join(root, 'reference-deck.json'))))
assertDeckMatchesPresentationPlan(deck, plan)
assert.equal(plan.sources.length, 7)
assert.equal(plan.claims.length, 11)
assert(plan.claims.every((claim) => claim.reviewStatus === 'needs_review'))
assert.equal(deck.slides.length, 8)
const articlePages = await pdfToPages(article)
const dictionaryPages = await pdfToPages(dictionary)
assert.equal(articlePages.length, 16)
assert.equal(dictionaryPages.length, 2)
assert.match(
  execFileSync('pdfinfo', [join(root, 'hess-peterson-2015-article.pdf')], { encoding: 'utf8' }),
  /^Pages:\s+16\s*$/m,
)
assert.match(
  execFileSync('pdfinfo', [join(root, 'hess-peterson-2015-dictionary.pdf')], { encoding: 'utf8' }),
  /^Pages:\s+2\s*$/m,
)
for (const source of plan.sources) {
  const bytes = source.id.startsWith('article-')
    ? article
    : source.id === 'dictionary-fields'
      ? dictionary
      : csv
  assert.equal(source.snapshotAttachmentId, sha(bytes))
  const page = Number(source.locator?.match(/^第 (\d+) 页$/)?.[1])
  const extracted = source.id.startsWith('article-') ? articlePages : dictionaryPages
  const text = Number.isInteger(page) ? extracted[page - 1] : csv.toString('utf8')
  assert(text?.includes(source.excerpt), `source excerpt missing: ${source.id}`)
}
const lines = csv.toString('utf8').trim().split('\n')
const rows = lines.slice(1).map((line) => {
  const [treatment, n, permitted, safe] = line.split(',')
  return { treatment, n: Number(n), permitted: Number(permitted), safe: Number(safe) }
})
assert.equal(rows.length, 4)
assert.equal(
  rows.reduce((total, row) => total + row.n, 0),
  1824,
)
const chart = deck.slides[4].elements.find((element) => element.kind === 'chart')
assert(chart)
assert.equal(chart.categories.length, 4)
for (const [seriesIndex, field] of ['permitted', 'safe'].entries())
  for (const [index, row] of rows.entries()) {
    const expected = Number(((row[field] / row.n) * 100).toFixed(2))
    assert.equal(chart.series[seriesIndex].values[index], expected)
    assert(chart.categories[index].includes(`n=${row.n}`))
    const claim = plan.claims.find((item) => item.id === `${field}-${index + 1}`)
    assert(claim)
    assert.equal(reproducePresentationCalculation(claim).status, 'reproduced')
  }
const dataCheck = checkPresentationChartData(plan, 'p05', [chart])
assert(dataCheck.charts.every((item) => item.findings.length === 0))
const pptx = await JSZip.loadAsync(readFileSync(join(root, 'p0-03-reference.pptx')))
assert.equal(
  Object.keys(pptx.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length,
  8,
)
assert.equal(
  Object.keys(pptx.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name)).length,
  1,
)
const chartXml = await pptx.file('ppt/charts/chart1.xml')?.async('string')
assert(chartXml?.includes('Permitted2 同意率') && chartXml.includes('Safe2 同意率'))
const renderedSeries = chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/g)
assert.equal(renderedSeries?.length, 2)
assert.match(renderedSeries[0], /<c:spPr><a:solidFill><a:srgbClr val="087D83"\/>/)
assert.match(renderedSeries[1], /<c:spPr><a:solidFill><a:srgbClr val="173248"\/>/)
assert(chartXml.includes('formatCode="#,##0.########"'))
const workbook = await JSZip.loadAsync(
  await pptx.file('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx').async('uint8array'),
)
const sheet = await workbook.file('xl/worksheets/sheet1.xml')?.async('string')
assert(sheet)
for (const series of chart.series)
  for (const value of series.values)
    assert(chartXml.includes(`<c:v>${value}</c:v>`) && sheet.includes(`<v>${value}</v>`))
for (const line of readFileSync(join(root, 'SHA256SUMS'), 'utf8').trim().split('\n')) {
  const [expected, name] = line.split(/\s+/)
  assert.equal(sha(readFileSync(join(root, name))), expected, `hash mismatch: ${name}`)
}
console.log(
  'PPT-P0-03: article, dictionary, CSV denominators, reproduced chart values, eight native slides and frozen hashes verified; scientific/host review pending',
)
