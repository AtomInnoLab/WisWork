import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'
import { parsePresentationDeck } from '@wiswork/pptx-engine/presentation'
import { checkPresentationChartData } from '@wiswork/pptx-engine/presentation-chart-data'
import { pdfToPages } from '../../../../packages/file-parse/src/pdf.ts'

const root = dirname(fileURLToPath(import.meta.url))
const basis = JSON.parse(readFileSync(join(root, 'basis.json'), 'utf8'))
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')
assert.equal(basis.caseId, 'PPT-P0-01')
assert.match(basis.review, /^pending_/)
assert.equal(basis.sample.before, 14)
assert.equal(basis.sample.after, 12)
assert.deepEqual(basis.score, { before: 1.6, after: 2.2, maximum: 6, p: 0.318, sourcePage: 5 })
assert.deepEqual(basis.table1OpenSource, {
  before: 7,
  after: 10,
  beforeDenominator: 14,
  afterDenominator: 12,
  sourcePage: 5,
})
assert.equal(basis.limitationsPage, 9)
const article = join(root, basis.sourceFile)
const checklist = join(root, basis.checklistFile)
assert.equal(sha(readFileSync(article)), basis.sourceSha256)
assert.equal(sha(readFileSync(checklist)), basis.checklistSha256)
assert.match(execFileSync('pdfinfo', [article], { encoding: 'utf8' }), /^Pages:\s+11\s*$/m)
assert.match(execFileSync('pdfinfo', [checklist], { encoding: 'utf8' }), /^Pages:\s+1\s*$/m)
const page5 = execFileSync('pdftotext', ['-f', '5', '-l', '5', '-layout', article, '-'], {
  encoding: 'utf8',
})
for (const anchor of [
  'only 12 researchers',
  '1.6 out of 6',
  '1.6 to 2.2',
  'p = 0.318',
  'Use open source software',
])
  assert(page5.includes(anchor), `original page 5 missing: ${anchor}`)
assert.match(page5, /Use open source software\s+7\s+10/)
const page9 = execFileSync('pdftotext', ['-f', '9', '-l', '9', '-layout', article, '-'], {
  encoding: 'utf8',
})
for (const anchor of ['Limitations', 'small cohort', 'solely by the author', 'appropriate power'])
  assert(page9.includes(anchor), `original page 9 missing: ${anchor}`)
assert(
  execFileSync('pdftotext', [checklist, '-'], { encoding: 'utf8' }).includes(
    'Reproducibility Score Card',
  ),
)
const plan = parsePresentationPlan(
  JSON.parse(readFileSync(join(root, 'reference-plan.json'), 'utf8')),
)
const deck = parsePresentationDeck(
  JSON.parse(readFileSync(join(root, 'reference-deck.json'), 'utf8')),
)
assertDeckMatchesPresentationPlan(deck, plan)
assert.equal(plan.sources.length, 4)
assert(plan.sources.every((source) => source.snapshotAttachmentId === basis.sourceSha256))
const extractedPages = await pdfToPages(readFileSync(article))
for (const source of plan.sources) {
  const page = Number(source.locator?.match(/^第 (\d+) 页$/)?.[1])
  assert(Number.isInteger(page), `source ${source.id} has no PDF page locator`)
  assert(
    extractedPages[page - 1]?.includes(source.excerpt),
    `source ${source.id} excerpt is absent from PDF parser output`,
  )
}
assert.equal(plan.claims.length, 4)
assert(plan.claims.every((claim) => claim.reviewStatus === 'needs_review'))
assert.equal(plan.slides.length, 8)
assert.equal(deck.slides.length, 8)
const chartData = checkPresentationChartData(plan, 'p06', [
  deck.slides[5].elements.find((el) => el.kind === 'chart'),
])
assert(chartData.charts.every((chart) => chart.findings.length === 0))

const pptx = await JSZip.loadAsync(readFileSync(join(root, 'p0-01-reference.pptx')))
assert.equal(
  Object.keys(pptx.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length,
  8,
)
assert.equal(
  Object.keys(pptx.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name)).length,
  1,
)
const chart = await pptx.file('ppt/charts/chart1.xml')?.async('string')
assert(chart)
for (const value of ['前测 n=14', '三个月后 n=12', '<c:v>7</c:v>', '<c:v>10</c:v>'])
  assert(chart.includes(value), `native chart missing: ${value}`)
const workbook = await JSZip.loadAsync(
  await pptx.file('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')?.async('uint8array'),
)
const sheet = await workbook.file('xl/worksheets/sheet1.xml')?.async('string')
assert(sheet)
assert.match(sheet, /<c r="B2"><v>7<\/v><\/c>/)
assert.match(sheet, /<c r="B3"><v>10<\/v><\/c>/)
const slide5 = await pptx.file('ppt/slides/slide5.xml')?.async('string')
const slide6 = await pptx.file('ppt/slides/slide6.xml')?.async('string')
assert(slide5?.includes('p = 0.318'))
assert(slide6?.includes('7/14') && slide6.includes('10/12'))
for (let i = 1; i <= 8; i++) {
  const slide = await pptx.file(`ppt/slides/slide${i}.xml`)?.async('string')
  assert(slide?.includes('候选') || slide?.includes('来源：') || slide?.includes('原件已冻结'))
  assert(!slide.includes('attachment:'), `slide ${i}: raw attachment ID leaked into citation`)
}
const manifest = readFileSync(join(root, 'SHA256SUMS'), 'utf8').trim().split('\n')
for (const name of [
  basis.sourceFile,
  basis.checklistFile,
  'basis.json',
  'reference-plan.json',
  'reference-deck.json',
  'p0-01-reference.pptx',
])
  assert(
    manifest.includes(`${sha(readFileSync(join(root, name)))}  ${name}`),
    `hash missing: ${name}`,
  )
console.log(
  'PPT-P0-01: original anchors, eight native pages, chart cache/workbook and frozen hashes verified; scientific/host review pending',
)
