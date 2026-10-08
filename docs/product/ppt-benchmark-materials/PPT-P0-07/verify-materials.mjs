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
const files = [
  'apple-fy2024-form10k.pdf',
  'apple-fy2024-q4-financial-statements.pdf',
  'data-dictionary.json',
  'independent-recalc.csv',
  'p0-07-reference.pptx',
  'reference-plan.json',
  'reference-deck.json',
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
const dictionary = JSON.parse(readFileSync(join(root, 'data-dictionary.json'), 'utf8'))
assert.equal(dictionary.caseId, 'PPT-P0-07')
assert.equal(dictionary.asOf, '2024-11-01')
assert.equal(dictionary.currency, 'USD')
assert.equal(dictionary.unit, 'millions')
assert.match(dictionary.restatementReview, /^pending_financial_reviewer/)
assert.equal(dictionary.annualReport.auditedStatements, true)
assert.equal(dictionary.earningsRelease.unaudited, true)
assert.equal(dictionary.metrics.length, 3)
assert.equal(dictionary.requiredSlides.length, 8)
const annual = join(root, dictionary.annualReport.file)
const release = join(root, dictionary.earningsRelease.file)
assert.match(execFileSync('pdfinfo', [annual], { encoding: 'utf8' }), /^Pages:\s+121\s*$/m)
assert.match(execFileSync('pdfinfo', [release], { encoding: 'utf8' }), /^Pages:\s+4\s*$/m)
const page = (path, number) =>
  execFileSync('pdftotext', ['-f', String(number), '-l', String(number), '-layout', path, '-'], {
    encoding: 'utf8',
  })
const audit = page(annual, dictionary.annualReport.auditOpinionPdfPage)
assert.match(audit, /Opinion on the Financial Statements/)
assert.match(audit, /We have audited the accompanying consolidated balance sheets/)
assert.match(page(release, 1), /STATEMENTS OF OPERATIONS \(Unaudited\)/)
const annualPages = new Map()
const releasePages = new Map()
for (const metric of dictionary.metrics) {
  if (!annualPages.has(metric.annualPdfPage))
    annualPages.set(metric.annualPdfPage, page(annual, metric.annualPdfPage))
  if (!releasePages.has(metric.releasePdfPage))
    releasePages.set(metric.releasePdfPage, page(release, metric.releasePdfPage))
  const row = new RegExp(
    `${metric.annualRow}[^\\n]*${metric.fy2024.toLocaleString('en-US')}[^\\n]*${metric.fy2023.toLocaleString('en-US')}`,
  )
  assert.match(annualPages.get(metric.annualPdfPage), row)
  assert.match(releasePages.get(metric.releasePdfPage), row)
}
const csv = readFileSync(join(root, 'independent-recalc.csv'), 'utf8').trim().split('\n')
assert.equal(csv.length, 4)
assert.equal(
  csv[0],
  'metric,label,fy2024_usd_millions,fy2023_usd_millions,yoy_percent,evidence_type',
)
for (const [index, metric] of dictionary.metrics.entries()) {
  const sign = Math.sign(metric.fy2024 - metric.fy2023)
  const basisPoints =
    sign *
    Math.floor(
      (Math.abs(metric.fy2024 - metric.fy2023) * 10000 + metric.fy2023 / 2) / metric.fy2023,
    )
  assert.deepEqual(csv[index + 1].split(','), [
    metric.id,
    metric.label,
    String(metric.fy2024),
    String(metric.fy2023),
    (basisPoints / 100).toFixed(2),
    'independently_calculated',
  ])
}
const plan = parsePresentationPlan(
  JSON.parse(readFileSync(join(root, 'reference-plan.json'), 'utf8')),
)
const deck = parsePresentationDeck(
  JSON.parse(readFileSync(join(root, 'reference-deck.json'), 'utf8')),
)
assertDeckMatchesPresentationPlan(deck, plan)
assert.equal(plan.domain, 'finance')
assert.equal(plan.sources.length, 19)
assert.equal(plan.claims.length, 12)
assert.equal(plan.slides.length, 8)
assert.equal(deck.slides.length, 8)
assert(plan.slides.every((slide) => slide.domainSection))
assert(
  plan.claims.every(
    (claim) => claim.reviewStatus === 'needs_review' && claim.asOf === dictionary.asOf,
  ),
)
assert(plan.claims.every((claim) => claim.professionalContext?.domain === 'finance'))
assert(
  plan.claims.every(
    (claim) =>
      claim.professionalContext?.currency === 'USD' &&
      claim.professionalContext?.unit === 'USD millions',
  ),
)
assert.equal(
  plan.claims.find((claim) => claim.id === 'release-status')?.professionalContext?.materialKind,
  'disclosure',
)
const sourceFiles = new Map([
  [dictionary.annualReport.sourceUrl, dictionary.annualReport.file],
  [dictionary.earningsRelease.sourceUrl, dictionary.earningsRelease.file],
  ['local:data-dictionary.json', 'data-dictionary.json'],
  ['local:independent-recalc.csv', 'independent-recalc.csv'],
])
const parsedPages = new Map()
for (const source of plan.sources) {
  const file = sourceFiles.get(source.uri)
  assert(file, `unknown plan source: ${source.id}`)
  const bytes = readFileSync(join(root, file))
  assert.equal(source.snapshotAttachmentId, createHash('sha256').update(bytes).digest('hex'))
  if (file.endsWith('.pdf')) {
    if (!parsedPages.has(file)) parsedPages.set(file, await pdfToPages(bytes))
    const pageNumber = Number(/^PDF (?:第 )?(\d+)/.exec(source.locator)?.[1])
    assert(Number.isInteger(pageNumber) && pageNumber > 0)
    assert(
      parsedPages.get(file)[pageNumber - 1].includes(source.excerpt),
      `PDF excerpt: ${source.id}`,
    )
  } else {
    assert(bytes.toString('utf8').includes(source.excerpt), `literal source excerpt: ${source.id}`)
  }
}
for (const metric of dictionary.metrics) {
  const csvRow = csv.find((line) => line.startsWith(`${metric.id},`))?.split(',')
  assert(csvRow)
  for (const year of [2023, 2024]) {
    const source = plan.sources.find((item) => item.id === `${metric.id}-${year}-csv`)
    assert.equal(source?.excerpt, csvRow[year === 2024 ? 2 : 3])
    assert.equal(Number(source.excerpt), metric[`fy${year}`])
  }
  const claim = plan.claims.find((item) => item.id === `${metric.id}-yoy`)
  assert(claim)
  assert.equal(reproducePresentationCalculation(claim).status, 'reproduced')
}
for (let index = 2; index <= 4; index++) {
  const chart = deck.slides[index].elements.find((element) => element.kind === 'chart')
  assert(chart)
  const report = checkPresentationChartData(plan, plan.slides[index].id, [chart])
  assert(report.charts.every((item) => item.findings.length === 0))
}
const reference = await JSZip.loadAsync(readFileSync(join(root, 'p0-07-reference.pptx')))
assert.equal(
  Object.keys(reference.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length,
  8,
)
assert.equal(
  Object.keys(reference.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name)).length,
  3,
)
for (let index = 0; index < 8; index++) {
  const slide = await reference.file(`ppt/slides/slide${index + 1}.xml`)?.async('string')
  assert.ok(slide?.includes(`<a:t>${dictionary.requiredSlides[index]}</a:t>`))
}
for (const [index, metric] of dictionary.metrics.entries()) {
  const xml = await reference.file(`ppt/charts/chart${index + 1}.xml`)?.async('string')
  assert.match(xml, /<c:valAx>[\s\S]*?<c:scaling>[\s\S]*?<c:min val="0"\/>/)
  const values = [...xml.matchAll(/<c:v>([^<]+)<\/c:v>/g)].map((match) => match[1])
  assert.deepEqual(values, [
    `${metric.label}（百万美元）`,
    '2023 财年',
    '2024 财年',
    String(metric.fy2023),
    String(metric.fy2024),
  ])
}
const profitPage = await reference.file('ppt/slides/slide4.xml')?.async('string')
assert.ok(profitPage?.includes('同比 -3.36%，为下降'))
console.log(
  'PPT-P0-07: audited 10-K, release, three YoY calculations and three native charts verified',
)
