import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'

const root = dirname(fileURLToPath(import.meta.url))
const files = [
  'apple-fy2024-form10k.pdf',
  'apple-fy2024-q4-financial-statements.pdf',
  'data-dictionary.json',
  'independent-recalc.csv',
  'p0-07-reference.pptx',
]
const lines = readFileSync(join(root, 'SHA256SUMS'), 'utf8').trim().split('\n')
assert.equal(lines.length, files.length)
for (const [index, line] of lines.entries()) {
  const match = /^([a-f0-9]{64})  ([A-Za-z0-9.-]+)$/.exec(line)
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
