import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'

const root = dirname(fileURLToPath(import.meta.url))
const basis = JSON.parse(readFileSync(join(root, 'basis.json'), 'utf8'))
const report = join(root, 'apple-fy2024-form10k.pdf')
const digest = createHash('sha256').update(readFileSync(report)).digest('hex')
assert.equal(basis.caseId, 'PPT-P0-08')
assert.equal(basis.asOf, '2024-11-01')
assert.equal(basis.companies.length, 2)
assert.equal(basis.companies[0].localFile, 'apple-fy2024-form10k.pdf')
assert.equal(basis.companies[1].localFile, null)
assert.equal(basis.companies[0].valueMillions, 391035)
assert.equal(basis.companies[1].valueMillions, 13020768)
assert.equal(basis.companies[0].currency, 'USD')
assert.equal(basis.companies[1].currency, 'JPY')
assert.equal(basis.companies[1].valueSource, basis.companies[1].source)
assert.equal(basis.companies[1].valueSourcePdfPage, 48)
assert.equal(basis.illustrativeFx.sourcePdfPage, 24)
assert.equal(basis.companies[0].accounting, 'US GAAP')
assert.equal(basis.companies[1].accounting, 'IFRS')
assert.notEqual(basis.companies[0].fiscalYearEnd, basis.companies[1].fiscalYearEnd)
assert.equal(basis.illustrativeFx.jpyPerUsd, 144.4)
assert.equal(
  Math.round((basis.companies[1].valueMillions / basis.illustrativeFx.jpyPerUsd) * 100) / 100,
  basis.illustrativeFx.sonyUsdMillionsRounded2,
)
assert.equal(basis.illustrativeFx.status, 'illustrative_only')
assert.equal(basis.comparison.likeForLikeUsdRevenueDifferenceMillions, null)
assert.equal(basis.comparison.crossCompanyRevenueRanking, null)
assert.match(basis.review, /^pending_/)
assert.match(execFileSync('pdfinfo', [report], { encoding: 'utf8' }), /^Pages:\s+121\s*$/m)
const applePage = execFileSync('pdftotext', ['-f', '32', '-l', '32', '-layout', report, '-'], {
  encoding: 'utf8',
})
assert.match(applePage, /Total net sales[^\n]*391,035/)
const hashes = readFileSync(join(root, 'SHA256SUMS'), 'utf8').trim().split('\n')
const files = ['apple-fy2024-form10k.pdf', 'basis.json', 'independent-recalc.csv', 'p0-08-reference.pptx']
assert.equal(hashes.length, files.length)
for (const [index, name] of files.entries()) {
  const fileDigest = createHash('sha256').update(readFileSync(join(root, name))).digest('hex')
  assert.equal(hashes[index], `${fileDigest}  ${name}`)
}
assert.equal(hashes[0], `${digest}  apple-fy2024-form10k.pdf`)
const rows = readFileSync(join(root, 'independent-recalc.csv'), 'utf8').trim().split('\n')
assert.equal(rows.length, 4)
assert.deepEqual(rows[1].split(','), [
  'sony_total_sales_and_financial_services_revenue',
  String(basis.companies[1].valueMillions),
  'JPY millions',
  String(basis.illustrativeFx.jpyPerUsd),
  `${basis.companies[1].valueMillions}/${basis.illustrativeFx.jpyPerUsd}`,
  String(basis.illustrativeFx.sonyUsdMillionsRounded2),
  'illustrative_only',
])
assert.deepEqual(rows[2].split(','), ['like_for_like_revenue_difference', '', '', '', '', '', 'not_comparable'])
assert.deepEqual(rows[3].split(','), ['cross_company_revenue_ranking', '', '', '', '', '', 'not_comparable'])
const reference = await JSZip.loadAsync(readFileSync(join(root, 'p0-08-reference.pptx')))
assert.equal(Object.keys(reference.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length, 8)
assert.equal(Object.keys(reference.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name)).length, 2)
for (const [index, expected] of [
  [`${basis.companies[0].metric}（百万美元）`, basis.companies[0].name, String(basis.companies[0].valueMillions)],
  [`${basis.companies[1].metric}（百万日元）`, basis.companies[1].name, String(basis.companies[1].valueMillions)],
].entries()) {
  const xml = await reference.file(`ppt/charts/chart${index + 1}.xml`)?.async('string')
  assert.match(xml, /<c:valAx>[\s\S]*?<c:scaling>[\s\S]*?<c:min val="0"\/>/)
  assert.deepEqual([...xml.matchAll(/<c:v>([^<]+)<\/c:v>/g)].map((match) => match[1]), expected)
}
const slide7 = await reference.file('ppt/slides/slide7.xml')?.async('string')
assert.ok(slide7?.includes('同口径美元收入差额：空缺'))
assert.ok(slide7?.includes('跨公司收入排名：空缺'))
const tableRows = (xml) => {
  const table = /<a:tbl>([\s\S]*?)<\/a:tbl>/.exec(xml)?.[1]
  assert.ok(table)
  return [...table.matchAll(/<a:tr\b[^>]*>([\s\S]*?)<\/a:tr>/g)].map((row) =>
    [...row[1].matchAll(/<a:tc>([\s\S]*?)<\/a:tc>/g)].map((cell) =>
      [...cell[1].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((text) => text[1]).join(''),
    ),
  )
}
const slide3 = await reference.file('ppt/slides/slide3.xml')?.async('string')
assert.ok(slide3?.includes('name="basis-matrix"'))
assert.deepEqual(tableRows(slide3), [
  ['核对项', 'Apple FY2024', 'Sony FY2023'],
  ['财年结束', basis.companies[0].fiscalYearEnd, basis.companies[1].fiscalYearEnd],
  ['会计准则', basis.companies[0].accounting, basis.companies[1].accounting],
  ['币种与单位', 'USD millions', 'JPY millions'],
  ['收入范围', basis.companies[0].metric, basis.companies[1].metric],
])
assert.ok(slide7.includes('name="comparison-gaps"'))
assert.deepEqual(tableRows(slide7), [
  ['比较项', '结果', '理由'],
  ['同口径美元收入差额', '', '财年、准则和范围不同'],
  ['跨公司收入排名', '', '换算不能消除口径差异'],
])
console.log('PPT-P0-08 partial: Apple report, recalc, eight slides, native charts and blank comparison table verified; Sony PDF pending')
