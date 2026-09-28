import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

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
assert.deepEqual(hashes, [`${digest}  apple-fy2024-form10k.pdf`])
console.log('PPT-P0-08 partial: Apple report and comparison arithmetic verified; Sony source PDF pending')
