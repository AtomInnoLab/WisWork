import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = dirname(fileURLToPath(import.meta.url))
const sums = readFileSync(join(root, 'SHA256SUMS'), 'utf8').trim().split('\n')
if (sums.length !== 3) throw new Error('expected exactly three frozen inputs')
const names = []
for (const line of sums) {
  const match = /^([a-f0-9]{64})  ([A-Za-z0-9.-]+)$/.exec(line)
  if (!match) throw new Error('invalid SHA256SUMS entry')
  const actual = createHash('sha256')
    .update(readFileSync(join(root, match[2])))
    .digest('hex')
  if (actual !== match[1]) throw new Error(`checksum mismatch: ${match[2]}`)
  names.push(match[2])
}
const scenario = JSON.parse(readFileSync(join(root, 'scenario.json'), 'utf8'))
if (
  scenario.caseId !== 'PPT-P0-20' ||
  scenario.documentInputs?.join(',') !==
    'deardorff-2020-article.pdf,deardorff-2020-checklist.pdf' ||
  names.sort().join(',') !== [...scenario.documentInputs, 'scenario.json'].sort().join(',') ||
  scenario.pages?.length !== 8 ||
  scenario.pages.some((page, index) => page.number !== index + 1 || !page.purpose) ||
  scenario.claims?.length !== 4 ||
  scenario.claims.some((claim) => claim.review !== 'pending') ||
  scenario.faultSchedule?.map((fault) => fault.id).join(',') !== 'F1,F2,F3' ||
  !scenario.requiredEvidence?.length
)
  throw new Error('invalid scenario')
console.log(
  'PPT-P0-20: three frozen inputs, eight pages, four claims and three fault points verified',
)
