import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = dirname(fileURLToPath(import.meta.url))
const lines = readFileSync(join(root, 'SHA256SUMS'), 'utf8').trim().split('\n')
if (lines.length !== 6) throw new Error('expected six frozen files')
const hashes = new Map()
for (const line of lines) {
  const match = /^([a-f0-9]{64})  ([A-Za-z0-9./-]+)$/.exec(line)
  if (!match || match[2].includes('..') || hashes.has(match[2]))
    throw new Error('invalid checksum entry')
  const actual = createHash('sha256')
    .update(readFileSync(join(root, match[2])))
    .digest('hex')
  if (actual !== match[1]) throw new Error(`checksum mismatch: ${match[2]}`)
  hashes.set(match[2], actual)
}
const scenario = JSON.parse(readFileSync(join(root, 'scenario.json'), 'utf8'))
const rights = JSON.parse(readFileSync(join(root, 'asset-rights.json'), 'utf8'))
if (
  scenario.caseId !== 'PPT-P0-16' ||
  scenario.documentInputs?.join(',') !==
    'deardorff-2020-article.pdf,deardorff-2020-checklist.pdf' ||
  scenario.imageInputs?.join(',') !== 'images/schematic-01.png,images/schematic-03.png' ||
  [...hashes.keys()].sort().join(',') !==
    [...scenario.documentInputs, ...scenario.imageInputs, 'scenario.json', 'asset-rights.json']
      .sort()
      .join(',') ||
  scenario.pages?.length !== 8 ||
  scenario.pages.some((page, index) => page.number !== index + 1 || !page.purpose) ||
  scenario.faultSchedule?.map((fault) => fault.id).join(',') !== 'F1,F2' ||
  scenario.imageInputs?.length !== 2 ||
  rights.assets?.length !== 2 ||
  rights.assets.map((asset) => asset.file).join(',') !== scenario.imageInputs.join(',') ||
  rights.assets.some((asset) => hashes.get(asset.file) !== asset.sha256 || !asset.usePermission) ||
  !scenario.requiredEvidence?.length
)
  throw new Error('invalid scenario or rights')
console.log(
  'PPT-P0-16: six frozen files, eight pages, two authorized images and two fault points verified',
)
