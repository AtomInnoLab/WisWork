import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = dirname(fileURLToPath(import.meta.url))
const files = ['synthetic-nda.txt', 'review-policy.txt', 'revision-notes.txt', 'scenario.json']
const sums = readFileSync(join(root, 'SHA256SUMS'), 'utf8').trim().split('\n')
if (sums.length !== files.length) throw new Error('expected four frozen inputs')
for (let index = 0; index < files.length; index++) {
  const [digest, name] = sums[index].split('  ')
  if (name !== files[index] || !/^[a-f0-9]{64}$/.test(digest ?? ''))
    throw new Error('invalid frozen input list')
  const actual = createHash('sha256')
    .update(readFileSync(join(root, name)))
    .digest('hex')
  if (actual !== digest) throw new Error(`checksum mismatch: ${name}`)
}

const contract = readFileSync(join(root, files[0]), 'utf8')
const policy = readFileSync(join(root, files[1]), 'utf8')
const revisions = readFileSync(join(root, files[2]), 'utf8')
const scenarioText = readFileSync(join(root, files[3]), 'utf8')
const scenario = JSON.parse(scenarioText)
const canary = /^测试追踪码：([^\r\n]+)$/m.exec(contract)?.[1]
if (!canary || !/^PRIVATE-SIM-P0-06-[A-Z0-9]+$/.test(canary))
  throw new Error('missing synthetic disclosure probe')
if ([policy, revisions, scenarioText].some((text) => text.includes(canary)))
  throw new Error('disclosure probe copied into output guidance')

const clauseNumbers = [...contract.matchAll(/^第 (\d+) 条｜/gm)].map((match) => Number(match[1]))
if (clauseNumbers.join(',') !== '1,2,3,4,5,6,7,8')
  throw new Error('contract clause numbering changed')
const revisionNumbers = [...revisions.matchAll(/^R(\d+)｜/gm)].map((match) => Number(match[1]))
if (revisionNumbers.join(',') !== '1,2,3,4,5') throw new Error('revision numbering changed')
if (
  scenario.caseId !== 'PPT-P0-06' ||
  scenario.materialStatus !== 'candidate_synthetic_pending_legal_and_privacy_review' ||
  scenario.documentInputs?.join(',') !== files.slice(0, 3).join(',') ||
  scenario.classification !== 'internal_test_synthetic' ||
  scenario.forbiddenDisclosure?.contractFieldLabel !== '测试追踪码' ||
  scenario.forbiddenDisclosure?.requiresActualSendRecords !== true ||
  scenario.forbiddenDisclosure?.targets?.join(',') !==
    'presentation,model_request,remote_diagnostics,cross_project_asset' ||
  scenario.clauses?.map((clause) => clause.number).join(',') !== clauseNumbers.join(',') ||
  scenario.pages?.map((page) => page.number).join(',') !== clauseNumbers.join(',') ||
  scenario.pages.some((page) => !page.purpose) ||
  scenario.requiredEvidence?.length !== 6
)
  throw new Error('invalid P0-06 scenario')
const policyClauses = new Set(
  [...policy.matchAll(/第 ([0-9、]+) 条/g)].flatMap((match) => match[1].split('、').map(Number)),
)
for (const number of clauseNumbers)
  if (!policyClauses.has(number)) throw new Error(`review policy omits clause ${number}`)

console.log('PPT-P0-06: four frozen inputs, eight clauses/pages and disclosure boundary verified')
