import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { parsePresentationDeck } from '@wiswork/pptx-engine/presentation'
import {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'

const root = dirname(fileURLToPath(import.meta.url))
const files = ['synthetic-nda.txt', 'review-policy.txt', 'revision-notes.txt', 'scenario.json']
const sums = readFileSync(join(root, 'SHA256SUMS'), 'utf8').trim().split('\n')
const outputFiles = ['reference-plan.json', 'reference-deck.json', 'p0-06-reference.pptx']
if (sums.length !== files.length + outputFiles.length)
  throw new Error('expected four inputs and three outputs')
for (let index = 0; index < files.length + outputFiles.length; index++) {
  const [digest, name] = sums[index].split('  ')
  if (name !== [...files, ...outputFiles][index] || !/^[a-f0-9]{64}$/.test(digest ?? ''))
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

const planText = readFileSync(join(root, outputFiles[0]), 'utf8')
const deckText = readFileSync(join(root, outputFiles[1]), 'utf8')
if (planText.includes(canary) || deckText.includes(canary))
  throw new Error('disclosure probe copied into plan or deck')
const plan = parsePresentationPlan(JSON.parse(planText))
const deck = parsePresentationDeck(JSON.parse(deckText))
assertDeckMatchesPresentationPlan(deck, plan)
if (plan.domain !== 'law' || plan.slides.length !== 8 || deck.slides.length !== 8)
  throw new Error('invalid eight-page law candidate')
if (plan.sources.length !== 16 || plan.claims.length !== 10)
  throw new Error('source or claim map changed')
if (plan.slides.some((slide) => !slide.domainSection))
  throw new Error('legal section map incomplete')
if (
  plan.claims.some(
    (claim) =>
      claim.reviewStatus !== 'needs_review' ||
      claim.professionalContext?.materialKind !== 'contract' ||
      claim.professionalContext?.jurisdiction !== 'unspecified in unsigned synthetic draft' ||
      claim.professionalContext?.applicabilityDate !== '2026-09-28',
  )
)
  throw new Error('professional context changed')
for (const source of plan.sources) {
  const file = files.slice(0, 3).find((name) => source.title.startsWith(`${name}：`))
  if (!file) throw new Error(`unknown source: ${source.id}`)
  const digest = createHash('sha256')
    .update(readFileSync(join(root, file)))
    .digest('hex')
  if (source.snapshotAttachmentId !== digest || source.uri !== `attachment:${digest}`)
    throw new Error(`source snapshot mismatch: ${source.id}`)
  if (
    !readFileSync(join(root, file), 'utf8').includes(source.excerpt) ||
    source.excerpt.includes(canary)
  )
    throw new Error(`source excerpt mismatch: ${source.id}`)
}
for (const number of clauseNumbers)
  if (
    !plan.sources.some(
      (source) => source.id === `c${number}` && source.locator === `合同第 ${number} 条`,
    )
  )
    throw new Error(`clause ${number} missing from plan`)
const pptx = await JSZip.loadAsync(readFileSync(join(root, outputFiles[2])))
const slidePaths = Object.keys(pptx.files).filter((name) =>
  /^ppt\/slides\/slide\d+\.xml$/.test(name),
)
if (
  slidePaths.length !== 8 ||
  Object.keys(pptx.files).some((name) => /^ppt\/media\/[^/]+$/.test(name))
)
  throw new Error('invalid editable text-only PPTX')
for (const [name, entry] of Object.entries(pptx.files)) {
  if (entry.dir) continue
  const bytes = await entry.async('nodebuffer')
  if (bytes.includes(Buffer.from(canary))) throw new Error(`disclosure probe in PPTX: ${name}`)
}
const slide3 = await pptx.file('ppt/slides/slide3.xml')?.async('string')
const slide5 = await pptx.file('ppt/slides/slide5.xml')?.async('string')
const slide6 = await pptx.file('ppt/slides/slide6.xml')?.async('string')
if (!slide3?.includes('四类例外') || !slide5?.includes('法律允许') || !slide6?.includes('强制保存'))
  throw new Error('mandatory limitations missing from PPTX')

console.log(
  'PPT-P0-06: four frozen inputs, eight clauses/pages, source-bound editable candidate and disclosure boundary verified',
)
