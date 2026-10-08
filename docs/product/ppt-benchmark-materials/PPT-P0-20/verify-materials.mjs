import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { pdfToPages } from '../../../../packages/file-parse/src/pdf.ts'
import { checkPresentationChartData } from '@wiswork/pptx-engine/presentation-chart-data'
import {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'

const root = dirname(fileURLToPath(import.meta.url))
const sums = readFileSync(join(root, 'SHA256SUMS'), 'utf8').trim().split('\n')
if (sums.length !== 6) throw new Error('expected six frozen files')
const names = []
for (const line of sums) {
  const match = /^([a-f0-9]{64}) {2}([A-Za-z0-9.-]+)$/.exec(line)
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
  names.sort().join(',') !==
    [
      ...scenario.documentInputs,
      'scenario.json',
      'reference-plan.json',
      'reference-deck.json',
      'wiswork-generated-candidate.pptx',
    ]
      .sort()
      .join(',') ||
  scenario.pages?.length !== 8 ||
  scenario.pages.some((page, index) => page.number !== index + 1 || !page.purpose) ||
  scenario.claims?.length !== 4 ||
  scenario.claims.some((claim) => claim.review !== 'pending') ||
  scenario.faultSchedule?.map((fault) => fault.id).join(',') !== 'F1,F2,F3' ||
  !scenario.requiredEvidence?.length
)
  throw new Error('invalid scenario')
const plan = parsePresentationPlan(
  JSON.parse(readFileSync(join(root, 'reference-plan.json'), 'utf8')),
)
const deck = JSON.parse(readFileSync(join(root, 'reference-deck.json'), 'utf8'))
assertDeckMatchesPresentationPlan(deck, plan)
if (
  plan.projectId !== 'p0-20-fault-recovery-reference' ||
  plan.slides.length !== 8 ||
  plan.sources.length !== 5 ||
  deck.slides.length !== 8 ||
  plan.claims.length !== 5
)
  throw new Error('invalid source-bound candidate')
const articlePages = await pdfToPages(readFileSync(join(root, scenario.documentInputs[0])))
const checklistPages = await pdfToPages(readFileSync(join(root, scenario.documentInputs[1])))
if (articlePages.length !== 11 || checklistPages.length !== 1)
  throw new Error('source page count mismatch')
for (const source of plan.sources) {
  const supplement = source.id === 'supplement-checklist'
  const file = scenario.documentInputs[supplement ? 1 : 0]
  const bytes = readFileSync(join(root, file))
  if (source.snapshotAttachmentId !== createHash('sha256').update(bytes).digest('hex'))
    throw new Error(`source digest mismatch: ${source.id}`)
  const pages = supplement ? checklistPages : articlePages
  const page = Number(
    (supplement ? /^PDF 第 (\d+) 页$/ : /^第 (\d+) 页$/).exec(source.locator)?.[1],
  )
  if (!page || !pages[page - 1]?.includes(source.excerpt))
    throw new Error(`source excerpt mismatch: ${source.id}`)
}
if (
  plan.claims.some((claim) => claim.reviewStatus !== 'needs_review') ||
  !deck.slides.every((slide, index) => slide.title === scenario.pages[index].purpose)
)
  throw new Error('unreviewed claim or page purpose mismatch')
const chart = deck.slides[5].elements.find((element) => element.kind === 'chart')
if (
  !chart ||
  checkPresentationChartData(plan, 'p06', [chart]).charts.some((entry) => entry.findings.length)
)
  throw new Error('Table 1 source-bound chart mismatch')
const pptx = await JSZip.loadAsync(readFileSync(join(root, 'wiswork-generated-candidate.pptx')))
if (
  Object.keys(pptx.files).filter((file) => /^ppt\/slides\/slide\d+\.xml$/.test(file)).length !==
    8 ||
  Object.keys(pptx.files).filter((file) => /^ppt\/charts\/chart\d+\.xml$/.test(file)).length !== 1
)
  throw new Error('native PPTX structure mismatch')
const chartXml = await pptx.file('ppt/charts/chart1.xml')?.async('string')
for (const value of ['前测 n=14', '三个月后 n=12', '<c:v>7</c:v>', '<c:v>10</c:v>'])
  if (!chartXml?.includes(value)) throw new Error(`native chart value missing: ${value}`)
const workbook = await JSZip.loadAsync(
  await pptx.file('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')?.async('uint8array'),
)
const sheet = await workbook.file('xl/worksheets/sheet1.xml')?.async('string')
if (!sheet?.includes('<c r="B2"><v>7</v></c>') || !sheet.includes('<c r="B3"><v>10</v></c>'))
  throw new Error('embedded chart workbook mismatch')
for (let index = 0; index < 8; index++) {
  const xml = await pptx.file(`ppt/slides/slide${index + 1}.xml`)?.async('string')
  if (!xml?.includes(scenario.pages[index].purpose)) throw new Error(`slide ${index + 1} missing`)
}
console.log(
  'PPT-P0-20: six frozen files, source-bound eight-page PPTX, native chart and three fault points verified',
)
