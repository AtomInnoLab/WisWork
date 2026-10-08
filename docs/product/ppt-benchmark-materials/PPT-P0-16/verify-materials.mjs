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
const lines = readFileSync(join(root, 'SHA256SUMS'), 'utf8').trim().split('\n')
if (lines.length !== 9) throw new Error('expected nine frozen files')
const hashes = new Map()
for (const line of lines) {
  const match = /^([a-f0-9]{64}) {2}([A-Za-z0-9./-]+)$/.exec(line)
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
    [
      ...scenario.documentInputs,
      ...scenario.imageInputs,
      'scenario.json',
      'asset-rights.json',
      'reference-plan.json',
      'reference-deck.json',
      'wiswork-generated-candidate.pptx',
    ]
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
const plan = parsePresentationPlan(
  JSON.parse(readFileSync(join(root, 'reference-plan.json'), 'utf8')),
)
const deck = JSON.parse(readFileSync(join(root, 'reference-deck.json'), 'utf8'))
assertDeckMatchesPresentationPlan(deck, plan)
if (plan.slides.length !== 8 || deck.slides.length !== 8 || plan.sources.length !== 6)
  throw new Error('invalid eight-page plan')
const pdfs = await Promise.all(
  scenario.documentInputs.map((file) => pdfToPages(readFileSync(join(root, file)))),
)
for (const source of plan.sources) {
  const file = source.uri.startsWith('local:')
    ? 'asset-rights.json'
    : source.uri.endsWith('.s001')
      ? scenario.documentInputs[1]
      : scenario.documentInputs[0]
  if (source.snapshotAttachmentId !== hashes.get(file))
    throw new Error(`source digest mismatch: ${source.id}`)
  const content =
    file === 'asset-rights.json'
      ? readFileSync(join(root, file), 'utf8')
      : pdfs[file === scenario.documentInputs[1] ? 1 : 0].join('\n')
  if (!content.includes(source.excerpt)) throw new Error(`source excerpt mismatch: ${source.id}`)
}
for (const [index, file] of scenario.imageInputs.entries()) {
  const asset = deck.assets[index]
  if (
    createHash('sha256').update(Buffer.from(asset.base64, 'base64')).digest('hex') !==
      hashes.get(file) ||
    asset.license !== 'owned'
  )
    throw new Error(`image asset mismatch: ${file}`)
  if (
    !deck.slides[index === 0 ? 1 : 3].elements.some(
      (element) =>
        element.kind === 'image' &&
        element.assetId === asset.id &&
        /非研究测量数据/.test(element.altText),
    )
  )
    throw new Error(`image placement mismatch: ${file}`)
}
const chart = deck.slides[5].elements.find((element) => element.kind === 'chart')
if (
  !chart ||
  checkPresentationChartData(plan, deck.slides[5].id, [chart]).charts.some(
    (item) => item.findings.length,
  )
)
  throw new Error('native chart source mismatch')
const zip = await JSZip.loadAsync(readFileSync(join(root, 'wiswork-generated-candidate.pptx')))
const files = Object.keys(zip.files)
if (
  files.filter((file) => /^ppt\/slides\/slide\d+\.xml$/.test(file)).length !== 8 ||
  files.filter((file) => /^ppt\/charts\/chart\d+\.xml$/.test(file)).length !== 1 ||
  files.filter((file) => /^ppt\/media\//.test(file)).length < 2
)
  throw new Error('compiled PPTX structure mismatch')
for (let index = 0; index < 8; index++) {
  const xml = await zip.file(`ppt/slides/slide${index + 1}.xml`)?.async('string')
  if (!xml?.includes(deck.slides[index].title)) throw new Error(`missing slide ${index + 1}`)
  if ([1, 3].includes(index) && !xml.includes('<p:pic>'))
    throw new Error(`missing image on slide ${index + 1}`)
}
console.log(
  'PPT-P0-16: nine frozen files, source-backed eight-page PPTX, two authorized images and two fault points verified',
)
