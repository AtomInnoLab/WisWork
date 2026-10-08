import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
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

const folder = dirname(fileURLToPath(import.meta.url))
const scenario = JSON.parse(readFileSync(join(folder, 'scenario.json'), 'utf8'))
const expected = new Map(
  readFileSync(join(folder, 'SHA256SUMS'), 'utf8')
    .trim()
    .split('\n')
    .map((line) => {
      const match = /^([a-f0-9]{64}) {2}([a-z0-9-]+\.(?:pdf|json|pptx))$/.exec(line)
      if (!match) throw new Error('invalid SHA256SUMS line')
      return [match[2], match[1]]
    }),
)
if (scenario.taskId !== 'PPT-P0-17' || scenario.documents?.length !== 3 || expected.size !== 13)
  throw new Error('invalid three-document scenario')
const markers = new Set()
const files = new Set()
const required = new Set(['scenario.json'])
for (const doc of scenario.documents) {
  if (
    typeof doc.source !== 'string' ||
    !expected.has(doc.source) ||
    files.has(doc.source) ||
    typeof doc.marker !== 'string' ||
    !/^(?:SCI|LAW|FIN)-P0-17$/.test(doc.marker) ||
    markers.has(doc.marker) ||
    !Array.isArray(doc.slideTitles) ||
    doc.slideTitles.length !== 8 ||
    doc.slideTitles.some((title) => typeof title !== 'string' || !title.trim())
  )
    throw new Error('invalid document plan')
  files.add(doc.source)
  required.add(doc.source)
  markers.add(doc.marker)
  const path = join(folder, doc.source)
  const sha = createHash('sha256').update(readFileSync(path)).digest('hex')
  if (sha !== expected.get(doc.source)) throw new Error(`source hash mismatch: ${doc.source}`)
  const info = execFileSync('pdfinfo', [path], { encoding: 'utf8' })
  if (Number(/^Pages:\s+(\d+)$/m.exec(info)?.[1]) !== doc.pages)
    throw new Error(`page count mismatch: ${doc.source}`)
  const content = execFileSync('pdftotext', [path, '-'], {
    encoding: 'utf8',
    maxBuffer: 12 * 1024 * 1024,
  })
  const phrase = {
    science: 'computational reproducibility',
    legal: 'Guidelines 07/2020',
    finance: 'Apple Inc.',
  }[doc.key]
  if (!phrase || !content.toLowerCase().includes(phrase.toLowerCase()))
    throw new Error(`source identity mismatch: ${doc.source}`)
  const prefix = `reference-${doc.key}`
  const planFile = `${prefix}-plan.json`,
    deckFile = `${prefix}-deck.json`,
    pptxFile = `${prefix}.pptx`
  for (const file of [planFile, deckFile, pptxFile]) required.add(file)
  const plan = parsePresentationPlan(JSON.parse(readFileSync(join(folder, planFile), 'utf8')))
  const deck = JSON.parse(readFileSync(join(folder, deckFile), 'utf8'))
  assertDeckMatchesPresentationPlan(deck, plan)
  if (
    plan.projectId !== `p0-17-${doc.key}` ||
    deck.slides.length !== 8 ||
    plan.sources.length !== 5 ||
    plan.sources.some(
      (source) => source.snapshotAttachmentId !== sha || source.uri !== `local:${doc.source}`,
    ) ||
    deck.slides.some(
      (slide, index) =>
        slide.title !== doc.slideTitles[index] ||
        !slide.elements.some((element) => element.text?.includes(doc.marker)),
    )
  )
    throw new Error(`cross-document plan or marker mismatch: ${doc.key}`)
  const pages = await pdfToPages(readFileSync(path))
  for (const source of plan.sources) {
    const page = Number(/^PDF 第 (\d+) 页$/.exec(source.locator)?.[1])
    if (!page || !pages[page - 1]?.includes(source.excerpt))
      throw new Error(`source excerpt mismatch: ${source.id}`)
  }
  if (doc.key === 'science' || doc.key === 'finance') {
    const chart = deck.slides[5].elements.find((element) => element.kind === 'chart')
    if (
      !chart ||
      checkPresentationChartData(plan, 'p06', [chart]).charts.some((item) => item.findings.length)
    )
      throw new Error(`${doc.key} chart source mismatch`)
    if (doc.key === 'finance') {
      const raw = plan.sources[2].excerpt
      if (
        !raw.includes('Total net sales 391,035 383,285') ||
        chart.series[0]?.values?.join(',') !== '383285,391035'
      )
        throw new Error('finance audited sales values mismatch')
    }
  }
  const zip = await JSZip.loadAsync(readFileSync(join(folder, pptxFile)))
  if (
    Object.keys(zip.files).filter((file) => /^ppt\/slides\/slide\d+\.xml$/.test(file)).length !== 8
  )
    throw new Error(`slide count mismatch: ${doc.key}`)
  for (let index = 0; index < 8; index++) {
    const slideXml = await zip.file(`ppt/slides/slide${index + 1}.xml`)?.async('string')
    if (!slideXml?.includes(doc.slideTitles[index]) || !slideXml.includes(doc.marker))
      throw new Error(`native slide mismatch: ${doc.key}/${index + 1}`)
  }
}
for (const [file, digest] of expected) {
  if (
    !required.has(file) ||
    createHash('sha256')
      .update(readFileSync(join(folder, file)))
      .digest('hex') !== digest
  )
    throw new Error(`frozen file mismatch: ${file}`)
}
if (
  scenario.interleaving?.length !== 5 ||
  files.size !== 3 ||
  markers.size !== 3 ||
  required.size !== expected.size
)
  throw new Error('invalid interleaving scenario')
process.stdout.write('PPT-P0-17: 3 isolated source plans, 24 native slides and scenario verified\n')
