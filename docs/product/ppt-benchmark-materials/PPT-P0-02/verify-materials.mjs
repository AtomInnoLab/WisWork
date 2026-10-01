import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { pdfToPages } from '../../../../packages/file-parse/src/pdf.ts'
import { parsePresentationDeck } from '@wiswork/pptx-engine/presentation'
import {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'

const root = dirname(fileURLToPath(import.meta.url))
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex')
const entries = [
  ['helps', 'helps-2014-white-noise.pdf', 10],
  ['han', 'han-2013-speech-noise.pdf', 7],
  ['mohanathasan', 'mohanathasan-2025-conversation-noise.pdf', 22],
]
const plan = parsePresentationPlan(JSON.parse(readFileSync(join(root, 'reference-plan.json'))))
const deck = parsePresentationDeck(JSON.parse(readFileSync(join(root, 'reference-deck.json'))))
assertDeckMatchesPresentationPlan(deck, plan)
assert.equal(plan.sources.length, 3)
assert.equal(plan.claims.length, 3)
assert(plan.claims.every((claim) => claim.reviewStatus === 'needs_review'))
assert.equal(plan.slides.length, 8)
assert.equal(deck.slides.length, 8)
assert.deepEqual(plan.slides[5].claimIds, ['helps', 'han', 'mohanathasan'])
assert.deepEqual(plan.slides[6].claimIds, ['helps', 'han', 'mohanathasan'])
const comparison = deck.slides[5].elements.map((element) => element.text ?? '').join(' ')
for (const direction of ['有利', '未见显著差异', '不利'])
  assert(comparison.includes(direction), `comparison missing: ${direction}`)
const matrix = deck.slides[6].elements.map((element) => element.text ?? '').join(' ')
for (const dimension of ['人群', '噪声', '任务', '指标'])
  assert(matrix.includes(dimension), `comparison matrix missing: ${dimension}`)
for (const [id, name, pageCount] of entries) {
  const bytes = readFileSync(join(root, name))
  const source = plan.sources.find((item) => item.id === id)
  assert(source)
  assert.equal(source.snapshotAttachmentId, digest(bytes))
  assert.equal(source.locator, '第 1 页')
  assert.match(
    execFileSync('pdfinfo', [join(root, name)], { encoding: 'utf8' }),
    new RegExp(`^Pages:\\s+${pageCount}\\s*$`, 'm'),
  )
  const pages = await pdfToPages(bytes)
  assert.equal(pages.length, pageCount)
  assert(pages[0].includes(source.excerpt), `${id} excerpt absent from source PDF`)
}
const pptx = await JSZip.loadAsync(readFileSync(join(root, 'p0-02-reference.pptx')))
assert.equal(
  Object.keys(pptx.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length,
  8,
)
const compareXml = await pptx.file('ppt/slides/slide6.xml')?.async('string')
const matrixXml = await pptx.file('ppt/slides/slide7.xml')?.async('string')
assert(compareXml?.includes('有利') && compareXml.includes('不利'))
assert(matrixXml?.includes('人群') && matrixXml.includes('指标'))
for (const line of readFileSync(join(root, 'SHA256SUMS'), 'utf8').trim().split('\n')) {
  const [expected, name] = line.split(/\s+/)
  assert.equal(digest(readFileSync(join(root, name))), expected, `hash mismatch: ${name}`)
}
console.log(
  'PPT-P0-02: three original PDFs, exact excerpts, eight native pages and comparison matrix verified; scientific/host review pending',
)
