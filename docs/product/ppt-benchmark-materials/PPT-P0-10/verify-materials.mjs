import assert from 'node:assert/strict'
import { readFileSync, existsSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { join, dirname } from 'node:path'
import JSZip from 'jszip'
import {
  parsePresentationPlan,
  assertDeckMatchesPresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'
const root = dirname(fileURLToPath(import.meta.url))
const read = (n) => JSON.parse(readFileSync(join(root, n), 'utf8'))
const sha = (b) => createHash('sha256').update(b).digest('hex')
const basis = read('basis.json'),
  manifest = read('manifest.json'),
  plan = parsePresentationPlan(read('reference-plan.json')),
  deck = read('reference-deck.json')
const original = process.env.WISWORK_P0_10_PDF || basis.source.localPath
assert(
  existsSync(original),
  'Required real original missing; reacquire explicitly, never substitute a PDF',
)
const bytes = readFileSync(original)
assert.equal(bytes.length, 51870607)
assert.equal(sha(bytes), 'cadeefed4b0f0627384b6b7f3730afc729570270b8794b71759f9dc6511a36b2')
assert.match(execFileSync('pdfinfo', [original], { encoding: 'utf8' }), /Pages:\s+766/)
assert.equal(basis.scopeHumanConfirmed, false)
assert.equal(basis.review.professional, 'pending')
assert.equal(basis.slides.length, 8)
assert.equal(deck.slides.length, 8)
assert.equal(plan.slides.length, 8)
for (const slide of plan.slides) {
  assert.equal(slide.acceptanceCriteria[0], '自制中文摘要与可编辑文本')
  assert(slide.acceptanceCriteria.every((criterion) => !/表格|图表/.test(criterion)))
}
assert(deck.slides.every((slide) => slide.elements.every((element) => element.kind === 'text')))
assertDeckMatchesPresentationPlan(deck, plan)
for (const locator of basis.locators) {
  const text = execFileSync(
    'pdftotext',
    ['-f', String(locator.pdfPage), '-l', String(locator.pdfPage), '-layout', original, '-'],
    { encoding: 'utf8' },
  )
  assert(text.includes(locator.anchor), `Original locator ${locator.sourceId}`)
}
for (const source of plan.sources) assert.equal(source.uri, `attachment:${basis.source.sha256}`)
assert.equal(basis.numericBasis.conversionAuthenticatesClimateEstimate, false)
assert.equal(
  plan.claims.find((c) => c.id === 'ice-conversion').calculation.reproduction.expected,
  220 / 360,
)
const zip = await JSZip.loadAsync(readFileSync(join(root, 'p0-10-reference.pptx')))
assert.equal(Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length, 8)
assert.equal(
  Object.keys(zip.files).filter((n) => /^ppt\/media\//.test(n) && !zip.files[n].dir).length,
  0,
)
for (let i = 0; i < 8; i++) {
  const xml = await zip.file(`ppt/slides/slide${i + 1}.xml`).async('string')
  assert(xml.includes(basis.slides[i].title))
  for (const s of basis.slides[i].summary)
    assert(xml.includes(s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')))
}
for (const [name, expected] of Object.entries(manifest.files))
  assert.equal(sha(readFileSync(join(root, name))), expected, `Frozen artifact ${name}`)
console.log(
  'PPT-P0-10: original identity + original page anchors + strict plan/deck + 8 native-text slides + frozen artifact hashes verified. Candidate only; no professional/host/copyright certification.',
)
