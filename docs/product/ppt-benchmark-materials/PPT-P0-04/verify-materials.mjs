import { readFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import { pdfToPages } from '../../../../packages/file-parse/src/pdf.ts'
import { parsePresentationDeck } from '@wiswork/pptx-engine/presentation'
import {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'
const root = dirname(fileURLToPath(import.meta.url))
const load = (path) => readFileSync(resolve(root, path))
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')
const manifest = JSON.parse(load('materials-manifest.json'))
const basis = JSON.parse(load('basis.json'))
assert.equal(manifest.caseId, 'PPT-P0-04')
assert.equal(manifest.status, 'candidate_pending_review')
assert.equal(manifest.domain, 'law')
assert.equal(manifest.historicalAsOf, '2024-12-01')
assert.equal(manifest.effectiveFrom, '2023-12-01')
assert.equal(manifest.slideCount, 8)
assert.equal(basis.historicalAsOf, manifest.historicalAsOf)
assert.equal(manifest.sources.length, 3)
assert.equal(new Set(manifest.sources.map((s) => s.id)).size, 3)
const extracts = new Map()
for (const source of manifest.sources) {
  const bytes = load(source.path)
  assert(bytes.subarray(0, 5).equals(Buffer.from('%PDF-')))
  assert.equal(bytes.length, source.fileSize, `${source.id} original size`)
  assert.equal(sha(bytes), source.sha256, `${source.id} original SHA`)
  assert(['www.uscourts.gov', 'www.govinfo.gov'].includes(new URL(source.url).hostname))
  const pages = execFileSync('pdftotext', ['-layout', resolve(root, source.path), '-'], {
    maxBuffer: 8 * 1024 * 1024,
  })
    .toString('utf8')
    .split('\f')
  assert.equal(pages.length - 1, source.pdfPageCount)
  for (const page of source.extracts) {
    const bytes = load(page.path)
    assert.equal(bytes.length, page.fileSize)
    assert.equal(sha(bytes), page.sha256)
    assert.equal(
      bytes.toString('utf8'),
      pages[page.pdfPage - 1],
      `${source.id} actual PDF page ${page.pdfPage}`,
    )
    assert(!extracts.has(page.path))
    extracts.set(page.path, {
      text: bytes.toString('utf8'),
      sourceId: source.id,
      pdfPage: page.pdfPage,
    })
  }
}
assert.equal(new Set(basis.basis.map((b) => b.id)).size, basis.basis.length)
for (const item of basis.basis) {
  const page = extracts.get(item.extractPath)
  assert(page, `${item.id} has real scoped page`)
  assert.equal(item.sourceId, page.sourceId)
  assert.equal(item.pdfPage, page.pdfPage)
  assert.equal(item.offsetUnit, 'utf16_code_unit')
  assert(Number.isSafeInteger(item.offset) && item.offset >= 0)
  assert.equal(item.window.offset, 0)
  assert.equal(item.window.maxChars, page.text.length)
  assert.equal(
    page.text.slice(item.offset, item.offset + item.text.length),
    item.text,
    `${item.id} exact literal basis`,
  )
}
const plan = parsePresentationPlan(JSON.parse(load('reference-plan.json')))
const deck = parsePresentationDeck(JSON.parse(load('reference-deck.json')))
assertDeckMatchesPresentationPlan(deck, plan)
assert.equal(plan.domain, 'law')
assert.equal(plan.sources.length, 10)
assert.equal(plan.claims.length, 10)
assert(plan.claims.every((claim) => claim.reviewStatus === 'needs_review'))
assert.equal(plan.slides.length, 8)
assert.equal(deck.slides.length, 8)
assert(plan.slides.every((slide) => slide.domainSection))
assert(plan.claims.every((claim) => claim.professionalContext?.domain === 'law'))
assert(
  plan.claims.every(
    (claim) =>
      claim.professionalContext?.applicabilityDate ===
      (claim.id === 'old' ? '2022-12-01' : '2024-12-01'),
  ),
)
const parsedOriginals = new Map()
for (const source of plan.sources) {
  const original = manifest.sources.find((item) => item.sha256 === source.snapshotAttachmentId)
  assert(original, `unbound original: ${source.id}`)
  assert.equal(source.uri, original.url)
  const page = Number(source.locator?.match(/^第 (\d+) 页$/)?.[1])
  assert(Number.isInteger(page) && page > 0 && page <= original.pdfPageCount)
  if (!parsedOriginals.has(original.path))
    parsedOriginals.set(original.path, await pdfToPages(load(original.path)))
  const parsedPages = parsedOriginals.get(original.path)
  assert(parsedPages[page - 1].includes(source.excerpt), `PDF parser excerpt missing: ${source.id}`)
}
for (const id of ['standard', 'weight', 'scope', 'certainty', 'procedure']) {
  const claim = plan.claims.find((item) => item.id === id)
  assert(claim)
  assert.equal(claim.professionalContext.materialKind, undefined)
  assert.equal(claim.professionalContext.effectLevel, 'official explanatory committee note')
}
for (const id of ['old', 'rule-1101'])
  assert.equal(
    plan.claims.find((item) => item.id === id)?.professionalContext?.effectiveFrom,
    undefined,
  )
const pptx = await JSZip.loadAsync(load('p0-04-reference.pptx'))
assert.equal(
  Object.keys(pptx.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length,
  8,
)
const comparisonPage = await pptx.file('ppt/slides/slide4.xml')?.async('string')
assert(comparisonPage?.includes('2024') && comparisonPage.includes('旧 PDF 29'))
for (const line of load('SHA256SUMS').toString('utf8').trim().split('\n')) {
  const [digest, name] = line.split(/\s+/)
  assert.equal(sha(load(name)), digest, `hash mismatch: ${name}`)
}
console.log(
  `Candidate integrity OK: ${manifest.sources.length} official originals, ${extracts.size} exact page extracts, ${basis.basis.length} original anchors, 10 product PDF excerpts and eight native slides; legal/copyright/PowerPoint review pending.`,
)
