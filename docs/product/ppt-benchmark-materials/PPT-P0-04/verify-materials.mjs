import { readFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import assert from 'node:assert/strict'
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
console.log(
  `Candidate integrity OK: ${manifest.sources.length} official originals, ${extracts.size} exact PDF page extracts, ${basis.basis.length} UTF16 basis anchors; legal/copyright/layout/PowerPoint review pending.`,
)
