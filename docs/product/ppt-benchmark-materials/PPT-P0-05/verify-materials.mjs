import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'

const root = new URL('./', import.meta.url)
const readJson = (name) => JSON.parse(readFileSync(new URL(name, root), 'utf8'))
const manifest = readJson('manifest.json')
const basis = readJson('basis.json')
assert.equal(manifest.caseId, 'PPT-P0-05')
assert.equal(basis.caseId, manifest.caseId)
assert.equal(basis.asOf, '2024-12-01')
assert.equal(manifest.analyticalAsOf, basis.asOf)
assert.equal(manifest.captureIsNotHistoricalSnapshot, true)
assert.equal(manifest.status, 'candidate_pending_review_and_actual_powerpoint')
assert.equal(manifest.sources.length, 2)
assert.equal(new Set(manifest.sources.map((source) => source.docket)).size, 2)
const texts = new Map()
const sums = []
for (const source of manifest.sources) {
  assert.match(source.file, /^[a-z-]+\.pdf$/)
  assert.equal(new URL(source.url).hostname, 'www.supremecourt.gov')
  assert.equal(new URL(source.docketUrl).hostname, 'www.supremecourt.gov')
  assert.ok(source.decisionDate <= basis.asOf && source.judgmentIssuedDate <= basis.asOf)
  assert.equal(source.artifactStatus, 'official_slip_opinion_subject_to_formal_revision')
  const path = fileURLToPath(new URL(source.file, root))
  const bytes = readFileSync(path)
  assert.equal(bytes.subarray(0, 5).toString(), '%PDF-')
  assert.equal(bytes.length, source.bytes)
  const hash = createHash('sha256').update(bytes).digest('hex')
  assert.equal(hash, source.sha256)
  sums.push(`${hash}  ${source.file}\n`)
  const info = execFileSync('pdfinfo', [path], { encoding: 'utf8' })
  assert.equal(Number(/^Pages:\s+(\d+)/m.exec(info)?.[1]), source.pdfPages)
  const text = execFileSync('pdftotext', ['-layout', path, '-'], {
    encoding: 'utf8',
    maxBuffer: 8 * 1024 * 1024,
  })
  assert.ok(text.includes(source.docket.replace('-', '–')) || text.includes(source.docket))
  texts.set(source.id, text.split('\f'))
  console.log(`${source.id}: ${source.bytes} bytes, ${source.pdfPages} pages, SHA-256 OK`)
}
assert.equal(readFileSync(new URL('SHA256SUMS', root), 'utf8'), sums.join(''))
assert.equal(basis.slides.length, 8)
assert.deepEqual(
  basis.slides.map((slide) => slide.page),
  [1, 2, 3, 4, 5, 6, 7, 8],
)
const statements = new Map(basis.statements.map((statement) => [statement.id, statement]))
assert.equal(statements.size, basis.statements.length)
for (const slide of basis.slides) for (const id of slide.statementIds) assert.ok(statements.has(id))
for (const statement of basis.statements) {
  assert.ok(
    [
      'reported_facts',
      'issue',
      'court_judgment',
      'express_limitation',
      'research_inference_not_holding',
    ].includes(statement.kind),
  )
  assert.ok(statement.refs.length)
  for (const ref of statement.refs) {
    const source = manifest.sources.find((item) => item.id === ref.sourceId)
    assert.ok(source)
    const [from, to] = ref.pdfPages
    assert.ok(Number.isInteger(from) && Number.isInteger(to) && from <= to)
    assert.ok(from >= source.majorityPdfPages[0] && to <= source.majorityPdfPages[1])
    assert.deepEqual(ref.printedPages, [
      from - source.majorityPdfPages[0] + 1,
      to - source.majorityPdfPages[0] + 1,
    ])
    for (let page = from; page <= to; page++) {
      const text = texts.get(ref.sourceId)[page - 1]
      assert.ok(text.trim().length > 100)
      assert.match(text, /Opinion of the Court/)
    }
    for (const anchor of ref.anchors ?? []) {
      assert.ok(anchor.pdfPage >= from && anchor.pdfPage <= to)
      assert.ok(
        texts.get(ref.sourceId)[anchor.pdfPage - 1].replace(/\s+/g, ' ').includes(anchor.text),
      )
    }
  }
}
const googleQuestion = statements.get('g-question').refs[0]
assert.deepEqual(googleQuestion.pdfPages, [18, 19])
assert.match(googleQuestion.section, /^III B:/)
assert.equal(googleQuestion.anchors.length, 2)
assert.match(texts.get('google')[17], /^\s*B\s*$/m)
assert.match(texts.get('google')[18], /^\s*IV\s*$/m)
assert.match(texts.get('warhol')[26], /expresses no opinion/)
assert.ok(basis.prohibitedGeneralizations.length >= 5)
assert.ok(
  manifest.pending.includes('legal_review') && manifest.pending.includes('actual_powerpoint_task'),
)
console.log(
  '8-page scope and source positions OK; professional, rights and actual PowerPoint review remain pending.',
)
