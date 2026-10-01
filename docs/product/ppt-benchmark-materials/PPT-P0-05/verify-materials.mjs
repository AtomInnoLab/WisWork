import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { pdfToPages } from '../../../../packages/file-parse/src/pdf.ts'
import { parsePresentationDeck } from '@wiswork/pptx-engine/presentation'
import {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'

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
const checksumLines = readFileSync(new URL('SHA256SUMS', root), 'utf8').trim().split('\n')
assert.equal(checksumLines.slice(0, 2).join('\n') + '\n', sums.join(''))
for (const line of checksumLines) {
  const match = /^([a-f0-9]{64}) {2}([a-z0-9.-]+)$/.exec(line)
  assert(match, `invalid checksum line: ${line}`)
  assert.equal(
    createHash('sha256')
      .update(readFileSync(new URL(match[2], root)))
      .digest('hex'),
    match[1],
  )
}
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
const plan = parsePresentationPlan(readJson('reference-plan.json'))
const deck = parsePresentationDeck(readJson('reference-deck.json'))
assertDeckMatchesPresentationPlan(deck, plan)
assert.equal(plan.domain, 'law')
assert.equal(plan.sources.length, 12)
assert.equal(plan.claims.length, basis.statements.length)
assert.equal(plan.slides.length, 8)
assert.equal(deck.slides.length, 8)
assert.deepEqual(
  plan.slides.map((slide) => slide.claimIds),
  basis.slides.map((slide) => slide.statementIds),
)
assert(plan.slides.every((slide) => slide.domainSection))
assert(
  plan.claims.every((claim) => claim.reviewStatus === 'needs_review' && claim.asOf === basis.asOf),
)
assert.equal(plan.claims.find((claim) => claim.id === 'comparison')?.type, 'judgment')
for (const claim of plan.claims) {
  assert.equal(claim.professionalContext?.domain, 'law')
  assert.equal(claim.professionalContext?.applicabilityDate, basis.asOf)
  assert(claim.professionalContext?.limitations)
  if (claim.id !== 'comparison') {
    const original = manifest.sources.find((source) =>
      claim.id.startsWith(source.id === 'google' ? 'g-' : 'w-'),
    )
    assert.equal(claim.professionalContext?.caseNumber, original?.docket)
  }
}
const parsedPages = new Map()
for (const source of plan.sources) {
  const original = manifest.sources.find((item) => item.sha256 === source.snapshotAttachmentId)
  assert(original, `unbound original: ${source.id}`)
  assert.equal(source.uri, original.url)
  const page = Number(/^第 (\d+) 页$/.exec(source.locator)?.[1])
  assert(page >= original.majorityPdfPages[0] && page <= original.majorityPdfPages[1])
  if (!parsedPages.has(original.id))
    parsedPages.set(original.id, await pdfToPages(readFileSync(new URL(original.file, root))))
  assert(
    parsedPages.get(original.id)[page - 1].includes(source.excerpt),
    `literal excerpt missing: ${source.id}`,
  )
}
const pptx = await JSZip.loadAsync(readFileSync(new URL('p0-05-reference.pptx', root)))
assert.equal(
  Object.keys(pptx.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length,
  8,
)
assert.equal(Object.keys(pptx.files).filter((name) => /^ppt\/media\/[^/]+$/.test(name)).length, 0)
const matrix = await pptx.file('ppt/slides/slide6.xml')?.async('string')
assert(matrix?.includes('撤销并发回') && matrix.includes('维持'))
console.log(
  '8-page source-backed editable candidate OK; professional, rights and actual PowerPoint review remain pending.',
)
