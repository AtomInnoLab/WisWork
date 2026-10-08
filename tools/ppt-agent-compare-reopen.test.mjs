import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test } from 'node:test'
import JSZip from 'jszip'
import PptxGenJS from 'pptxgenjs'
import { compareStagedReopen } from './ppt-agent-compare-reopen.mjs'

const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex')

async function staged(root, name, bytes, time) {
  const directory = join(root, name)
  await mkdir(directory)
  await writeFile(join(directory, 'presentation.pptx'), bytes)
  const draft = {
    version: 1,
    status: 'needs_human_review',
    caseId: 'PPT-P0-01',
    documentId: 'doc-1',
    projectId: 'project-1',
    requestId: 'request-1',
    planRevision: 1,
    inputDigest: sha256('input'),
    planDigest: sha256('plan'),
    sourceBundleSha256: sha256(name),
    bundleCreatedAt: time,
    pptxFile: 'presentation.pptx',
    pptxSha256: sha256(bytes),
    structureGate: 'passed',
  }
  await writeFile(join(directory, 'acceptance-draft.json'), JSON.stringify(draft))
  return { directory, draft }
}

test('matches OOXML part bytes across different ZIP metadata without claiming a real reopen', async () => {
  const root = await mkdtemp(join(tmpdir(), 'ppt-reopen-'))
  try {
    const pptx = new PptxGenJS()
    for (let index = 0; index < 8; index++) pptx.addSlide().addText(`Page ${index + 1}`)
    const original = Buffer.from(await pptx.write({ outputType: 'nodebuffer' }))
    const rezip = await JSZip.loadAsync(original)
    rezip.file('ppt/slides/slide1.xml').date = new Date('2001-01-01T00:00:00.000Z')
    const rezipped = await rezip.generateAsync({
      type: 'nodebuffer',
      compression: 'STORE',
    })
    assert.notEqual(sha256(original), sha256(rezipped))
    const before = await staged(root, 'before', original, '2026-09-30T00:00:00.000Z')
    const after = await staged(root, 'after', rezipped, '2026-09-30T00:10:00.000Z')
    const report = await compareStagedReopen(before.directory, after.directory)
    assert.equal(report.packageComparison, 'exact_part_bytes')
    assert.deepEqual(report.changedParts, [])
    assert.equal(report.hostReopenVerified, false)
    assert.match(report.requiredHumanEvidence, /现场录屏/)
  } finally {
    await rm(root, { recursive: true, force: true })
  }
})

test('reports a changed slide and refuses mismatched or tampered staged exports', async () => {
  const root = await mkdtemp(join(tmpdir(), 'ppt-reopen-'))
  try {
    const pptx = new PptxGenJS()
    for (let index = 0; index < 8; index++) pptx.addSlide().addText(`Page ${index + 1}`)
    const original = Buffer.from(await pptx.write({ outputType: 'nodebuffer' }))
    const zip = await JSZip.loadAsync(original)
    const path = 'ppt/slides/slide2.xml'
    const xml = await zip.file(path).async('string')
    zip.file(path, xml.replace('Page 2', 'Changed 2'))
    const changed = await zip.generateAsync({ type: 'nodebuffer' })
    const before = await staged(root, 'before', original, '2026-09-30T00:00:00.000Z')
    const after = await staged(root, 'after', changed, '2026-09-30T00:10:00.000Z')
    const report = await compareStagedReopen(before.directory, after.directory)
    assert.equal(report.packageComparison, 'changed_requires_review')
    assert.deepEqual(report.changedParts, [path])
    after.draft.documentId = 'other'
    await writeFile(join(after.directory, 'acceptance-draft.json'), JSON.stringify(after.draft))
    await assert.rejects(
      compareStagedReopen(before.directory, after.directory),
      /reopen_identity_invalid/,
    )
    after.draft.documentId = 'doc-1'
    await writeFile(join(after.directory, 'acceptance-draft.json'), JSON.stringify(after.draft))
    await writeFile(join(after.directory, 'presentation.pptx'), 'changed')
    await assert.rejects(
      compareStagedReopen(before.directory, after.directory),
      /reopen_pptx_invalid/,
    )
    assert.notEqual(
      sha256(await readFile(join(after.directory, 'presentation.pptx'))),
      after.draft.pptxSha256,
    )
  } finally {
    await rm(root, { recursive: true, force: true })
  }
})
