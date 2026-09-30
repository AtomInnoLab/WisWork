import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test } from 'node:test'
import JSZip from 'jszip'
import PptxGenJS from 'pptxgenjs'
import { stagePresentationHostBundle } from './ppt-agent-stage-host-bundle.mjs'

const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex')

async function fixture() {
  const root = await mkdtemp(join(tmpdir(), 'ppt-host-stage-'))
  const deck = new PptxGenJS()
  for (let index = 0; index < 8; index++) deck.addSlide().addText(`Page ${index + 1}`)
  const pptx = Buffer.from(await deck.write({ outputType: 'nodebuffer' }))
  const files = {
    'presentation.pptx': pptx,
    'evidence.json': Buffer.from(
      JSON.stringify({ documentId: 'doc-1', projectId: 'project-1', requestId: 'request-1' }),
    ),
    'evidence.md': Buffer.from('# Evidence'),
    'claims.json': Buffer.from('[]'),
    'sources.json': Buffer.from('[]'),
    'quality.json': Buffer.from(
      JSON.stringify({ scope: 'historical_records_only', needsRecapture: true }),
    ),
    'checkpoints.json': Buffer.from('{}'),
    'README.md': Buffer.from('# Host bundle'),
  }
  const manifest = {
    version: 1,
    scope: 'current_office_document',
    documentId: 'doc-1',
    projectId: 'project-1',
    requestId: 'request-1',
    files: Object.entries(files).map(([name, bytes]) => ({
      name,
      sizeBytes: bytes.length,
      sha256: sha256(bytes),
    })),
    checks: {
      completion: 'not_verified',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
      roundTrip: 'not_run',
      hostQa: 'not_checked',
      pdf: 'not_requested',
    },
  }
  const zip = new JSZip()
  for (const [name, bytes] of Object.entries(files)) zip.file(name, bytes)
  zip.file('manifest.json', JSON.stringify(manifest))
  const bundle = join(root, 'bundle.zip')
  await writeFile(bundle, await zip.generateAsync({ type: 'nodebuffer' }))
  return { root, bundle, zip, files, manifest }
}

test('stages a verified current-host bundle without creating an acceptance pass', async () => {
  const f = await fixture()
  try {
    const output = join(f.root, 'stage')
    const draft = await stagePresentationHostBundle(f.bundle, output, 'PPT-P0-01')
    assert.equal(draft.status, 'needs_human_review')
    assert.equal(draft.structureGate, 'passed')
    assert.equal(draft.pptxSha256, sha256(f.files['presentation.pptx']))
    assert.equal(draft.missingEvidence.length, 4)
    assert.equal(draft.outcome, undefined)
    assert.deepEqual(
      await readFile(join(output, 'presentation.pptx')),
      f.files['presentation.pptx'],
    )
    assert.deepEqual(JSON.parse(await readFile(join(output, 'acceptance-draft.json'))), draft)
    await assert.rejects(
      stagePresentationHostBundle(f.bundle, output, 'PPT-P0-01'),
      /stage_output_exists/,
    )
  } finally {
    await rm(f.root, { recursive: true, force: true })
  }
})

test('rejects changed manifest bytes and mismatched task identity before creating output', async () => {
  const f = await fixture()
  try {
    f.zip.file('claims.json', '{}')
    await writeFile(f.bundle, await f.zip.generateAsync({ type: 'nodebuffer' }))
    await assert.rejects(
      stagePresentationHostBundle(f.bundle, join(f.root, 'tampered'), 'PPT-P0-01'),
      /stage_digest_mismatch/,
    )
    f.zip.file('claims.json', f.files['claims.json'])
    f.zip.file(
      'evidence.json',
      JSON.stringify({ documentId: 'other', projectId: 'project-1', requestId: 'request-1' }),
    )
    f.manifest.files.find((file) => file.name === 'evidence.json').sha256 = sha256(
      Buffer.from(await f.zip.file('evidence.json').async('nodebuffer')),
    )
    f.manifest.files.find((file) => file.name === 'evidence.json').sizeBytes = (
      await f.zip.file('evidence.json').async('nodebuffer')
    ).length
    f.zip.file('manifest.json', JSON.stringify(f.manifest))
    await writeFile(f.bundle, await f.zip.generateAsync({ type: 'nodebuffer' }))
    await assert.rejects(
      stagePresentationHostBundle(f.bundle, join(f.root, 'wrong-id'), 'PPT-P0-01'),
      /stage_evidence_invalid/,
    )
  } finally {
    await rm(f.root, { recursive: true, force: true })
  }
})

test('stages an incomplete current document as a draft with a failed eight-page gate', async () => {
  const f = await fixture()
  try {
    const short = new PptxGenJS()
    for (let index = 0; index < 7; index++) short.addSlide().addText(`Page ${index + 1}`)
    const bytes = Buffer.from(await short.write({ outputType: 'nodebuffer' }))
    f.zip.file('presentation.pptx', bytes)
    const entry = f.manifest.files.find((file) => file.name === 'presentation.pptx')
    entry.sizeBytes = bytes.length
    entry.sha256 = sha256(bytes)
    f.zip.file('manifest.json', JSON.stringify(f.manifest))
    await writeFile(f.bundle, await f.zip.generateAsync({ type: 'nodebuffer' }))
    const draft = await stagePresentationHostBundle(f.bundle, join(f.root, 'stage'), 'PPT-P0-02')
    assert.equal(draft.structureGate, 'acceptance_pptx_page_count')
    assert.equal(draft.status, 'needs_human_review')
    assert.equal(draft.outcome, undefined)
  } finally {
    await rm(f.root, { recursive: true, force: true })
  }
})
