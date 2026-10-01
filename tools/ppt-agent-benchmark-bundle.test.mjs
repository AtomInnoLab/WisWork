import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test } from 'node:test'
import { loadBenchmarkBundle } from './ppt-agent-benchmark-bundle.mjs'

test('finds each frozen source by digest across the case directory and uploads shared snapshots once', async () => {
  const root = await mkdtemp(join(tmpdir(), 'ppt-benchmark-bundle-'))
  try {
    const directory = join(root, 'docs/product/ppt-benchmark-materials/PPT-P0-05')
    await mkdir(join(directory, 'originals'), { recursive: true })
    const first = Buffer.from('%PDF-first')
    const second = Buffer.from('treatment,n\nA,2\n')
    const third = Buffer.from('{"asOf":"2024-11-01"}\n')
    const hashes = [first, second, third].map((bytes) =>
      createHash('sha256').update(bytes).digest('hex'),
    )
    await writeFile(join(directory, 'originals/source.pdf'), first)
    await writeFile(join(directory, 'data.csv'), second)
    await writeFile(join(directory, 'data-dictionary.json'), third)
    await writeFile(
      join(directory, 'reference-plan.json'),
      JSON.stringify({
        projectId: 'candidate',
        sources: [
          { snapshotAttachmentId: hashes[0] },
          { snapshotAttachmentId: hashes[0] },
          { snapshotAttachmentId: hashes[1] },
          { snapshotAttachmentId: hashes[2] },
        ],
      }),
    )
    await writeFile(
      join(directory, 'reference-deck.json'),
      JSON.stringify({
        id: 'candidate',
        slides: [{ id: 'p01' }],
      }),
    )
    const bundle = await loadBenchmarkBundle(root, 'PPT-P0-05')
    assert.deepEqual(
      bundle.sourceAttachments.map((item) => item.name),
      ['source.pdf', 'data.csv', 'data-dictionary.json'],
    )
    assert.deepEqual(
      bundle.sourceAttachments.map((item) => item.sha256),
      hashes,
    )
    await rm(join(directory, 'data.csv'))
    await assert.rejects(loadBenchmarkBundle(root, 'PPT-P0-05'), /benchmark source missing/)
    assert.equal((await readFile(join(directory, 'originals/source.pdf'))).toString(), '%PDF-first')
  } finally {
    await rm(root, { recursive: true, force: true })
  }
})

test('loads only the selected P0-17 document plan and its own frozen PDF', async () => {
  const root = await mkdtemp(join(tmpdir(), 'ppt-benchmark-bundle-'))
  try {
    const directory = join(root, 'docs/product/ppt-benchmark-materials/PPT-P0-17')
    await mkdir(directory, { recursive: true })
    for (const key of ['science', 'legal', 'finance']) {
      const bytes = Buffer.from(`%PDF-${key}`)
      const digest = createHash('sha256').update(bytes).digest('hex')
      await writeFile(join(directory, `${key}.pdf`), bytes)
      await writeFile(
        join(directory, `reference-${key}-plan.json`),
        JSON.stringify({ projectId: key, sources: [{ snapshotAttachmentId: digest }] }),
      )
      await writeFile(
        join(directory, `reference-${key}-deck.json`),
        JSON.stringify({ id: key, slides: [{ id: 'p01' }] }),
      )
    }
    const bundle = await loadBenchmarkBundle(root, 'PPT-P0-17', 'legal')
    assert.equal(bundle.plan.projectId, 'legal')
    assert.deepEqual(
      bundle.sourceAttachments.map((item) => item.name),
      ['legal.pdf'],
    )
    await assert.rejects(
      loadBenchmarkBundle(root, 'PPT-P0-17', '../science'),
      /invalid benchmark variant/,
    )
  } finally {
    await rm(root, { recursive: true, force: true })
  }
})
