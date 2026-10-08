import { describe, it, expect, afterEach } from 'vitest'
import { mkdtempSync, writeFileSync, readdirSync, symlinkSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PresentationAcquisitionStore } from '../src/presentation-acquisition-store'
import { parsePresentationAcquisitionHistory } from '../src/presentation-acquisition'
const roots: string[] = []
const temporaryRoot = () => {
  const root = mkdtempSync(join(tmpdir(), 'acquisitions-'))
  roots.push(root)
  return root
}
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
describe('acquisition store', () => {
  it('recovers attempts, preserves idempotent finish and bounds history', async () => {
    const root = temporaryRoot()
    const store = new PresentationAcquisitionStore(root)
    const input = {
      kind: 'image' as const,
      source: 'https://example.com/image',
      sourceUrlHash: 'a'.repeat(64),
    }
    const first = await store.begin('doc', input)
    const result = { state: 'rejected' as const, error: 'remote_image_unavailable' as const }
    const done = await store.finish('doc', first.id, result)
    expect(
      await store.finish('doc', first.id, { error: result.error, state: result.state }),
    ).toEqual(done)
    await expect(store.finish('doc', first.id, { ...result, error: 'aborted' })).rejects.toThrow(
      'invalid_state',
    )
    for (let i = 0; i < 64; i++) await store.begin('doc', input)
    const history = await new PresentationAcquisitionStore(root).read('doc')
    expect(history.totalAttempts).toBe(65)
    expect(history.records).toHaveLength(64)
    expect(history.records[0]?.attempt).toBe(2)
    expect(() => parsePresentationAcquisitionHistory({ ...history, extra: true })).toThrow(
      'invalid_state',
    )
    expect(() => parsePresentationAcquisitionHistory({ ...history, revision: 130 })).toThrow(
      'invalid_state',
    )
  })
  it('rejects tampering, symlinks, mismatched identity and noncanonical times safely', async () => {
    const root = temporaryRoot()
    const store = new PresentationAcquisitionStore(root)
    await store.begin('doc', {
      kind: 'webpage',
      source: 'https://example.com/',
      sourceUrlHash: 'b'.repeat(64),
    })
    const h = await store.read('doc')
    const record = h.records[0]!
    expect(parsePresentationAcquisitionHistory(h)).toEqual(h)
    expect(() =>
      parsePresentationAcquisitionHistory({
        ...h,
        records: [{ ...record, source: 'https://example.com/?secret=1' }],
      }),
    ).toThrow('invalid_state')
    for (const invalid of [
      { ...h, documentId: '' },
      { ...h, records: [{ ...record, startedAt: '2026-02-30T00:00:00.000Z' }] },
      { ...h, records: [{ ...record, id: '../unsafe' }] },
    ])
      expect(() => parsePresentationAcquisitionHistory(invalid)).toThrow('invalid_state')
    const file = join(
      root,
      'presentation-acquisition-history',
      readdirSync(join(root, 'presentation-acquisition-history'))[0]!,
    )
    writeFileSync(file, 'secret failure')
    await expect(store.read('doc')).rejects.toThrow(/^invalid_state$/)
    const other = temporaryRoot()
    symlinkSync(root, join(other, 'presentation-acquisition-history'))
    await expect(new PresentationAcquisitionStore(other).read('doc')).rejects.toThrow(
      'invalid_state',
    )
  })
})
