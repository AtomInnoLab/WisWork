import { expect, it } from 'vitest'
import {
  parseSavedPresentationPreference,
  parsePresentationPreferenceSource,
  presentationPreferenceReuseId,
} from '../src/presentation-preference'
const source = { documentId: 'doc', projectId: 'source', changeId: 'edit' }
const original = { projectId: 'target', changeId: 'edit', text: 'Short titles' }
const imported = {
  ...original,
  reuse: {
    version: 1,
    source,
    sourceTextDigest: 'a'.repeat(64),
    approvedAt: '2026-09-29T00:00:00.000Z',
    approvalId: '12345678-1234-4123-8123-123456789abc',
  },
}
it('preserves old exact shape and clones complete reuse metadata', () => {
  expect(parseSavedPresentationPreference(original)).toEqual(original)
  expect(parseSavedPresentationPreference(original)).not.toHaveProperty('reuse')
  expect(parseSavedPresentationPreference(imported)).toEqual(imported)
  expect(parsePresentationPreferenceSource(source)).toEqual(source)
})
it.each([
  { ...original, reuse: undefined },
  { ...imported, reuse: { ...imported.reuse, approvedAt: '2026-02-30T00:00:00.000Z' } },
  { ...imported, reuse: { ...imported.reuse, approvalId: 'agent-approved' } },
  { ...imported, reuse: { ...imported.reuse, source: { ...source, approved: true } } },
  { ...imported, reuse: { ...imported.reuse, sourceTextDigest: 'A'.repeat(64) } },
  { ...original, approved: true },
])('rejects malformed or unrecognized preference metadata %j', (value) => {
  expect(() => parseSavedPresentationPreference(value)).toThrow('invalid_request')
})

it('computes a six-field snapshot identity independently of source key ordering', async () => {
  const first = await presentationPreferenceReuseId('target-doc', 'target', source, 'a'.repeat(64))
  expect(first).toMatch(/^reuse_[a-f0-9]{64}$/)
  expect(
    await presentationPreferenceReuseId(
      'target-doc',
      'target',
      { changeId: source.changeId, projectId: source.projectId, documentId: source.documentId },
      'a'.repeat(64),
    ),
  ).toBe(first)
  expect(
    await presentationPreferenceReuseId('target-doc', 'target', source, 'b'.repeat(64)),
  ).not.toBe(first)
})
