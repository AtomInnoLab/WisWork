import { expect, it } from 'vitest'
import {
  parsePresentationManualObservation,
  parsePresentationManualObservationShape,
  presentationManualObservationDigest,
} from '../src/presentation-manual-observation.js'
import { parseSavedPresentationPreference } from '../src/presentation-preference.js'
const shape = {
  id: 'shape',
  name: 'Title',
  type: 'TextBox' as const,
  left: 1,
  top: 2,
  width: 300,
  height: 40,
  rotation: 15,
  text: 'hello',
  font: { name: null, size: 20, color: null, bold: null },
}
it('preserves bounded aggregate literal state and canonical digest without attributing authorship', async () => {
  const digest = await presentationManualObservationDigest(shape)
  expect(await presentationManualObservationDigest({ ...shape, text: 'other' })).not.toBe(digest)
  expect(await presentationManualObservationDigest({ ...shape, rotation: 45 })).not.toBe(digest)
  const input = {
    version: 1,
    source: 'host_difference_unattributed',
    observationId: 'one',
    documentId: 'doc',
    projectId: 'p',
    slideId: 'slide',
    shapeId: 'shape',
    before: { capturedAt: '2026-09-29T00:00:00.000Z', shape, digest },
    atomicSnapshot: false,
    coverage: 'text_geometry_aggregate_font',
  }
  expect(parsePresentationManualObservation(input)).toEqual(input)
  for (const change of [
    { author: 'user' },
    { atomicSnapshot: true },
    { source: 'user' },
    { shapeId: 'other' },
    { before: { ...input.before, capturedAt: '2026-02-30T00:00:00.000Z' } },
    { after: { ...input.before, capturedAt: '2026-09-28T00:00:00.000Z' } },
    { after: { ...input.before, shape: { ...shape, type: 'Chart' } } },
  ])
    expect(() => parsePresentationManualObservation({ ...input, ...change })).toThrow()
  const output = parsePresentationManualObservation(input)
  output.before.shape.text = 'changed'
  expect(input.before.shape.text).toBe('hello')
  const { rotation: _rotation, ...legacyShape } = shape
  const legacy = {
    ...input,
    before: {
      ...input.before,
      shape: legacyShape,
      digest: await presentationManualObservationDigest(legacyShape),
    },
  }
  expect(parsePresentationManualObservation(legacy)).toEqual(legacy)
})
it('rejects unsupported shapes, extra fields, invalid aggregate fonts and oversized literal records', () => {
  for (const change of [
    { type: 'Chart' },
    { width: 0 },
    { height: NaN },
    { rotation: '45' },
    { rotation: 361 },
    { author: 'user' },
    { font: { name: 'Arial', size: 12, color: null, extra: true } },
    { font: { name: null, size: null, color: null, bold: undefined } },
    { text: undefined },
  ])
    expect(() => parsePresentationManualObservationShape({ ...shape, ...change })).toThrow()
  expect(
    parsePresentationManualObservationShape({ ...shape, type: 'GeometricShape', text: 'a\nb' })
      .text,
  ).toBe('a\nb')
})
it('keeps legacy preference shape and strictly preserves optional manual provenance', () => {
  const old = { projectId: 'p', changeId: 'edit', text: 'short titles' }
  expect(parseSavedPresentationPreference(old)).toEqual(old)
  const item = {
    ...old,
    changeId: 'manual_one',
    origin: {
      version: 1,
      observationId: 'one',
      beforeDigest: 'a'.repeat(64),
      afterDigest: 'b'.repeat(64),
    },
  }
  expect(parseSavedPresentationPreference(item)).toEqual(item)
  for (const change of [
    { origin: undefined },
    { changeId: 'other' },
    { origin: { ...item.origin, author: 'user' } },
    { origin: { ...item.origin, afterDigest: 'bad' } },
  ])
    expect(() => parseSavedPresentationPreference({ ...item, ...change })).toThrow()
})
