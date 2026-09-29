import { expect, it } from 'vitest'
import {
  presentationImportContent,
  parsePresentationImportSource,
  type PresentationImportSource,
} from '../src/presentation-import-source.js'
it('preserves the original full-file and ordered production receipt digest inputs', () => {
  const artifact = {
    documentId: 'doc',
    projectId: 'p',
    requestId: 'r',
    planRevision: 7,
    pptxBase64: 'UEs=',
    pages: [{ id: 'page', title: '标题', sourceSlideId: '256#' }],
    pagePptxBase64: ['UEs='],
  }
  expect(presentationImportContent(artifact)).toBe(
    '{"documentId":"doc","projectId":"p","requestId":"r","planRevision":7,"pages":[{"id":"page","title":"标题","sourceSlideId":"256#"}],"pagePptxBase64":["UEs="]}',
  )
  const { pagePptxBase64: _pages, ...whole } = artifact
  expect(presentationImportContent(whole)).toBe('UEs=')
})
it('validates bounded namespace-specific descriptors without binary bodies and returns detached data', () => {
  const source: PresentationImportSource = {
    version: 1,
    documentId: 'doc',
    projectId: 'p',
    requestId: 'r',
    source: 'production',
    planRevision: 1,
    artifactDigest: 'a'.repeat(64),
    pages: [
      { id: 'a', title: 'A', sourceSlideId: '256#' },
      { id: 'b', title: 'B', sourceSlideId: '256#' },
    ],
  }
  const copy = parsePresentationImportSource(source)
  copy.pages[0]!.id = 'changed'
  expect(source.pages[0]!.id).toBe('a')
  for (const patch of [
    { pptxBase64: 'private body' },
    { source: 'compiled' },
    { planRevision: undefined },
    { artifactDigest: 'bad' },
    { pages: [{ id: 'a', title: 'A' }] },
    { pages: [{ id: 'a', title: 'A', sourceSlideId: '255#' }] },
    {
      pages: Array.from({ length: 33 }, (_, index) => ({
        id: `p${index}`,
        title: 'A',
        sourceSlideId: '256#',
      })),
    },
  ])
    expect(() => parsePresentationImportSource({ ...source, ...patch })).toThrow(
      'presentation_import_source_invalid',
    )
  const { planRevision: _revision, ...legacy } = source
  expect(
    parsePresentationImportSource({
      ...legacy,
      source: 'compiled',
      pages: [{ id: 'a', title: 'A' }],
    }).planRevision,
  ).toBeUndefined()
})
