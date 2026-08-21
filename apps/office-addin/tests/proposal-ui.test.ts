import { describe, expect, it } from 'vitest'
import { proposalPresentation } from '../src/App.js'

describe('generic proposal presentation', () => {
  it('preserves structured impact, preview, before, after, and code', () => {
    expect(
      proposalPresentation({
        id: 'p1',
        operation: 'edit_slide_xml',
        title: 'Edit XML',
        impact: { host: 'powerpoint', targets: ['slide-1'], count: 1 },
        preview: { nodes: 2 },
        fingerprint: 'fp',
        before: '<old/>',
        after: '<new/>',
        code: 'sync()',
      }),
    ).toEqual({
      title: 'Edit XML',
      host: 'powerpoint',
      count: 1,
      targets: ['slide-1'],
      before: '<old/>',
      after: '<new/>',
      preview: '{\n  "nodes": 2\n}',
      code: 'sync()',
    })
  })

  it('keeps the legacy append preview compatible', () => {
    expect(
      proposalPresentation({
        id: 'p2',
        operation: 'append',
        before: 'old',
        value: ' new',
        fingerprint: 'fp',
      }),
    ).toMatchObject({ title: 'Append to selection', before: 'old', after: 'old new' })
  })
})
