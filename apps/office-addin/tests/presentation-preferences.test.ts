import { describe, expect, it } from 'vitest'
import { presentationPreferenceCandidates } from '../src/skills/powerpoint/presentation-preferences.js'
import type { PresentationHistoryEntry } from '../src/skills/powerpoint/presentation-change-history.js'

const base = {
  version: 1 as const,
  documentId: 'doc-1',
  projectId: 'project-1',
  requestId: 'request-1',
  artifactDigest: 'a'.repeat(64),
  pageId: 'page-1',
  hostSlideId: 'slide-1',
  shapeId: 'shape-1',
}
const text = (
  changeId: string,
  state: 'applied' | 'undone',
  after: string,
): PresentationHistoryEntry => ({
  id: `text:${changeId}`,
  sequence: Number(changeId.slice(1)),
  legacy: false,
  kind: 'text',
  record: { ...base, changeId, before: '旧标题', after, state },
})

describe('presentation preference candidates', () => {
  it('keeps only applied and scoped observations, without treating them as brand rules', () => {
    const other = text('c3', 'applied', '其它文档')
    if (other.kind !== 'text') throw new Error('invalid_fixture')
    const entries: PresentationHistoryEntry[] = [
      text('c1', 'applied', '更短的标题'),
      text('c2', 'undone', '撤销文本'),
      { ...other, record: { ...other.record, documentId: 'doc-2' } },
    ]
    expect(presentationPreferenceCandidates(entries, 'doc-1', 'project-1')).toEqual([
      {
        changeId: 'c1',
        kind: 'text',
        pageId: 'page-1',
        before: '旧标题',
        after: '更短的标题',
        status: 'candidate',
      },
    ])
  })
  it('rejects corrupted history instead of learning from it', () => {
    expect(() =>
      presentationPreferenceCandidates(
        [{ ...text('c1', 'applied', '新标题'), id: 'wrong' }],
        'doc-1',
        'project-1',
      ),
    ).toThrow('presentation_change_history_invalid')
  })
})
