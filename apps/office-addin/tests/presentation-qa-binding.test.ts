import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { expect, it } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { type PresentationQaRecord } from '../src/skills/powerpoint/presentation-qa'
import { PresentationQaCard } from '../src/agent/presentation-qa-card'

function fixture() {
  const values = new Map<string, string>()
  let location = 'file://deck.pptx',
    fail = false
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save: async () => {
      if (fail) throw new Error('save_failed')
    },
    location: () => location,
  }
  const create = () => createPresentationDocumentBinding(settings, () => 'document-qa')
  return {
    binding: create(),
    create,
    values,
    fail: () => {
      fail = true
    },
    move: () => {
      location = 'file://other.pptx'
    },
  }
}
function record(documentId: string, requestId = 'request-1'): PresentationQaRecord {
  return {
    version: 1,
    documentId,
    projectId: 'project',
    requestId,
    artifactDigest: 'a'.repeat(64),
    pages: [
      {
        pageId: 'page1',
        title: '第一页',
        hostSlideId: '256',
        capturedAt: '2026-09-23T08:00:00.000Z',
        screenshotDigest: 'b'.repeat(64),
        screenshotBytes: 128,
        structure: {
          status: 'passed',
          shapeCount: 1,
          overflowCount: 0,
          overlapCount: 0,
          shapesTruncated: false,
          overlapsTruncated: false,
        },
        visual: { status: 'needs_review' },
      },
    ],
  }
}
it('persists bounded QA metadata and reads it through a fresh binding', async () => {
  const f = fixture(),
    value = record(await f.binding.documentId())
  await f.binding.writeQa('project/request-1', value)
  expect(f.create().readQa('project/request-1')).toEqual(value)
  expect([...f.values.values()].join('')).not.toContain('base64')
})
it('restores the prior review if settings save fails and rejects document mismatch', async () => {
  const f = fixture(),
    value = record(await f.binding.documentId())
  await f.binding.writeQa('project/request-1', value)
  f.fail()
  const next = structuredClone(value)
  next.pages[0]!.visual = {
    status: 'pass',
    reviewer: 'agent',
    notes: '布局清晰',
    reviewedAt: '2026-09-23T08:01:00.000Z',
  }
  await expect(f.binding.writeQa('project/request-1', next)).rejects.toThrow('save_failed')
  expect(f.binding.readQa('project/request-1')).toEqual(value)
  f.move()
  await expect(f.binding.writeQa('project/request-1', value)).rejects.toThrow(
    'presentation_document_changed',
  )
})
it('rejects malformed persisted QA and limits retained request history', async () => {
  const f = fixture(),
    doc = await f.binding.documentId()
  for (let i = 0; i < 8; i++)
    await f.binding.writeQa(`project/request-${i}`, record(doc, `request-${i}`))
  await expect(f.binding.writeQa('project/request-8', record(doc, 'request-8'))).rejects.toThrow(
    'presentation_qa_history_full',
  )
  f.values.set(
    'wiswork.presentation.qa.v1',
    JSON.stringify({ 'project/request-1': { ...record(doc), extra: 'untrusted' } }),
  )
  expect(() => f.binding.readQa('project/request-1')).toThrow('presentation_qa_state_invalid')
})
it('labels a restored Agent assessment as historical and escapes reviewer text', () => {
  const value = record('doc')
  value.pages[0]!.visual = {
    status: 'pass',
    reviewer: 'agent',
    notes: '<script>bad()</script>',
    reviewedAt: '2026-09-23T08:01:00.000Z',
  }
  const html = renderToStaticMarkup(
    createElement(PresentationQaCard, {
      controller: { read: () => value, revision: () => 0, subscribe: () => () => undefined },
    }),
  )
  expect(html).toContain('历史检查记录')
  expect(html).toContain('Agent 判断通过')
  expect(html).not.toContain('<script>')
})
