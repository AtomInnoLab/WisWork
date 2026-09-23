import React from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { expect, it } from 'vitest'
import { PresentationQaCard } from '../src/agent/presentation-qa-card.js'
import type { PresentationQaRecord } from '../src/skills/powerpoint/presentation-qa.js'
it('makes post-edit recapture take precedence over a historical visual pass', () => {
  const record: PresentationQaRecord = {
    version: 1,
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    artifactDigest: 'a'.repeat(64),
    pages: [
      {
        pageId: 'page',
        title: 'Title',
        hostSlideId: 'host',
        capturedAt: '2026-09-23T00:00:00.000Z',
        screenshotDigest: 'b'.repeat(64),
        screenshotBytes: 68,
        recheckRequired: true,
        structure: {
          status: 'passed',
          shapeCount: 0,
          overflowCount: 0,
          overlapCount: 0,
          shapesTruncated: false,
          overlapsTruncated: false,
        },
        visual: {
          status: 'pass',
          reviewer: 'agent',
          notes: 'Earlier title review',
          reviewedAt: '2026-09-23T00:00:00.000Z',
        },
      },
    ],
  }
  const html = renderToStaticMarkup(
    React.createElement(PresentationQaCard, {
      controller: { read: () => record, revision: () => 0, subscribe: () => () => {} },
    }),
  )
  expect(html).toContain('修改后需重新采集')
  expect(html).toContain('历史视觉')
  expect(html).toContain('Earlier title review')
  expect(html).not.toContain('视觉：Agent 判断通过')
})
