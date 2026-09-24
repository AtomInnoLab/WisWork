import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { expect, it } from 'vitest'
import type { PresentationImportProgress } from '../src/skills/powerpoint/presentation-page-delivery'
import { PresentationImportProgressCard } from '../src/agent/presentation-import-progress'

it('shows completed, pending and uncertain pages as saved records without claiming visual QA', () => {
  const markup = renderToStaticMarkup(
    createElement(PresentationImportProgressCard, {
      controller: {
        revision: () => 1,
        subscribe: () => () => undefined,
        read: (): PresentationImportProgress => ({
          total: 3,
          completed: 1,
          status: 'uncertain',
          pages: [
            { id: 'one', title: '第一页', state: 'complete', completedAt: '2026-09-24T00:00:00.000Z' },
            { id: 'two', title: '第二页', state: 'uncertain', startedAt: '2026-09-24T00:01:00.000Z' },
            { id: 'three', title: '第三页', state: 'pending' },
          ],
        }),
      },
    }),
  )
  expect(markup).toContain('1 / 3')
  expect(markup).toContain('停止自动重试')
  expect(markup).toContain('已记录完成')
  expect(markup).toContain('待导入')
  expect(markup).toContain('dateTime="2026-09-24T00:00:00.000Z"')
  expect(markup).toContain('dateTime="2026-09-24T00:01:00.000Z"')
})
it('shows invalid persisted state as an actionable error', () => {
  const markup = renderToStaticMarkup(
    createElement(PresentationImportProgressCard, {
      controller: {
        revision: () => 0,
        subscribe: () => () => undefined,
        read: () => {
          throw new Error('invalid_state')
        },
      },
    }),
  )
  expect(markup).toContain('避免重复插页')
})
