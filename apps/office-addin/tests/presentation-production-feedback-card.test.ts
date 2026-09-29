// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import { PresentationProjectCard } from '../src/agent/presentation-project-card.js'
import { createPresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'
import type { PresentationProductionFeedbackLedger } from '@wiswork/project-store/presentation-feedback'
async function fixture() {
  const production = {
    projectId: 'p',
    requestId: 'r',
    planRevision: 1,
    status: 'compiled',
    compiledCount: 2,
    total: 2,
    pages: [
      { id: 'a', title: '甲', state: 'compiled', attempt: 1 },
      { id: 'b', title: '乙', state: 'compiled', attempt: 1 },
    ],
  }
  const project = {
    projectId: 'p',
    title: '反馈项目',
    status: 'compiled',
    latestRequestId: 'r',
    latestCompiledRequestId: 'r',
    slideCount: 2,
    slides: [
      { id: 'a', title: '甲' },
      { id: 'b', title: '乙' },
    ],
    history: [{ requestId: 'r', sequence: 1, status: 'compiled', slideCount: 2 }],
    production,
    checks: {
      structure: 'passed',
      geometry: 'passed',
      render: 'not_run',
      sources: 'not_verified',
      roundTrip: 'not_run',
    },
  }
  let feedback: PresentationProductionFeedbackLedger | null = null
  const request = vi.fn(async (body: unknown) => {
    const input = body as {
      operation: string
      pages: PresentationProductionFeedbackLedger['snapshots'][number]['pages']
    }
    if (input.operation === 'status') return new Response(JSON.stringify(project))
    if (input.operation === 'production_job_status')
      return new Response(JSON.stringify({ error: 'invalid_request' }))
    if (input.operation === 'production_feedback_record')
      feedback = {
        version: 1,
        source: 'user_reported',
        documentId: 'doc',
        projectId: 'p',
        requestId: 'r',
        inputDigest: 'a'.repeat(64),
        planDigest: 'b'.repeat(64),
        planRevision: 1,
        pageIds: ['a', 'b'],
        revision: 1,
        snapshots: [
          {
            revision: 1,
            recordedAt: '2026-09-29T00:00:00.000Z',
            pages: ['a', 'b'].map(
              (pageId) =>
                input.pages.find((page) => page.pageId === pageId) ?? {
                  pageId,
                  status: 'not_evaluated',
                },
            ),
          },
        ],
      }
    return new Response(JSON.stringify({ feedback }))
  })
  const executeTool = vi.fn(async () => ({ output: '{}', mutated: false, summary: '' })),
    controller = createPresentationProjectController({
      request,
      executeTool,
      documentId: async () => 'doc',
      available: () => true,
      lastProject: () => 'p',
    })
  await controller.refresh()
  const container = document.createElement('div'),
    root = createRoot(container)
  const render = async (disabled = false) =>
    act(async () =>
      root.render(React.createElement(PresentationProjectCard, { controller, disabled })),
    )
  await render()
  const button = (text: string) =>
    Array.from(container.querySelectorAll('button')).find((item) => item.textContent === text)!
  return {
    controller,
    request,
    executeTool,
    container,
    root,
    render,
    button,
    feedback: () => feedback,
  }
}
it('records only a real save click, keeps unknown denominator and downloads the exact saved user feedback locally', async () => {
  const f = await fixture()
  const anchor = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {}),
    create = vi.fn((_blob: Blob) => 'blob:feedback'),
    revoke = vi.fn()
  vi.stubGlobal(
    'URL',
    class extends URL {
      static createObjectURL = create
      static revokeObjectURL = revoke
    },
  )
  try {
    expect(
      f.request.mock.calls.some(
        ([body]) => (body as { operation: string }).operation === 'production_feedback_record',
      ),
    ).toBe(false)
    expect(f.container.querySelector('[aria-label="人工修正评价 甲"]')).toBeNull()
    await act(async () => f.button('读取当前任务反馈').click())
    expect(f.container.textContent).toContain('已评估 0 / 2 页')
    expect(f.container.textContent).toContain('未评估（未知） 2 页')
    expect(f.container.textContent).toContain('需要修正比例：未计算')
    const select = f.container.querySelector('[aria-label="人工修正评价 甲"]') as HTMLSelectElement
    await act(async () => {
      select.value = 'needs_correction'
      select.dispatchEvent(new Event('change', { bubbles: true }))
    })
    const textarea = f.container.querySelector(
      '[aria-label="人工修正说明 甲"]',
    ) as HTMLTextAreaElement
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(
        textarea,
        '<script>人工判断</script>',
      )
      textarea.dispatchEvent(new Event('input', { bubbles: true }))
    })
    await act(async () => {
      f.button('保存人工修正反馈').click()
      f.button('保存人工修正反馈').click()
    })
    expect(
      f.request.mock.calls.filter(
        ([body]) => (body as { operation: string }).operation === 'production_feedback_record',
      ),
    ).toHaveLength(1)
    expect(f.container.textContent).toContain('已评估 1 / 2 页')
    expect(f.container.textContent).toContain('需要人工修正 1 页')
    expect(f.container.textContent).toContain('未评估（未知） 1 页')
    expect(f.container.textContent).toContain('已评估页需要修正比例：1 / 1')
    expect(f.container.querySelector('script')).toBeNull()
    const count = f.request.mock.calls.length
    await act(async () => f.button('下载人工修正反馈 JSON').click())
    expect(f.request).toHaveBeenCalledTimes(count)
    expect(anchor).toHaveBeenCalledOnce()
    expect(create).toHaveBeenCalledOnce()
    const json = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = () => resolve(String(reader.result))
      reader.onerror = () => reject(reader.error)
      reader.readAsText(create.mock.calls[0]![0])
    })
    expect(JSON.parse(json)).toEqual(f.feedback())
    expect(JSON.parse(json).source).toBe('user_reported')
    expect(f.executeTool).not.toHaveBeenCalled()
    await act(async () => f.controller.clear())
    expect(f.container.querySelector('[aria-label="人工修正反馈"]')).toBeNull()
  } finally {
    await act(async () => f.root.unmount())
    anchor.mockRestore()
    vi.unstubAllGlobals()
  }
})
it('blocks saving overlong notes and disables feedback while the product is busy', async () => {
  const f = await fixture()
  try {
    await act(async () => f.button('读取当前任务反馈').click())
    const textarea = f.container.querySelector(
      '[aria-label="人工修正说明 甲"]',
    ) as HTMLTextAreaElement
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(
        textarea,
        '长'.repeat(1000),
      )
      textarea.dispatchEvent(new Event('input', { bubbles: true }))
    })
    expect(f.container.textContent).toContain('说明过长，请缩短后保存。')
    expect(f.button('保存人工修正反馈').disabled).toBe(true)
    await f.render(true)
    expect(f.button('读取当前任务反馈').disabled).toBe(true)
    expect(
      f.request.mock.calls.some(
        ([body]) => (body as { operation: string }).operation === 'production_feedback_record',
      ),
    ).toBe(false)
  } finally {
    await act(async () => f.root.unmount())
  }
})

it('saves both edited pages together without silently discarding another page draft', async () => {
  const f = await fixture()
  try {
    await act(async () => f.button('读取当前任务反馈').click())
    for (const title of ['甲', '乙']) {
      const select = f.container.querySelector(
        `[aria-label="人工修正评价 ${title}"]`,
      ) as HTMLSelectElement
      await act(async () => {
        select.value = 'needs_correction'
        select.dispatchEvent(new Event('change', { bubbles: true }))
      })
    }
    await act(async () => f.button('保存人工修正反馈').click())
    expect(
      (f.container.querySelector('[aria-label="人工修正评价 乙"]') as HTMLSelectElement).value,
    ).toBe('needs_correction')
    expect(
      f
        .feedback()
        ?.snapshots.at(-1)
        ?.pages.map((page) => page.status),
    ).toEqual(['needs_correction', 'needs_correction'])
  } finally {
    await act(async () => f.root.unmount())
  }
})
