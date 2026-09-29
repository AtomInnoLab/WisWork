// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import { PresentationProjectCard } from '../src/agent/presentation-project-card.js'
import { comparisonFixture, comparisonReport } from './presentation-feedback-comparison-fixture.js'
async function mount(report = comparisonReport()) {
  const f = comparisonFixture()
  f.setComparison(report)
  await f.controller.refresh()
  const container = document.createElement('div'),
    root = createRoot(container)
  const render = async (disabled = false) =>
    act(async () =>
      root.render(
        React.createElement(PresentationProjectCard, { controller: f.controller, disabled }),
      ),
    )
  await render()
  const button = (label: string) =>
    Array.from(container.querySelectorAll('button')).find((item) => item.textContent === label)!
  const select = async (value = 'baseline') =>
    act(async () => {
      const input = container.querySelector('[aria-label="选择反馈对照基线"]') as HTMLSelectElement
      input.value = value
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })
  return { ...f, root, container, render, button, select }
}
it('compares only on a real click and exports the original frozen observation locally', async () => {
  const report = comparisonReport(),
    f = await mount(report),
    create = vi.fn((_blob: Blob) => 'blob:comparison'),
    anchor = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})
  vi.stubGlobal(
    'URL',
    class extends URL {
      static createObjectURL = create
      static revokeObjectURL = vi.fn()
    },
  )
  try {
    const selector = f.container.querySelector(
      '[aria-label="选择反馈对照基线"]',
    ) as HTMLSelectElement
    expect(Array.from(selector.options).map((option) => option.value)).toEqual(['', 'baseline'])
    expect(f.button('读取两次制作反馈对照').disabled).toBe(true)
    await f.select()
    expect(
      f.request.mock.calls.some(
        ([body]) => (body as { operation: string }).operation === 'production_feedback_compare',
      ),
    ).toBe(false)
    await act(async () => f.button('读取两次制作反馈对照').click())
    const panel = f.container.querySelector('[aria-label="两次制作反馈对照"]')!
    for (const text of [
      '两次制作反馈对照，不代表技能效果或验收通过',
      '通用基线',
      '当前制作',
      '冻结计划第 1 版',
      '冻结计划第 2 版',
      '反馈版本：1',
      '已评估 2 / 2 页',
      '已评估 5 / 5 页',
      '未评估（未知） 0 页',
      '完整',
      '制作要求：相同',
      '来源声明：相同',
      '主张声明：相同',
      '研究绑定：相同',
      '样式：相同',
      '品牌规范：相同',
      '有效并行方式：相同',
      '需要修正页数 -1 页',
      '-80.0 个百分点',
      '不作因果或显著性结论',
    ])
      expect(panel.textContent).toContain(text)
    f.setComparison(comparisonReport({ candidateRevision: 2, candidateNeeds: 3 }))
    const count = f.request.mock.calls.length
    await act(async () => f.button('下载制作反馈对照 JSON').click())
    expect(f.request).toHaveBeenCalledTimes(count)
    expect(anchor).toHaveBeenCalledOnce()
    const json = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = () => resolve(String(reader.result))
      reader.onerror = () => reject(reader.error)
      reader.readAsText(create.mock.calls[0]![0])
    })
    expect(JSON.parse(json)).toEqual(report)
    expect(JSON.parse(json).effect).toBe('not_verified')
    expect(f.executeTool).not.toHaveBeenCalled()
    await f.select('')
    expect(f.container.querySelector('[aria-label="对照基线版本"]')).toBeNull()
    expect(f.button('下载制作反馈对照 JSON')).toBeUndefined()
  } finally {
    await act(async () => f.root.unmount())
    anchor.mockRestore()
    vi.unstubAllGlobals()
  }
})
it.each([{ missing: true }, { differentStyle: true }])(
  'keeps incomplete evaluation or unequal conditions unknown rather than zero %j',
  async (options) => {
    const f = await mount(comparisonReport(options))
    try {
      await f.select()
      await act(async () => f.button('读取两次制作反馈对照').click())
      const panel = f.container.querySelector('[aria-label="两次制作反馈对照"]')!
      expect(panel.textContent).toContain('不可计算')
      expect(panel.textContent).not.toContain('需要修正页数 0 页')
      if (options.missing) {
        expect(panel.textContent).toContain('未评估（未知） 5 页')
        expect(panel.textContent).toContain('已评估页需要修正比例：未知')
        expect(panel.textContent).toContain('当前任务没有保存人工评价')
      } else {
        expect(panel.textContent).toContain('样式：不同')
        expect(panel.textContent).toContain('两次制作的声明输入条件不同')
      }
      await f.render(true)
      expect(
        (f.container.querySelector('[aria-label="选择反馈对照基线"]') as HTMLSelectElement)
          .disabled,
      ).toBe(true)
      expect(f.button('读取两次制作反馈对照').disabled).toBe(true)
      await act(async () => f.controller.clear())
      expect(f.container.querySelector('[aria-label="两次制作反馈对照"]')).toBeNull()
    } finally {
      await act(async () => f.root.unmount())
    }
  },
)
