// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { PresentationProjectCard } from '../src/agent/presentation-project-card.js'
import type { PresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'

type Snapshot = ReturnType<PresentationProjectController['snapshot']>
const pending: Snapshot = {
  phase: 'idle',
  project: {
    projectId: 'p1',
    title: '季度计划',
    status: 'pending',
    latestRequestId: 'latest',
    latestCompiledRequestId: 'old',
    slideCount: 1,
    slides: [{ id: 's1', title: '目标' }],
    history: [
      { requestId: 'latest', sequence: 2, status: 'pending', slideCount: 1 },
      { requestId: 'old', sequence: 1, status: 'compiled', slideCount: 1 },
    ],
  },
}
const roots: ReturnType<typeof createRoot>[] = []
afterEach(async () => {
  for (const root of roots.splice(0)) await act(async () => root.unmount())
})
async function mount(snapshot: Snapshot, disabled = false) {
  const listeners = new Set<() => void>()
  const controller: PresentationProjectController = {
    pdfAvailable: vi.fn(() => false),
    exportProductionPdf: vi.fn(async () => {}),
    auditSources: vi.fn(async () => {}),
    snapshot: () => snapshot,
    subscribe: (listener) => {
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
    readDeliveryReport: vi.fn(async () => {}),
    exportDeliveryReport: vi.fn(async () => {}),
    recordIssueAction: vi.fn(async () => {}),
    refresh: vi.fn(async () => {}),
    restore: vi.fn(async () => {}),
    resume: vi.fn(async () => {}),
    runProduction: vi.fn(async () => {}),
    selectProduction: vi.fn(async () => {}),
    startProductionJob: vi.fn(async () => {}),
    pauseProductionJob: vi.fn(async () => {}),
    resumeProductionJob: vi.fn(async () => {}),
    cancelProductionJob: vi.fn(async () => {}),
    downloadProductionPage: vi.fn(async () => {}),
    prepareProduction: vi.fn(async () => {}),
    cancel: vi.fn(),
    clear: vi.fn(),
    prepareReconnect: vi.fn(),
  }
  const container = document.createElement('div')
  const root = createRoot(container)
  roots.push(root)
  await act(async () =>
    root.render(React.createElement(PresentationProjectCard, { controller, disabled })),
  )
  const button = (label: string) =>
    Array.from(container.querySelectorAll('button')).find((item) => item.textContent === label)!
  return {
    container,
    controller,
    button,
    update: async (next: Snapshot) => {
      snapshot = next
      await act(async () => listeners.forEach((listener) => listener()))
    },
  }
}
describe('presentation project recovery card', () => {
  it('shows planned attachment source readiness and recovers after refresh', async () => {
    const plan = benchmarkPlan()
    const attachmentId = 'a'.repeat(64)
    plan.sources[0]!.uri = `attachment:${attachmentId}`
    const project = {
      ...pending.project!,
      plan: { revision: 1, value: plan },
      sourcePreparation: [
        { sourceId: plan.sources[0]!.id, attachmentId, status: 'missing' as const },
      ],
    }
    const view = await mount({ phase: 'idle', project })
    expect(view.container.textContent).toContain('附件来源 0 / 1 已就绪')
    expect(view.container.textContent).toContain('合成基准：附件缺失')
    await view.update({
      phase: 'idle',
      project: {
        ...project,
        production: {
          projectId: 'p1',
          requestId: 'pages',
          planRevision: 1,
          status: 'partial',
          compiledCount: 0,
          total: 1,
          pages: [
            { id: 's1', title: '目标', state: 'failed', attempt: 1, error: 'source_unavailable' },
          ],
        },
      },
    })
    expect(view.container.textContent).toContain('引用的附件来源不可用')
    await view.update({
      phase: 'idle',
      project: {
        ...project,
        sourcePreparation: [{ sourceId: plan.sources[0]!.id, attachmentId, status: 'ready' }],
      },
    })
    expect(view.container.textContent).toContain('附件来源 1 / 1 已就绪')
    expect(view.container.textContent).toContain('合成基准：已解析，编译前仍会复核')
  })
  it('shows the PDF preview action only for a completed task with the PC capability', async () => {
    const project = {
      ...pending.project!,
      production: {
        projectId: 'p1',
        requestId: 'pages',
        planRevision: 1,
        status: 'compiled' as const,
        compiledCount: 1,
        total: 1,
        pages: [{ id: 's1', title: '目标', state: 'compiled' as const, attempt: 1 }],
      },
    }
    const view = await mount({ phase: 'idle', project })
    expect(view.button('导出 PDF 预览')).toBeUndefined()
    vi.mocked(view.controller.pdfAvailable!).mockReturnValue(true)
    await view.update({ phase: 'idle', project })
    await act(async () => view.button('导出 PDF 预览').click())
    expect(view.controller.exportProductionPdf).toHaveBeenCalledOnce()
    await view.update({
      phase: 'idle',
      project: {
        ...project,
        production: { ...project.production, status: 'partial', compiledCount: 0 },
      },
    })
    expect(view.button('导出 PDF 预览')).toBeUndefined()
  })

  it('shows literal source audit findings without claiming factual verification', async () => {
    const plan = benchmarkPlan()
    const attachmentId = 'a'.repeat(64)
    plan.sources[0]!.uri = `attachment:${attachmentId}`
    const view = await mount({
      phase: 'idle',
      project: { ...pending.project!, plan: { revision: 2, value: plan } },
      sourceAudit: {
        planRevision: 2,
        sources: [{ sourceId: plan.sources[0]!.id, attachmentId, status: 'not_found' }],
      },
    })
    expect(view.container.textContent).toContain('未在完整原文中找到引文')
    expect(view.container.textContent).toContain('不核验事实')
    await act(async () => view.button('核对计划引文与附件原文').click())
    expect(view.controller.auditSources).toHaveBeenCalledOnce()
    await view.update({
      phase: 'idle',
      project: { ...pending.project!, plan: { revision: 3, value: plan } },
      sourceAudit: {
        planRevision: 2,
        sources: [{ sourceId: plan.sources[0]!.id, attachmentId, status: 'not_found' }],
      },
    })
    expect(view.container.textContent).not.toContain('未在完整原文中找到引文')
  })
  it('shows local review comments and marks old-plan notes as historical', async () => {
    const view = await mount({
      phase: 'idle',
      project: {
        ...pending.project!,
        plan: { revision: 2, value: benchmarkPlan() },
        reviewComments: {
          revision: 1,
          openCount: 1,
          resolvedCount: 0,
          recent: [
            {
              id: 'comment-1',
              targetKind: 'source',
              targetId: 'source-1',
              authorLabel: '审阅人甲',
              text: '请复核来源',
              state: 'open',
              planRevision: 1,
              createdAt: '2026-09-28T00:00:00.000Z',
            },
          ],
        },
      },
    })
    expect(view.container.textContent).toContain('待处理 1')
    expect(view.container.textContent).toContain('旧计划第 1 版')
    expect(view.container.textContent).toContain('请复核来源')
    expect(view.container.textContent).toContain('未验证的显示标签')
  })
  it('shows saved pending state and resumes only the explicit latest request', async () => {
    const view = await mount(pending)
    expect(view.container.textContent).toContain('已保存，待编译')
    expect(view.container.textContent).toContain('目标')
    expect(view.container.querySelectorAll('details')).toHaveLength(1)
    await act(async () => view.button('继续编译').click())
    expect(view.controller.resume).toHaveBeenCalledWith('latest')
    await act(async () => view.button('恢复最近完成版本').click())
    expect(view.controller.restore).toHaveBeenCalledOnce()
    await act(async () => view.button('刷新').click())
    expect(view.controller.refresh).toHaveBeenCalledOnce()
  })
  it('shows honest checks and never offers resume for a completed project', async () => {
    const view = await mount({
      phase: 'idle',
      project: {
        ...pending.project!,
        status: 'compiled',
        checks: {
          structure: 'passed',
          geometry: 'warning',
          render: 'not_run',
          sources: 'not_verified',
          roundTrip: 'not_run',
        },
      },
    })
    expect(view.button('继续编译')).toBeUndefined()
    expect(view.container.textContent).toContain('已编译，尚未完成视觉验证')
    expect(view.container.textContent).toContain('几何检查：有警告')
    expect(view.container.textContent).toContain('来源：未核验')
    expect(view.container.textContent).toContain('Office 往返检查：未执行')
  })
  it('disables project operations while another agent action is active', async () => {
    const view = await mount(pending, true)
    expect(
      Array.from(view.container.querySelectorAll('button')).every((button) => button.disabled),
    ).toBe(true)
  })
  it('subscribes to operation state and permits cancellation', async () => {
    const view = await mount(pending)
    await view.update({ ...pending, phase: 'resuming' })
    expect(view.button('继续编译').disabled).toBe(true)
    expect(view.container.textContent).toContain('正在编译已保存版本')
    await act(async () => view.button('取消').click())
    expect(view.controller.cancel).toHaveBeenCalledOnce()
  })
  it('shows empty and upgrade states without presenting an error as an empty project', async () => {
    const view = await mount({ phase: 'idle' })
    expect(view.container.textContent).toContain('当前文档暂无已保存项目')
    await view.update({ phase: 'idle', error: '请升级 WisWork PC 后重试。' })
    expect(view.container.querySelector('[role="alert"]')?.textContent).toBe(
      '请升级 WisWork PC 后重试。',
    )
    expect(view.container.textContent).not.toContain('当前文档暂无已保存项目')
  })
})

it('shows a durable plan before compilation and labels its unverified evidence', async () => {
  const { benchmarkPlan } =
    await import('../../../packages/pptx-engine/tests/fixtures/presentation-plan.js')
  const plan = benchmarkPlan()
  const view = await mount({
    phase: 'idle',
    project: {
      projectId: plan.projectId,
      title: plan.title,
      status: 'planned',
      slideCount: plan.slides.length,
      slides: plan.slides.map(({ id, title }) => ({ id, title })),
      history: [],
      plan: { revision: 1, value: plan },
    },
  })
  expect(view.container.textContent).toContain('计划已保存，尚未编译')
  expect(view.container.textContent).toContain(plan.brief.objective)
  expect(view.container.textContent).toContain('未核验')
  expect(view.button('继续编译')).toBeUndefined()
})
it('shows page production states and a separate continue action without claiming delivery', async () => {
  const view = await mount({
    ...pending,
    project: {
      ...pending.project!,
      production: {
        projectId: 'p1',
        requestId: 'pages',
        planRevision: 1,
        status: 'partial',
        compiledCount: 1,
        total: 2,
        pages: [
          { id: 'a', title: 'First', state: 'compiled', attempt: 1 },
          { id: 'b', title: 'Second', state: 'failed', attempt: 2, error: 'compile_failed' },
        ],
      },
    },
  })
  expect(view.container.textContent).toContain('1 / 2')
  expect(view.container.textContent).toContain('已编译（未导入验收）')
  expect(view.container.textContent).toContain('失败待重试')
  await act(async () => view.button('继续页任务').click())
  expect(view.controller.runProduction).toHaveBeenCalledWith('pages')
})
it('warns when a planned project has production frozen against an older plan', async () => {
  const { benchmarkPlan } =
    await import('../../../packages/pptx-engine/tests/fixtures/presentation-plan.js')
  const view = await mount({
    phase: 'idle',
    project: {
      ...pending.project!,
      status: 'planned',
      latestRequestId: undefined,
      latestCompiledRequestId: undefined,
      history: [],
      plan: { revision: 2, value: benchmarkPlan() },
      production: {
        projectId: 'p1',
        requestId: 'old-pages',
        planRevision: 1,
        status: 'pending',
        compiledCount: 0,
        total: 1,
        pages: [{ id: 'a', title: 'First', state: 'pending', attempt: 0 }],
      },
    },
  })
  expect(view.container.textContent).toContain('页任务使用旧计划，继续任务按原快照，不代表当前计划')
  await act(async () => view.button('继续页任务').click())
  expect(view.controller.runProduction).toHaveBeenCalledWith('old-pages')
})
it('labels a derived page task with its parent and pending host replacement', async () => {
  const view = await mount({
    ...pending,
    project: {
      ...pending.project!,
      production: {
        projectId: 'p1',
        requestId: 'child',
        planRevision: 1,
        status: 'pending',
        compiledCount: 0,
        total: 1,
        pages: [{ id: 'a', title: 'First', state: 'pending', attempt: 0 }],
        revision: { parentRequestId: 'parent', pageId: 'a', parentInputDigest: 'a'.repeat(64) },
      },
    },
  })
  expect(view.container.textContent).toContain('parent')
  expect(view.container.textContent).toContain('目标页：a')
  expect(view.container.textContent).toContain('尚未替换当前页')
})
it('shows background pause, retained event history and compiled page download', async () => {
  const view = await mount({
    ...pending,
    project: {
      ...pending.project!,
      productionJob: {
        version: 1,
        projectId: 'p1',
        documentId: 'd',
        requestId: 'pages',
        inputDigest: 'a'.repeat(64),
        planDigest: 'b'.repeat(64),
        planRevision: 1,
        revision: 130,
        state: 'running',
        events: [{ sequence: 130, createdAt: '2026-09-24T00:00:00.000Z', type: 'run.started' }],
      },
      production: {
        projectId: 'p1',
        requestId: 'pages',
        planRevision: 1,
        status: 'partial',
        compiledCount: 1,
        total: 2,
        pages: [
          { id: 'a', title: 'A', state: 'compiled', attempt: 1 },
          { id: 'b', title: 'B', state: 'pending', attempt: 0 },
        ],
      },
    },
  })
  expect(view.container.textContent).toContain('更早历史已截断')
  expect(view.button('继续页任务')).toBeUndefined()
  await act(async () => view.button('暂停后台任务').click())
  expect(view.controller.pauseProductionJob).toHaveBeenCalledWith('pages')
  await act(async () => view.button('保存单页到附件：A').click())
  expect(view.controller.downloadProductionPage).toHaveBeenCalledWith('a')
})
it('lets users select an older paused task by request and progress', async () => {
  const view = await mount({
    ...pending,
    project: {
      ...pending.project!,
      productionTasks: [
        {
          requestId: 'new',
          sequence: 2,
          planRevision: 1,
          status: 'pending',
          compiledCount: 0,
          total: 1,
        },
        {
          requestId: 'old',
          sequence: 1,
          planRevision: 1,
          status: 'pending',
          compiledCount: 0,
          total: 1,
          jobState: 'paused',
        },
      ],
      production: {
        projectId: 'p1',
        requestId: 'new',
        planRevision: 1,
        status: 'pending',
        compiledCount: 0,
        total: 1,
        pages: [{ id: 'a', title: 'A', state: 'pending', attempt: 0 }],
      },
    },
  })
  const select = view.container.querySelector('select')!
  expect(select.getAttribute('aria-label')).toBe('选择页生产任务')
  expect(select.textContent).toContain('old')
  expect(select.textContent).toContain('已暂停')
  await act(async () => {
    select.value = 'old'
    select.dispatchEvent(new Event('change', { bubbles: true }))
  })
  expect(view.controller.selectProduction).toHaveBeenCalledWith('old')
})
it('offers downloads but no bulk import preparation for a compiled revision', async () => {
  const view = await mount({
    ...pending,
    project: {
      ...pending.project!,
      production: {
        projectId: 'p1',
        requestId: 'child',
        planRevision: 1,
        status: 'compiled',
        compiledCount: 1,
        total: 1,
        pages: [{ id: 'a', title: 'A', state: 'compiled', attempt: 1 }],
        revision: { parentRequestId: 'parent', pageId: 'a', parentInputDigest: 'a'.repeat(64) },
      },
    },
  })
  expect(view.button('准备完整成果导入')).toBeUndefined()
  expect(view.button('保存单页到附件：A')).toBeTruthy()
  expect(view.container.textContent).toContain('宿主页替换需单独确认')
})

it('shows evidence warnings, explicit issue disposition inputs, and export action', async () => {
  const { deliveryReportFixture } = await import('./presentation-delivery-fixture.js')
  const report = await deliveryReportFixture()
  const view = await mount({
    ...pending,
    deliveryReport: report,
    project: {
      ...pending.project!,
      production: {
        projectId: report.projectId,
        requestId: 'r',
        planRevision: 1,
        status: 'pending',
        compiledCount: 0,
        total: 1,
        pages: [{ id: 'a', title: 'A', state: 'pending', attempt: 0 }],
      },
    },
  })
  expect(view.container.textContent).toContain('算术复现')
  expect(view.container.textContent).toContain('已说明不会关闭机器发现')
  expect(view.container.querySelector('textarea')?.value).toBe('')
  expect(view.controller.recordIssueAction).not.toHaveBeenCalled()
  const form = view.container.querySelector('form')!
  const textarea = form.querySelector('textarea')!
  const disposition = form.querySelector('select')!
  await act(async () => {
    disposition.value = 'deferred'
    disposition.dispatchEvent(new Event('change', { bubbles: true }))
  })
  await act(async () => {
    Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(
      textarea,
      'Waiting for evidence',
    )
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
  })
  await act(async () =>
    form.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })),
  )
  expect(view.controller.recordIssueAction).toHaveBeenCalledWith(
    expect.objectContaining({ state: 'deferred', note: 'Waiting for evidence' }),
  )
  await act(async () => view.button('导出证据 JSON + Markdown 到附件').click())
  expect(view.controller.exportDeliveryReport).toHaveBeenCalledOnce()
})
