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
async function mount(snapshot: Snapshot, disabled = false, onEndFrontend?: () => void) {
  const listeners = new Set<() => void>()
  const controller: PresentationProjectController = {
    editPlan: vi.fn(async () => {}),
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
    root.render(
      React.createElement(PresentationProjectCard, { controller, disabled, onEndFrontend }),
    ),
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
  it('shows recorded host association, missing pages and historical identity without claiming current content or QA', async () => {
    const plan = benchmarkPlan()
    const f = await mount({
      ...pending,
      project: {
        ...pending.project!,
        plan: { revision: 3, value: plan },
        hostAssociations: {
          pages: [
            {
              pageId: plan.slides[0]!.id,
              hostPages: [
                {
                  requestId: 'old',
                  slideId: 'host-old',
                  planRevision: 1,
                  revisionRelation: 'historical',
                  presence: 'present',
                  sourceProof: 'digest',
                },
                {
                  requestId: 'unknown',
                  slideId: 'gone',
                  revisionRelation: 'unknown',
                  presence: 'missing',
                },
              ],
            },
          ],
          legacyImports: 2,
          uncertainImports: 1,
        },
      },
    })
    expect(f.container.querySelector('[aria-label="计划与宿主页关联"]')).not.toBeNull()
    expect(f.container.textContent).toContain('历史计划第 1 版')
    expect(f.container.textContent).toContain('源产物摘要已匹配')
    expect(f.container.textContent).toContain('宿主页已缺失')
    expect(f.container.textContent).toContain('计划修订未知')
    expect(f.container.textContent).toContain('2 份旧导入记录缺少页级身份')
    expect(f.container.textContent).toContain('不证明页面内容仍匹配计划或已通过验收')
  })
  it('shows persistent page locks and dispatches an explicit unlock while preventing page removal', async () => {
    const plan = benchmarkPlan()
    plan.slides[1]!.locked = true
    const f = await mount({
      ...pending,
      project: { ...pending.project!, plan: { revision: 5, value: plan } },
    })
    expect(f.container.textContent).toContain('计划页已锁定')
    const unlock = f.container.querySelector<HTMLButtonElement>(
      'button[aria-label="解除计划页锁定 2：' + plan.slides[1]!.title + '"]',
    )!
    expect(unlock).not.toBeNull()
    await act(async () => unlock.click())
    expect(f.controller.editPlan).toHaveBeenCalledWith(5, {
      kind: 'lock',
      pageId: plan.slides[1]!.id,
      locked: false,
    })
    expect(
      f.container.querySelector<HTMLButtonElement>(
        'button[aria-label="删除计划页 2：' + plan.slides[1]!.title + '"]',
      )!.disabled,
    ).toBe(true)
  })
  it('restores the chosen historical plan into the displayed current revision and resets stale choices', async () => {
    const plan = benchmarkPlan()
    const revisions = [1, 2, 3].map((revision) => ({
      revision,
      inputDigest: 'a'.repeat(64),
      createdAt: '2026-09-29T00:00:00.000Z',
    }))
    const project = { ...pending.project!, plan: { revision: 3, value: plan, revisions } }
    const view = await mount({ phase: 'idle', project })
    const select = view.container.querySelector(
      '[aria-label="选择历史计划版本"]',
    ) as HTMLSelectElement
    expect(select.value).toBe('2')
    await act(async () => {
      select.value = '1'
      select.dispatchEvent(new Event('change', { bubbles: true }))
    })
    await act(async () => view.button('恢复为新计划修订').click())
    expect(view.controller.editPlan).toHaveBeenCalledWith(3, { kind: 'restore', revision: 1 })
    expect(view.container.textContent).toContain('恢复将创建新修订；已有 PowerPoint 页面保留')
    await view.update({
      phase: 'idle',
      project: {
        ...project,
        plan: {
          ...project.plan,
          revision: 4,
          revisions: [...revisions, { ...revisions[0]!, revision: 4 }],
        },
      },
    })
    expect(
      (view.container.querySelector('[aria-label="选择历史计划版本"]') as HTMLSelectElement).value,
    ).toBe('3')
    await view.update({ phase: 'planning', project })
    expect(view.button('恢复为新计划修订').disabled).toBe(true)
  })
  it('edits the visible plan revision through accessible page controls and disables boundaries', async () => {
    const plan = benchmarkPlan()
    const view = await mount({
      phase: 'idle',
      project: { ...pending.project!, plan: { revision: 3, value: plan } },
    })
    const rows = view.container.querySelectorAll('[aria-label="逐页施工图"] > li')
    const up = rows[0]!.querySelector('button[aria-label^="上移计划页"]') as HTMLButtonElement
    const down = rows[1]!.querySelector('button[aria-label^="下移计划页"]') as HTMLButtonElement
    const remove = rows[1]!.querySelector('button[aria-label^="删除计划页"]') as HTMLButtonElement
    expect(up.disabled).toBe(true)
    expect(
      (
        rows[rows.length - 1]!.querySelector(
          'button[aria-label^="下移计划页"]',
        ) as HTMLButtonElement
      ).disabled,
    ).toBe(true)
    await act(async () => down.click())
    expect(view.controller.editPlan).toHaveBeenCalledWith(3, {
      kind: 'move',
      pageId: plan.slides[1]!.id,
      direction: 'down',
    })
    await act(async () => remove.click())
    expect(view.controller.editPlan).toHaveBeenLastCalledWith(3, {
      kind: 'delete',
      pageId: plan.slides[1]!.id,
    })
    expect(view.container.textContent).toContain('只调整制作计划；已有 PowerPoint 页面保留')
    await view.update({
      phase: 'planning',
      project: { ...pending.project!, plan: { revision: 3, value: plan } },
    })
    expect(
      Array.from(view.container.querySelectorAll('[aria-label="逐页施工图"] button')).every(
        (button) => (button as HTMLButtonElement).disabled,
      ),
    ).toBe(true)
  })
  it('shows retained compile provenance without marking host delivery or QA complete', async () => {
    const view = await mount({
      phase: 'idle',
      project: {
        ...pending.project!,
        production: {
          projectId: 'p1',
          requestId: 'revised',
          planRevision: 2,
          status: 'compiled',
          compiledCount: 1,
          total: 1,
          pages: [
            {
              id: 's1',
              title: '保留页',
              state: 'compiled',
              attempt: 1,
              reusedFromRequestId: 'original',
            },
          ],
        },
      },
    })
    expect(view.container.textContent).toContain('保留既有编译成果 · 来源任务 original')
    expect(view.container.textContent).toContain('当前内容与视觉仍需核验')
    expect(view.container.textContent).toContain('已编译（未导入验收）')
  })
  it('projects the current blueprint and style without treating saved plans as verified', async () => {
    const plan = benchmarkPlan()
    plan.style.fontFallbacks = ['Arial', '微软雅黑']
    plan.brief.constraints = ['禁止虚构数字']
    plan.slides[0]!.layout = 'chart'
    plan.slides[0]!.requiredAssets = ['收入趋势图']
    plan.slides[0]!.claimIds = [plan.claims[0]!.id]
    plan.slides[1]!.dependsOn = [plan.slides[0]!.id]
    const view = await mount({
      phase: 'idle',
      project: { ...pending.project!, plan: { revision: 1, value: plan } },
    })
    const blueprint = view.container.querySelector('[aria-label="逐页施工图"]')!
    expect(blueprint.textContent).toContain('页型：图表')
    expect(blueprint.textContent).toContain(plan.claims[0]!.statement)
    expect(blueprint.textContent).toContain('依据：合成基准 · 待核验')
    expect(blueprint.textContent).toContain('所需素材：收入趋势图')
    expect(blueprint.textContent).toContain(`依赖页面：${plan.slides[0]!.title}`)
    expect(blueprint.textContent).toContain('可编辑文本')
    expect(view.container.querySelector('[aria-label="样式规范"]')!.textContent).toContain(
      'Arial、微软雅黑',
    )
    expect(view.container.textContent).toContain('约束：禁止虚构数字')
    const updated = structuredClone(plan)
    updated.claims[0]!.statement = '修订后的结论'
    updated.slides[0]!.requiredAssets = []
    await view.update({
      phase: 'idle',
      project: { ...pending.project!, plan: { revision: 2, value: updated } },
    })
    expect(view.container.querySelector('[aria-label="逐页施工图"]')!.textContent).toContain(
      '修订后的结论',
    )
    expect(view.container.querySelector('[aria-label="逐页施工图"]')!.textContent).not.toContain(
      '收入趋势图',
    )
    expect(view.controller.prepareProduction).not.toHaveBeenCalled()
  })

  it('shows at most three accessible brand layout previews without fetching brand assets', async () => {
    const plan = benchmarkPlan()
    plan.brandKit = {
      id: 'brand',
      revision: 2,
      name: '示例品牌',
      allowedColors: ['112233'],
      logo: { assetId: 'logo', assetDigest: 'a'.repeat(64), placement: 'all' },
      layoutComponents: Array.from({ length: 4 }, (_, index) => ({
        id: `component-${index}`,
        name: `布局 ${index}`,
        layout: 'content' as const,
        slots: [{ id: 'image', kind: 'image' as const, x: 1, y: 2, w: 3, h: 4 }],
      })),
    }
    plan.slides[0]!.layoutComponentId = 'component-0'
    const view = await mount({
      phase: 'idle',
      project: { ...pending.project!, plan: { revision: 1, value: plan } },
    })
    const style = view.container.querySelector('[aria-label="样式规范"]')!
    expect(style.querySelectorAll('svg[role="img"]')).toHaveLength(3)
    expect(
      style.querySelector('[aria-label="布局 0布局示意"] rect[x="1"]')?.getAttribute('height'),
    ).toBe('4')
    expect(style.textContent).toContain('品牌色板：#112233')
    expect(style.textContent).toContain('每页使用')
    expect(style.textContent).toContain('非成品预览')
    expect(view.container.textContent).toContain('布局组件：布局 0')
    expect(style.querySelectorAll('img, image')).toHaveLength(0)
  })

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
    expect(view.container.textContent).toContain('附件不可读或来源摘录未匹配')
    expect(view.container.textContent).toContain('修订计划摘录后启动新任务')
    await view.update({
      phase: 'idle',
      project: {
        ...project,
        sourcePreparation: [
          { sourceId: plan.sources[0]!.id, attachmentId, status: 'excerpt_mismatch' },
        ],
      },
    })
    expect(view.container.textContent).toContain('合成基准：摘录不在附件原文中')
    await view.update({
      phase: 'idle',
      project: {
        ...project,
        sourcePreparation: [
          { sourceId: plan.sources[0]!.id, attachmentId, status: 'excerpt_matched' },
        ],
      },
    })
    expect(view.container.textContent).toContain('附件来源 1 / 1 已就绪')
    expect(view.container.textContent).toContain('合成基准：摘录已匹配，编译前仍会复核')
    await view.update({
      phase: 'idle',
      project: {
        ...project,
        sourcePreparation: [{ sourceId: plan.sources[0]!.id, attachmentId, status: 'ready' }],
      },
    })
    expect(view.container.textContent).toContain('合成基准：旧版 PC 仅确认已解析')
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
    expect(view.container.textContent).toContain('页面、版本与检查')
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

function completionSnapshot(state?: string | null): Snapshot {
  return {
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
          { id: 'a', title: 'A', state: 'compiled', attempt: 1 },
          { id: 'b', title: 'B', state: 'pending', attempt: 0 },
        ],
      },
      productionJob:
        state === null
          ? null
          : state === undefined
            ? undefined
            : ({
                version: 1,
                projectId: 'p1',
                documentId: 'd',
                requestId: 'pages',
                inputDigest: 'a'.repeat(64),
                planDigest: 'b'.repeat(64),
                planRevision: 1,
                revision: 1,
                state,
                events: [],
              } as NonNullable<Snapshot['project']>['productionJob']),
    },
  }
}

it.each(['idle', 'loading', 'producing'] as const)(
  'opens completion during %s and ends only frontend while retaining results',
  async (phase) => {
    const end = vi.fn()
    const view = await mount(
      { ...completionSnapshot('running'), phase, error: '暂不可读取' },
      true,
      end,
    )
    const details = view.container.querySelector(
      'details[aria-label="完成操作"]',
    ) as HTMLDetailsElement
    expect(details).not.toBeNull()
    await act(async () => (details.querySelector('summary') as HTMLElement).click())
    expect(details.open).toBe(true)
    const button = view.button('保留成果并结束前台')
    expect(button.disabled).toBe(false)
    await act(async () => button.click())
    expect(end).toHaveBeenCalledOnce()
    expect(view.controller.cancel).toHaveBeenCalledTimes(phase === 'idle' ? 0 : 1)
    expect(view.controller.clear).not.toHaveBeenCalled()
    expect(view.controller.startProductionJob).not.toHaveBeenCalled()
    expect(view.controller.cancelProductionJob).not.toHaveBeenCalled()
    expect(view.container.textContent).toContain('正在提交的修改请核对恢复记录')
  },
)

it('ends frontend waiting for a running job without restarting it even while busy', async () => {
  const end = vi.fn()
  const view = await mount({ ...completionSnapshot('running'), phase: 'producing' }, true, end)
  await act(async () => view.button('继续后台制作并结束前台').click())
  expect(end).toHaveBeenCalledOnce()
  expect(view.controller.cancel).toHaveBeenCalledOnce()
  expect(view.controller.startProductionJob).not.toHaveBeenCalled()
  expect(view.controller.resumeProductionJob).not.toHaveBeenCalled()
})

it.each(['paused', 'interrupted', 'failed'])(
  'resumes exact %s task and ends frontend',
  async (state) => {
    const end = vi.fn()
    const view = await mount(completionSnapshot(state), false, end)
    await act(async () => view.button('继续后台制作并结束前台').click())
    expect(view.controller.resumeProductionJob).toHaveBeenCalledWith('pages')
    expect(view.controller.startProductionJob).not.toHaveBeenCalled()
    expect(end).toHaveBeenCalledOnce()
  },
)

it('starts only a known absent unfinished job', async () => {
  const end = vi.fn()
  const view = await mount(completionSnapshot(null), false, end)
  await act(async () => view.button('继续后台制作并结束前台').click())
  expect(view.controller.startProductionJob).toHaveBeenCalledWith('pages')
  expect(view.controller.resumeProductionJob).not.toHaveBeenCalled()
  expect(end).toHaveBeenCalledOnce()
})

it.each(['pausing', 'cancelling', 'completed', 'cancelled', 'unknown', undefined])(
  'fails closed for non-continuable or unknown %s job',
  async (state) => {
    const view = await mount(completionSnapshot(state), false, vi.fn())
    expect(view.button('继续后台制作并结束前台')).toBeUndefined()
    expect(view.controller.startProductionJob).not.toHaveBeenCalled()
    expect(view.controller.resumeProductionJob).not.toHaveBeenCalled()
  },
)

it.each(['busy', 'applying', 'loading'])(
  'cannot resume or cancel background during %s frontend/project work',
  async (mode) => {
    const view = await mount(
      { ...completionSnapshot('paused'), phase: mode === 'loading' ? 'loading' : 'idle' },
      mode !== 'loading',
      vi.fn(),
    )
    expect(view.button('继续后台制作并结束前台').disabled).toBe(true)
    expect(view.button('取消剩余后台页面').disabled).toBe(true)
    await act(async () => {
      view.button('继续后台制作并结束前台').click()
      view.button('取消剩余后台页面').click()
    })
    expect(view.controller.resumeProductionJob).not.toHaveBeenCalled()
    expect(view.controller.cancelProductionJob).not.toHaveBeenCalled()
  },
)

it('explicit cancellation binds the current task without clearing or restarting it', async () => {
  const view = await mount(completionSnapshot('paused'), false, vi.fn())
  expect(view.container.textContent).toContain('当前编译页结束后生效，已完成成果保留')
  await act(async () => view.button('取消剩余后台页面').click())
  expect(view.controller.cancelProductionJob).toHaveBeenCalledWith('pages')
  expect(view.controller.clear).not.toHaveBeenCalled()
  expect(view.controller.startProductionJob).not.toHaveBeenCalled()
})

it('does not invent continuation when PC lacks support or current job identity differs', async () => {
  const snapshot = completionSnapshot('running')
  const view = await mount(
    { ...snapshot, project: { ...snapshot.project!, jobsUnavailable: true } },
    false,
    vi.fn(),
  )
  expect(view.button('继续后台制作并结束前台')).toBeUndefined()
  expect(view.button('取消剩余后台页面')).toBeUndefined()
  await view.update({
    ...snapshot,
    project: {
      ...snapshot.project!,
      productionJob: { ...snapshot.project!.productionJob!, requestId: 'other' },
    },
  })
  expect(view.button('继续后台制作并结束前台')).toBeUndefined()
  expect(view.button('取消剩余后台页面')).toBeUndefined()
})

it('shows an openable completion menu without a project or optional callback', async () => {
  const view = await mount({ phase: 'loading', error: '读取失败' }, true)
  const details = view.container.querySelector(
    'details[aria-label="完成操作"]',
  ) as HTMLDetailsElement
  await act(async () => (details.querySelector('summary') as HTMLElement).click())
  expect(details.open).toBe(true)
  expect(view.button('保留成果并结束前台').disabled).toBe(true)
  expect(view.button('继续后台制作并结束前台')).toBeUndefined()
  expect(view.container.textContent).toContain('当前没有可确认继续的后台任务')
})

it('fails closed at both completion and existing background entries for unknown or unavailable jobs', async () => {
  const view = await mount(completionSnapshot(undefined), false, vi.fn())
  expect(view.button('后台制作剩余页面')).toBeUndefined()
  await view.update(completionSnapshot('unknown'))
  expect(view.button('取消后台任务')).toBeUndefined()
  const snapshot = completionSnapshot('paused')
  await view.update({ ...snapshot, project: { ...snapshot.project!, jobsUnavailable: true } })
  expect(view.button('继续后台任务')).toBeUndefined()
  expect(view.button('取消后台任务')).toBeUndefined()
})

it('keeps one collapsed page event row across retries and preserves run chronology', async () => {
  const snapshot = completionSnapshot('running')
  const job = snapshot.project!.productionJob!
  job.revision = 4
  job.events = [
    { sequence: 1, createdAt: '2026-09-24T00:00:01.000Z', type: 'run.started' },
    {
      sequence: 2,
      createdAt: '2026-09-24T00:00:02.000Z',
      type: 'page.started',
      pageId: 'b',
      attempt: 1,
    },
    {
      sequence: 3,
      createdAt: '2026-09-24T00:00:03.000Z',
      type: 'page.failed',
      pageId: 'b',
      attempt: 1,
      error: 'compile_failed',
    },
    { sequence: 4, createdAt: '2026-09-24T00:00:04.000Z', type: 'run.paused' },
  ]
  const view = await mount(snapshot)
  const pageRows = () =>
    view.container.querySelectorAll('.presentation-job-events > ol > li[data-page-id]')
  expect(pageRows()).toHaveLength(1)
  const detail = pageRows()[0]!.querySelector('details') as HTMLDetailsElement
  expect(detail.open).toBe(false)
  expect(detail.querySelector('summary')!.textContent).toContain('B')
  expect(detail.querySelector('summary')!.textContent).toContain('第 1 次')
  await act(async () => detail.querySelector('summary')!.click())
  expect(detail.open).toBe(true)
  await view.update({
    ...snapshot,
    project: {
      ...snapshot.project!,
      productionJob: {
        ...job,
        revision: 6,
        events: [
          ...job.events,
          {
            sequence: 5,
            createdAt: '2026-09-24T00:00:05.000Z',
            type: 'page.started',
            pageId: 'b',
            attempt: 2,
          },
          {
            sequence: 6,
            createdAt: '2026-09-24T00:00:06.000Z',
            type: 'page.compiled',
            pageId: 'b',
            attempt: 2,
          },
        ],
      },
    },
  })
  expect(pageRows()).toHaveLength(1)
  const updatedDetail = pageRows()[0]!.querySelector('details') as HTMLDetailsElement
  expect(updatedDetail).toBe(detail)
  expect(updatedDetail.open).toBe(true)
  expect(updatedDetail.querySelector('summary')!.textContent).toContain('第 2 次')
  expect(updatedDetail.querySelector('summary')!.textContent).toContain('未导入验收')
  expect(updatedDetail.querySelectorAll('ol li')).toHaveLength(4)
  expect(
    Array.from(view.container.querySelectorAll('.presentation-job-events > ol > li > time')).map(
      (time) => time.textContent,
    ),
  ).toEqual(['2026-09-24T00:00:01.000Z', '2026-09-24T00:00:04.000Z'])
})
