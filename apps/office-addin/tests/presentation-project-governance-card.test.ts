// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import { PresentationProjectGovernanceCard } from '../src/agent/presentation-project-governance-card'
import { createPresentationProjectGovernanceController } from '../src/agent/presentation-project-governance'
it('requires visible explicit confirmation, describes retained data, and never initializes on rendering', async () => {
  const calls: Record<string, unknown>[] = []
  const controller = createPresentationProjectGovernanceController({
    available: () => true,
    documentId: async () => 'doc',
    currentProjectId: () => 'p',
    readAttempt: () => undefined,
    writeAttempt: () => {},
    request: async (body) => {
      calls.push(body)
      return new Response(
        JSON.stringify({
          preview: {
            expectedRevision: null,
            confirmationToken: 'a'.repeat(64),
            resources: [
              {
                kind: 'master_backups',
                ownership: 'document_shared',
                disposition: 'retained',
                resourceId: 'shared',
                fileCount: 1,
                bytes: 12,
              },
            ],
            governanceRetained: true,
          },
        }),
      )
    },
  })
  const element = document.createElement('div'),
    root = createRoot(element)
  try {
    await act(async () =>
      root.render(React.createElement(PresentationProjectGovernanceCard, { controller })),
    )
    expect(calls).toEqual([])
    expect(element.textContent).toContain('不修改当前 PowerPoint 内容')
    expect(element.textContent).toContain('本机默认不自动清理')
    expect(element.textContent).toContain('PC 明确启用保留期自动执行')
    const button = (text: string) =>
      Array.from(element.querySelectorAll('button')).find((b) => b.textContent === text)!
    await act(async () => button('预览本机项目删除范围').click())
    expect(calls).toHaveLength(1)
    expect(element.textContent).toContain('共享 1 项')
    expect(button('确认删除本机项目资料').disabled).toBe(true)
    expect(button('启用本机项目治理')).toBeDefined()
    await act(async () => element.querySelector<HTMLInputElement>('input[type=checkbox]')!.click())
    expect(button('确认删除本机项目资料').disabled).toBe(false)
    expect(calls).toHaveLength(1)
  } finally {
    await act(async () => root.unmount())
  }
})
it('renders no retained private governance data when unavailable', async () => {
  const controller = createPresentationProjectGovernanceController({
    available: () => false,
    documentId: async () => 'doc',
    currentProjectId: () => 'p',
    readAttempt: () => undefined,
    writeAttempt: () => {},
    request: vi.fn(),
  })
  const element = document.createElement('div'),
    root = createRoot(element)
  try {
    await act(async () =>
      root.render(React.createElement(PresentationProjectGovernanceCard, { controller })),
    )
    expect(element.textContent).toContain('当前连接不支持')
    expect(element.querySelector('button')).toBeNull()
  } finally {
    await act(async () => root.unmount())
  }
})
