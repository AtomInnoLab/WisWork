// @vitest-environment jsdom

import React, { act } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { renderToStaticMarkup } from 'react-dom/server'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { PresentationActivityGroup, PresentationEmptyState, PresentationMessage } from '@wiswork/ui'

;(
  globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }
).IS_REACT_ACT_ENVIRONMENT = true

let container: HTMLDivElement
let root: Root

beforeEach(() => {
  container = document.createElement('div')
  document.body.append(container)
  root = createRoot(container)
})

afterEach(() => {
  act(() => root.unmount())
  container.remove()
})

function render(node: React.ReactNode) {
  act(() => root.render(node))
}

describe('shared presentation agent UI', () => {
  it('uses the same quiet message hierarchy in desktop Slides and Office PowerPoint', () => {
    const markup = renderToStaticMarkup(
      React.createElement(
        'div',
        { className: 'ai-chat' },
        React.createElement(PresentationMessage, { role: 'user' }, 'Build a launch deck'),
        React.createElement(PresentationMessage, { role: 'assistant' }, 'I’ll plan and build it.'),
        React.createElement(
          PresentationMessage,
          { role: 'error' },
          'The run could not be completed.',
        ),
      ),
    )

    expect(markup).toContain('class="ai-msg ai-msg-user"')
    expect(markup).toContain('class="ai-msg ai-msg-assistant"')
    expect(markup).toContain('class="ai-msg ai-msg-error"')
    expect(markup).not.toContain('message-role')
  })

  it('opens while work is running and can be inspected after completion', () => {
    render(
      React.createElement(PresentationActivityGroup, {
        items: [{ id: 'context', label: 'Reading presentation', status: 'running' }],
        workingLabel: 'Working',
        workedLabel: (count: number) => `Worked · ${count} step${count === 1 ? '' : 's'}`,
      }),
    )

    const working = container.querySelector<HTMLButtonElement>('.ai-work-group-summary')!
    expect(working.textContent).toContain('Working')
    expect(working.getAttribute('aria-expanded')).toBe('true')
    expect(container.textContent).toContain('Reading presentation')

    render(
      React.createElement(PresentationActivityGroup, {
        items: [{ id: 'context', label: 'Read presentation', status: 'done' }],
        workingLabel: 'Working',
        workedLabel: (count: number) => `Worked · ${count} step${count === 1 ? '' : 's'}`,
      }),
    )

    const worked = container.querySelector<HTMLButtonElement>('.ai-work-group-summary')!
    expect(worked.textContent).toContain('Worked · 1 step')
    expect(worked.getAttribute('aria-expanded')).toBe('false')
    act(() => worked.click())
    expect(worked.getAttribute('aria-expanded')).toBe('true')
    expect(container.textContent).toContain('Read presentation')
  })

  it('renders a PowerPoint-specific, compact empty state without desktop-only chrome', () => {
    const choose = vi.fn()
    render(
      React.createElement(PresentationEmptyState, {
        title: 'Create a presentation',
        body: 'Describe the audience, purpose, or source material.',
        prompts: ['Create a project update', 'Turn this document into slides'],
        onChoose: choose,
      }),
    )

    const prompt = Array.from(container.querySelectorAll('button')).find(
      (button) => button.textContent === 'Create a project update',
    )!
    act(() => prompt.click())
    expect(choose).toHaveBeenCalledWith('Create a project update')
    expect(container.textContent).not.toMatch(/connected|session/i)
  })

  it('renders bounded display-safe activity only', () => {
    render(
      React.createElement(PresentationActivityGroup, {
        items: [
          {
            id: 'proposal',
            label: 'Prepared changes for Slides 1–3',
            status: 'done',
            detail: '3 slides · titles and layouts',
          },
        ],
        workingLabel: 'Working',
        workedLabel: (count: number) => `Worked · ${count} steps`,
      }),
    )
    const detail = container.querySelector<HTMLButtonElement>('.ai-step-title')!
    act(() => detail.click())
    expect(container.textContent).toContain('Prepared changes for Slides 1–3')
    expect(container.textContent).toContain('3 slides · titles and layouts')
    expect(container.textContent).not.toMatch(/callId|fingerprint|capability|arguments/)
  })
})
