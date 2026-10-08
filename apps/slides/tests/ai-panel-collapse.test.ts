// Slides: the AI panel stays mounted while collapsed (rail only),
// so the conversation, draft, and in-flight runs survive collapse/expand.
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

// react-konva's node entry requires the native 'canvas' package; nothing here draws
vi.mock('react-konva', () => {
  const stub = () => null
  return {
    Stage: stub,
    Layer: stub,
    Rect: stub,
    Group: stub,
    Transformer: stub,
    Line: stub,
    Arrow: stub,
    Text: stub,
    Ellipse: stub,
    Image: stub,
    Path: stub,
    Circle: stub,
    Arc: stub,
  }
})

import { AiPanel } from '../src/renderer/ai/AiPanel'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((p) => [p.id, { apiKey: '', model: p.defaultModel }]),
  ) as AiSettings['providers'],
}

function mount(element: React.ReactElement): {
  container: HTMLElement
  root: Root
  cleanup: () => void
} {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  return {
    container,
    root,
    cleanup: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

function panelProps(overrides: Record<string, unknown> = {}) {
  return {
    slides: [],
    current: 0,
    selectedIds: [],
    images: new Map<string, HTMLImageElement>(),
    applySlide: () => {},
    applyDeck: () => {},
    fitWidthPx: 960,
    settings,
    open: true,
    onExpand: () => {},
    onCollapse: () => {},
    ...overrides,
  }
}

/** Simulate typing into React's controlled textarea */
function typeInto(textarea: HTMLTextAreaElement, text: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
  act(() => {
    setter.call(textarea, text)
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

beforeAll(() => {
  // jsdom has no scrollTo; the panel auto-scrolls its chat log
  Element.prototype.scrollTo ??= () => {}
})

describe('AiPanel collapse (slides)', () => {
  it('keeps the draft input across a collapse/expand cycle', () => {
    const { container, root, cleanup } = mount(createElement(AiPanel, panelProps()))

    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')
    expect(textarea).not.toBeNull()
    typeInto(textarea!, 'unsent draft')
    expect(textarea!.value).toBe('unsent draft')

    // collapse: only the rail is rendered, but the component stays mounted
    act(() => root.render(createElement(AiPanel, panelProps({ open: false }))))
    expect(container.querySelector('.ai-input-box textarea')).toBeNull()
    expect(container.querySelector('.ai-rail')).not.toBeNull()

    // expand: the draft is still there
    act(() => root.render(createElement(AiPanel, panelProps({ open: true }))))
    const restored = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')
    expect(restored).not.toBeNull()
    expect(restored!.value).toBe('unsent draft')

    cleanup()
  })

  it('expands back through the rail button', () => {
    const onExpand = vi.fn()
    const { container, cleanup } = mount(
      createElement(AiPanel, panelProps({ open: false, onExpand })),
    )

    const rail = container.querySelector<HTMLButtonElement>('.ai-rail')
    expect(rail).not.toBeNull()
    act(() => rail!.click())
    expect(onExpand).toHaveBeenCalledTimes(1)

    cleanup()
  })

  it('shows the stop action while a selection-scoped run is busy', async () => {
    const originalApi = window.slidesApi
    let cleanup: () => void = () => undefined
    try {
      window.slidesApi = {
        getDesignSidecar: vi.fn(async () => ({ ok: true })),
        captureAgentSelection: vi.fn(async () => ({
          status: 'captured' as const,
          documentId: 'document-1',
          sessionId: 'session-1',
          generation: 1,
          slides: [
            {
              slideId: 'slide-1',
              elements: [
                {
                  elementId: 'shape-1',
                  expectedType: 'text' as const,
                  expectedFingerprint: 'fingerprint-1',
                },
              ],
            },
          ],
        })),
        beginHistoryBatch: vi.fn(async () => true),
        endHistoryBatch: vi.fn(async () => 1),
        onAiStream: vi.fn(() => () => undefined),
        aiStream: vi.fn(() => new Promise<void>(() => undefined)),
        aiStreamCancel: vi.fn(async () => undefined),
      } as unknown as typeof window.slidesApi
      const mounted = mount(createElement(AiPanel, panelProps({ selectedIds: ['shape-1'] })))
      cleanup = mounted.cleanup
      const scope = mounted.container.querySelector<HTMLInputElement>(
        '.ai-selection-scope-toggle input',
      )!
      act(() => scope.click())
      expect(scope.checked).toBe(true)
      const textarea =
        mounted.container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
      typeInto(textarea, 'Edit this shape')
      act(() => mounted.container.querySelector<HTMLButtonElement>('.ai-send-btn')!.click())
      await vi.waitFor(() => expect(mounted.container.querySelector('.ai-stop-btn')).not.toBeNull())
    } finally {
      cleanup()
      ;(window as typeof window & { slidesApi?: unknown }).slidesApi = originalApi
    }
  })
})
