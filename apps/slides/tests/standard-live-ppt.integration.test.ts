import { createAgentHarness } from '@wiswork/agent-harness'
import type { AgentTransport } from '@wiswork/agent-core'
import { AI_IPC_LIMITS, streamForProvider } from '@wiswork/ai-provider'
import type { RenderSlide, ShapeRenderNode } from '@wiswork/pptx-render'
import { describe, expect, it } from 'vitest'
import { createSlidesSkill } from '../src/renderer/ai/slides-skill'
import { executePreparedGeometryFamilyTransaction } from '../src/renderer/ai/presentation-geometry-transactions'
import { writePresentationE2eArtifact } from './presentation-e2e-artifact'

const token = process.env.WISWORK_REAL_WISUSAGE_TOKEN
const liveIt = token ? it : it.skip

describe('Standard Slides live presentation workflow', () => {
  liveIt(
    'drives a concrete new-deck request through the production model protocol',
    async () => {
      const blankSlide = (): RenderSlide => ({
        widthPx: 1280,
        heightPx: 720,
        scale: 1,
        background: { kind: 'solid', color: '#FFFFFF' },
        nodes: [],
      })
      let slides = [blankSlide()]
      let activeSlideIndex = 0
      let generated = 0
      const calls: string[] = []
      const transactionIds: string[] = []
      ;(globalThis as unknown as { window: Record<string, unknown> }).window = {
        slidesApi: {
          addSlide: async ({ sourceIndex }: { sourceIndex: number }) => {
            const next = slides.slice()
            const index = sourceIndex + 1
            next.splice(index, 0, blankSlide())
            return { slides: next, index }
          },
        },
      }
      const hostApi: Parameters<typeof executePreparedGeometryFamilyTransaction>[0] = {
        preparePresentationTarget: async (request) => {
          activeSlideIndex = request.slideIndex
          return {
            status: 'prepared' as const,
            expectedDeckRevision: `sha256:${String(transactionIds.length).padStart(64, '0')}`,
            target: {
              slideId: `ppt/slides/slide${request.slideIndex + 1}.xml`,
              expectedFingerprint: `sha256:${String(request.slideIndex + 1).padStart(64, '0')}`,
            },
          }
        },
        cancelPresentationTransaction: async () => true,
        executePresentationTransaction: async (transaction) => {
          const created = new Map<string, string>()
          let nodes = slides[activeSlideIndex]!.nodes.slice()
          for (const rawOperation of transaction.operations) {
            const operation = rawOperation as unknown as Record<string, unknown>
            if (operation.kind === 'add_text_box' || operation.kind === 'add_shape') {
              const sourceId = `generated-${++generated}`
              if (typeof operation.clientId === 'string') created.set(operation.clientId, sourceId)
              const geometry = operation.geometry as {
                x: number
                y: number
                width: number
                height: number
                rotation?: number
              }
              nodes.push({
                id: `render-${sourceId}`,
                sourceId,
                type: 'text',
                box: {
                  x: geometry.x,
                  y: geometry.y,
                  w: geometry.width,
                  h: geometry.height,
                  rotationDeg: geometry.rotation ?? 0,
                  flipH: false,
                  flipV: false,
                  centerX: geometry.x + geometry.width / 2,
                  centerY: geometry.y + geometry.height / 2,
                },
                fill: { kind: 'none' },
                text: {
                  lines: [],
                  insets: { l: 0, t: 0, r: 0, b: 0 },
                  anchor: 'top',
                  fontScale: 1,
                  wrap: true,
                  contentHeight: 0,
                },
              } as unknown as ShapeRenderNode)
            } else if (operation.kind === 'set_text') {
              const target = operation.target as {
                createdByClientId?: string
                elementId?: string
              }
              const sourceId = target.createdByClientId
                ? created.get(target.createdByClientId)
                : target.elementId
              const paragraphs = operation.paragraphs as Array<{
                runs: Array<{ text: string }>
              }>
              nodes = nodes.map((node) =>
                node.sourceId !== sourceId
                  ? node
                  : ({
                      ...node,
                      text: {
                        ...(node as ShapeRenderNode).text!,
                        lines: paragraphs.map((paragraph) => ({
                          runs: paragraph.runs.map((run) => ({
                            text: run.text,
                            x: 0,
                            baselineY: 24,
                            fontFamily: 'Arial',
                            fontSizePx: 24,
                            color: '#172033',
                            bold: false,
                            italic: false,
                            underline: false,
                            widthPx: run.text.length * 12,
                          })),
                          top: 0,
                          height: 28,
                        })),
                      },
                    } as ShapeRenderNode),
              )
            }
          }
          slides = slides.map((slide, index) =>
            index === activeSlideIndex ? { ...slide, nodes } : slide,
          )
          transactionIds.push(transaction.transactionId)
          return {
            status: 'applied' as const,
            transactionId: transaction.transactionId,
            resultingDeckRevision: `sha256:${String(transactionIds.length).padStart(64, '0')}`,
            operationCount: transaction.operations.length,
          }
        },
      }
      const baseSkill = createSlidesSkill({
        getSlides: () => slides,
        getCurrent: () => 0,
        getSelectedIds: () => [],
        applySlide: (index, slide) => {
          slides = slides.map((candidate, candidateIndex) =>
            candidateIndex === index ? slide : candidate,
          )
        },
        applyDeck: (next) => {
          slides = next
        },
        executePresentationOperation: (request, signal) => {
          if (!('operations' in request))
            return Promise.reject(new Error('standard_live_unexpected_transaction_family'))
          return executePreparedGeometryFamilyTransaction(
            hostApi,
            request,
            signal,
            async () => true,
          )
        },
        askClarification: async () => ({ answers: '3 pages; concise blue style; proceed now' }),
        generateStyleSkill: async () => ({
          ok: true,
          styleSkill:
            'Blue minimal business style. Background #F6F8FC, text #172033, accent #2457A7.',
        }),
        planDeckOutline: async () => ({
          ok: true,
          outline: {
            core_hook: '从预测下一个词到理解和生成内容',
            pages: [
              {
                title: '认识大语言模型',
                type: 'cover',
                brief: 'LLM 的一句话定义与核心价值',
                layout: 'cover_dark_minimal',
              },
              {
                title: 'LLM 如何工作',
                type: 'content',
                brief: '训练、上下文和生成过程',
                layout: 'timeline_horizontal',
              },
              {
                title: '能力、风险与使用建议',
                type: 'closing',
                brief: '适用场景、幻觉风险和人机协作',
                layout: 'closing_cta',
              },
            ],
          },
        }),
        fitWidthPx: 1280,
      })
      const skill = {
        ...baseSkill,
        async executeTool(call: Parameters<typeof baseSkill.executeTool>[0], signal?: AbortSignal) {
          calls.push(call.name)
          return baseSkill.executeTool(call, signal)
        },
      }
      const transport: AgentTransport = {
        stream(request, callbacks) {
          const controller = new AbortController()
          void streamForProvider(
            'wiswork',
            { apiKey: '', model: 'openai/gpt-5.6-sol' },
            request.system,
            [...request.messages],
            [...request.tools],
            AI_IPC_LIMITS.maxTokens,
            {
              signal: controller.signal,
              onDelta: callbacks.onDelta,
              onToolCall: callbacks.onToolCall,
              onStopReason: callbacks.onStopReason,
            },
            async (factory) => factory(token!),
          ).then(callbacks.onDone, (error: unknown) =>
            callbacks.onError(error instanceof Error ? error.message : 'standard_live_failed'),
          )
          return { cancel: () => controller.abort() }
        },
      }

      const result = await new Promise<{ text?: string; error?: string }>((resolve) => {
        const harness = createAgentHarness({
          transport,
          skill,
          // The production loop is unbounded, while this live gate is bounded to
          // prove the whole-deck tool avoids a long sequence of blank-page turns.
          maxTurns: 12,
          events: {
            onDone: ({ text }) => resolve({ text }),
            onError: (error) => resolve({ error }),
          },
        })
        expect(
          harness.run('制作一个介绍 LLM 的 PPT。细节请你决定并直接完成，不要停在澄清阶段。'),
        ).toBe(true)
      })

      expect(result.error).toBeUndefined()
      expect(slides.length).toBeGreaterThanOrEqual(3)
      expect(slides.length).toBeLessThanOrEqual(12)
      const pageTexts = slides.map((slide) =>
        slide.nodes
          .flatMap((node) =>
            node.type === 'text' || node.type === 'shape'
              ? (node.text?.lines ?? []).flatMap((line) => line.runs.map((run) => run.text))
              : [],
          )
          .join('\n'),
      )
      const visibleText = pageTexts.join('\n')
      expect(transactionIds.length, `live tool sequence: ${calls.join(',')}`).toBeGreaterThan(0)
      expect(
        pageTexts.every((text) => text.trim().length >= 10),
        `empty generated pages: ${pageTexts
          .map((text, index) => (text.trim() ? '' : index + 1))
          .filter(Boolean)
          .join(',')}`,
      ).toBe(true)
      expect(
        visibleText,
        `live tool sequence: ${calls.join(',')}; transactions: ${transactionIds.length}; slides: ${slides.length}`,
      ).toMatch(/LLM|大语言模型|语言模型/)
      const artifact = await writePresentationE2eArtifact(
        slides,
        process.env.WISWORK_STANDARD_PPT_E2E_OUTPUT ?? '/tmp/wiswork-standard-ppt-e2e.pptx',
      )
      expect(artifact.slideCount).toBe(slides.length)
      expect(artifact.text).toMatch(/LLM|大语言模型|语言模型/)
    },
    180_000,
  )
})
