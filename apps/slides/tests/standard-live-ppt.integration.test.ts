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
            core_hook: '入职第一天，快速进入状态',
            pages: [
              {
                title: '欢迎加入',
                type: 'cover',
                brief: '新人入职欢迎',
                layout: 'cover_dark_minimal',
              },
              {
                title: '第一天安排',
                type: 'content',
                brief: '上午下午时间线',
                layout: 'timeline_horizontal',
              },
              {
                title: '协作方式',
                type: 'closing',
                brief: '沟通渠道与行动',
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
          maxTurns: 8,
          events: {
            onDone: ({ text }) => resolve({ text }),
            onError: (error) => resolve({ error }),
          },
        })
        expect(
          harness.run(
            '创建一份3页新人入职培训PPT：简洁蓝色风格，包含欢迎、第一天安排、协作方式。需求明确，请直接制作，不要再询问。',
          ),
        ).toBe(true)
      })

      expect(result.error).toBeUndefined()
      expect(slides).toHaveLength(3)
      const visibleText = slides
        .flatMap((slide) => slide.nodes)
        .flatMap((node) =>
          node.type === 'text' || node.type === 'shape'
            ? (node.text?.lines ?? []).flatMap((line) => line.runs.map((run) => run.text))
            : [],
        )
        .join('\n')
      expect(visibleText).toMatch(/新人|入职|欢迎/)
      expect(transactionIds.length, `live tool sequence: ${calls.join(',')}`).toBeGreaterThan(0)
      const artifact = await writePresentationE2eArtifact(
        slides,
        process.env.WISWORK_STANDARD_PPT_E2E_OUTPUT ?? '/tmp/wiswork-standard-ppt-e2e.pptx',
      )
      expect(artifact.slideCount).toBe(3)
      expect(artifact.text).toMatch(/新人|入职|欢迎/)
    },
    120_000,
  )
})
