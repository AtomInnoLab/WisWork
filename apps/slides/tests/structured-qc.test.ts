import { describe, expect, it } from 'vitest'
import type { PlacedBox, RenderNode, RenderSlide, ShapeRenderNode } from '@wiswork/pptx-render'
import {
  auditSlideQuality,
  compareQualityFindings,
  createDeterministicQualityReceipt,
} from '../src/renderer/ai/layout-audit'
import { shouldRunVisualQc, VISUAL_QC_LIMITS } from '../src/renderer/ai/slide-qc'

const box = (x: number, y: number, w: number, h: number): PlacedBox => ({
  x,
  y,
  w,
  h,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
  centerX: x + w / 2,
  centerY: y + h / 2,
})
const text = (
  id: string,
  b: PlacedBox,
  value: string,
  height = 20,
  width = 40,
  size = 20,
): ShapeRenderNode => ({
  id,
  sourceId: id,
  type: 'shape',
  box: b,
  fill: { kind: 'none' },
  text: {
    lines: [
      {
        runs: [
          {
            text: value,
            x: 4,
            baselineY: 15,
            fontFamily: 'Arial',
            fontSizePx: size,
            color: '#000000',
            bold: false,
            italic: false,
            underline: false,
            widthPx: width,
          },
        ],
        top: 0,
        height: 20,
      },
    ],
    insets: { l: 4, t: 4, r: 4, b: 4 },
    anchor: 'top',
    fontScale: 1,
    wrap: true,
    contentHeight: height,
  },
})
const slide = (nodes: RenderNode[]): RenderSlide => ({
  widthPx: 1000,
  heightPx: 600,
  scale: 1,
  background: { kind: 'solid', color: '#FFFFFF' },
  nodes,
})

describe('structured deterministic quality audit', () => {
  it('returns stable identity-only findings for both overflow axes and bounds', () => {
    const findings = auditSlideQuality(
      slide([text('overflow', box(950, -20, 100, 40), 'PRIVATE', 80, 120)]),
      { slideId: 'slide-1' },
    )
    expect(findings.map((f) => f.code)).toEqual([
      'element_off_slide',
      'text_overflow_horizontal',
      'text_overflow_vertical',
    ])
    expect(findings.every((f) => f.slideId === 'slide-1' && f.elementId === 'overflow')).toBe(true)
    expect(JSON.stringify(findings)).not.toContain('PRIVATE')
  })

  it('detects collisions, tiny text, and empty placeholders while excluding backgrounds', () => {
    const bg = {
      ...text('bg', box(0, 0, 1000, 600), 'bg'),
      fill: { kind: 'solid', color: '#000000' },
    } as ShapeRenderNode
    const empty = {
      ...text('empty', box(20, 20, 100, 50), ''),
      type: 'placeholder-chip',
    } as unknown as RenderNode
    const findings = auditSlideQuality(
      slide([
        bg,
        empty,
        text('a', box(100, 100, 300, 100), 'A', 20, 20, 8),
        text('b', box(150, 120, 300, 100), 'B'),
      ]),
      { slideId: 'slide-1' },
    )
    expect(findings.map((f) => f.code)).toEqual(
      expect.arrayContaining(['element_collision', 'tiny_text', 'empty_placeholder']),
    )
    expect(findings.some((f) => f.elementId === 'bg' && f.code === 'element_collision')).toBe(false)
  })

  it('reports contrast only when foreground and solid background are reliably known', () => {
    const low = text('low', box(20, 20, 200, 50), 'Visible')
    low.fill = { kind: 'solid', color: '#FFFFFF' }
    low.text!.lines[0]!.runs[0]!.color = '#EEEEEE'
    expect(auditSlideQuality(slide([low]), { slideId: 'slide-1' }).map((f) => f.code)).toContain(
      'low_contrast',
    )
    low.fill = { kind: 'gradient', angleDeg: 0, stops: [] }
    expect(
      auditSlideQuality(slide([low]), { slideId: 'slide-1' }).map((f) => f.code),
    ).not.toContain('low_contrast')
  })

  it('is deterministic and bounded', () => {
    const nodes = Array.from({ length: 80 }, (_, i) =>
      text(`t${i}`, box(-100, -100, 10, 10), 'x', 40, 40, 6),
    )
    const a = auditSlideQuality(slide(nodes), { slideId: 'slide-1' })
    const b = auditSlideQuality(slide(nodes), { slideId: 'slide-1' })
    expect(a).toEqual(b)
    expect(a.length).toBeLessThanOrEqual(50)
  })

  it('emits a separate receipt with durable creation IDs and comparable stable keys', () => {
    const before = createDeterministicQualityReceipt(
      slide([text('source-1', box(-20, 10, 50, 30), 'x')]),
      {
        qualityRunId: 'qc-1',
        slideId: 'ppt/slides/slide1.xml',
        elementCreationId: () => 'creation-1',
      },
    )
    expect(before.status).toBe('available')
    if (before.status !== 'available') throw new Error('unreachable')
    expect(before.findings[0]).toMatchObject({
      slideId: 'ppt/slides/slide1.xml',
      elementId: 'creation-1',
    })
    expect(compareQualityFindings(before.findings, [])).toEqual({
      introduced: [],
      resolved: ['ppt/slides/slide1.xml:creation-1::element_off_slide'],
    })
  })
})

describe('visual QC policy', () => {
  it('runs only after an applied transaction and only once per transaction/session', () => {
    expect(
      shouldRunVisualQc({
        receiptStatus: 'applied',
        requested: true,
        transactionId: 'tx',
        sessionId: 's',
        completedKeys: new Set(),
      }),
    ).toBe(true)
    expect(
      shouldRunVisualQc({
        receiptStatus: 'conflict',
        requested: true,
        transactionId: 'tx',
        sessionId: 's',
        completedKeys: new Set(),
      }),
    ).toBe(false)
    expect(
      shouldRunVisualQc({
        receiptStatus: 'applied',
        requested: true,
        transactionId: 'tx',
        sessionId: 's',
        completedKeys: new Set(['s:tx']),
      }),
    ).toBe(false)
    expect(VISUAL_QC_LIMITS.maxScreenshots).toBeLessThanOrEqual(20)
    expect(VISUAL_QC_LIMITS.maxScreenshotBytes).toBeLessThanOrEqual(2_000_000)
  })
})
