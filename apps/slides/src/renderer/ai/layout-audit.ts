import type {
  GroupRenderNode,
  RenderNode,
  RenderSlide,
  ShapeRenderNode,
} from '@wiswork/pptx-render'
import type {
  PresentationQualityFinding,
  PresentationQualityReceipt,
} from '@wiswork/presentation-ops'

/**
 * Deterministic layout audit (modeled on the Google Slides add-in review_google_slides_addin geometry-only checks):
 * pure geometric computation, no LLM calls, no screenshots. Checks three kinds of problems:
 *  1. Elements extending past the canvas
 *  2. Pairwise overlap of content elements (text-text / text-image/media)
 *  3. Text overflowing its text box (uses the render layer's already-laid-out text.contentHeight — exact, not estimated)
 * Results are appended to layout tools' return values so the AI "sees" the real post-edit state (write → verify → fix loop).
 */

interface AuditEntry {
  id: string
  type: string
  x: number
  y: number
  w: number
  h: number
  hasText: boolean
  preview: string
  /** Pixels by which the text content height exceeds the box height (only meaningful when >0) */
  overflowPx: number
  /** Pixels by which the widest laid-out line exceeds the text box inner width */
  overflowXPx: number
  minFontSizePx: number
  emptyPlaceholder: boolean
  contrastRatio: number | null
}

const PREVIEW_MAX = 18

function textPreview(node: ShapeRenderNode): string {
  const t = (node.text?.lines ?? [])
    .map((l) => l.runs.map((r) => r.text).join(''))
    .join(' ')
    .trim()
  return t.length > PREVIEW_MAX ? `${t.slice(0, PREVIEW_MAX)}…` : t
}

/** Collect the top-level nodes that take part in the audit (skip master/layout decoration; a group counts as one box). */
function collectEntries(nodes: RenderNode[]): AuditEntry[] {
  const out: AuditEntry[] = []
  for (const n of nodes) {
    if (n.decoration) continue
    const { x, y, w, h } = n.box
    let hasText = false
    let preview = ''
    let overflowPx = 0
    let overflowXPx = 0
    let minFontSizePx = Infinity
    let contrastRatio: number | null = null
    if (n.type === 'shape' || n.type === 'text') {
      const sn = n as ShapeRenderNode
      preview = textPreview(sn)
      hasText = preview.length > 0
      if (sn.text && hasText) {
        minFontSizePx = sn.text.lines.reduce(
          (minimum, line) =>
            line.runs.reduce((value, run) => Math.min(value, run.fontSizePx), minimum),
          Infinity,
        )
        const inner = h - sn.text.insets.t - sn.text.insets.b
        overflowPx = Math.round(sn.text.contentHeight - inner)
        const innerWidth = w - sn.text.insets.l - sn.text.insets.r
        const widestRight = sn.text.lines.reduce(
          (widest, line) =>
            Math.max(
              widest,
              line.runs.reduce((right, run) => Math.max(right, run.x + run.widthPx), -Infinity),
            ),
          -Infinity,
        )
        overflowXPx = Math.round(Number.isFinite(widestRight) ? widestRight - innerWidth : 0)
        const colors = new Set(sn.text.lines.flatMap((line) => line.runs.map((run) => run.color)))
        const background = sn.fill.kind === 'solid' ? sn.fill.color : null
        if (colors.size === 1 && background)
          contrastRatio = colorContrast([...colors][0]!, background)
      }
    } else if (n.type === 'group') {
      // If any child in the group has text, treat it as text content for overlap detection
      hasText = groupHasText(n as GroupRenderNode)
      preview = '(group)'
    }
    out.push({
      id: n.sourceId,
      type: n.type,
      x,
      y,
      w,
      h,
      hasText,
      preview,
      overflowPx,
      overflowXPx,
      minFontSizePx,
      emptyPlaceholder: n.type === 'placeholder-chip' && !hasText,
      contrastRatio,
    })
  }
  return out
}

function groupHasText(g: GroupRenderNode): boolean {
  for (const c of g.children) {
    if (c.type === 'group') {
      if (groupHasText(c as GroupRenderNode)) return true
    } else if (c.type === 'shape' || c.type === 'text') {
      const t = (c as ShapeRenderNode).text?.lines ?? []
      if (t.some((l) => l.runs.some((r) => r.text.trim()))) return true
    }
  }
  return false
}

const MEDIA_TYPES = new Set(['picture', 'table', 'chart', 'placeholder-chip'])

/** Whether it's a content element (participates in overlap detection): has text, or is a picture/table/chart. */
function isContent(e: AuditEntry): boolean {
  return e.hasText || MEDIA_TYPES.has(e.type)
}

function label(e: AuditEntry): string {
  return e.preview && e.preview !== '(group)' ? `${e.id}"${e.preview}"` : `${e.id}(${e.type})`
}

const EDGE_TOLERANCE_PX = 8
const OVERFLOW_TOLERANCE_PX = 4
/** Threshold for overlap area as a fraction of the smaller element's area */
const OVERLAP_RATIO = 0.12
/** Absolute overlap area floor (px²), filtering out noise like touching trims */
const OVERLAP_MIN_AREA = 400
/** Background color blocks (≥70% of canvas area) don't participate in overlap detection */
const BACKGROUND_AREA_RATIO = 0.7
const MAX_ISSUES = 12
export const MAX_QUALITY_FINDINGS = 50
export const MAX_QUALITY_ELEMENTS = 500
export const MAX_COLLISION_CANDIDATES = 5_000

function colorContrast(foreground: string, background: string): number | null {
  const parse = (value: string) => {
    const match = /^#([0-9a-f]{6})$/i.exec(value)
    if (!match) return null
    return [0, 2, 4].map((offset) => Number.parseInt(match[1]!.slice(offset, offset + 2), 16) / 255)
  }
  const fg = parse(foreground),
    bg = parse(background)
  if (!fg || !bg) return null
  const luminance = (rgb: number[]) =>
    0.2126 * channel(rgb[0]!) + 0.7152 * channel(rgb[1]!) + 0.0722 * channel(rgb[2]!)
  const channel = (value: number) =>
    value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4
  const [lighter, darker] = [luminance(fg), luminance(bg)].sort((a, b) => b - a)
  return (lighter! + 0.05) / (darker! + 0.05)
}

/** Pure, bounded and content-free QC contract for transaction quality receipts. */
export function auditSlideQuality(
  slide: RenderSlide,
  identity: {
    slideId: string
    elementCreationId?: (sourceId: string) => string | undefined
    stats?: { candidatePairs: number; elements: number }
  },
): PresentationQualityFinding[] {
  if (identity.stats) {
    identity.stats.elements = Math.min(slide.nodes.length, MAX_QUALITY_ELEMENTS)
    identity.stats.candidatePairs = 0
  }
  if (slide.nodes.length > MAX_QUALITY_ELEMENTS) {
    return [
      {
        code: 'quality_capacity_exceeded',
        severity: 'critical',
        slideId: identity.slideId,
        evidence: {
          elementCount: Math.min(slide.nodes.length, 1_000_000),
          elementLimit: MAX_QUALITY_ELEMENTS,
        },
      },
    ]
  }
  const entries = collectEntries(slide.nodes)
  const findings: PresentationQualityFinding[] = []
  const add = (finding: PresentationQualityFinding) => {
    if (findings.length >= MAX_QUALITY_FINDINGS) return
    const elementId = finding.elementId
      ? identity.elementCreationId?.(finding.elementId)
      : undefined
    const relatedElementId = finding.relatedElementId
      ? identity.elementCreationId?.(finding.relatedElementId)
      : undefined
    const durableFinding = {
      ...finding,
    }
    delete durableFinding.elementId
    delete durableFinding.relatedElementId
    findings.push({
      ...durableFinding,
      ...(elementId ? { elementId } : {}),
      ...(relatedElementId ? { relatedElementId } : {}),
    })
  }
  for (const e of entries) {
    const leftPx = Math.max(0, -e.x)
    const topPx = Math.max(0, -e.y)
    const rightPx = Math.max(0, e.x + e.w - slide.widthPx)
    const bottomPx = Math.max(0, e.y + e.h - slide.heightPx)
    if (Math.max(leftPx, topPx, rightPx, bottomPx) > EDGE_TOLERANCE_PX)
      add({
        code: 'element_off_slide',
        severity: 'important',
        slideId: identity.slideId,
        elementId: e.id,
        evidence: {
          leftPx: Math.round(leftPx),
          topPx: Math.round(topPx),
          rightPx: Math.round(rightPx),
          bottomPx: Math.round(bottomPx),
        },
      })
    if (e.overflowXPx > OVERFLOW_TOLERANCE_PX)
      add({
        code: 'text_overflow_horizontal',
        severity: 'important',
        slideId: identity.slideId,
        elementId: e.id,
        evidence: { overflowPx: e.overflowXPx },
      })
    if (e.overflowPx > OVERFLOW_TOLERANCE_PX)
      add({
        code: 'text_overflow_vertical',
        severity: 'important',
        slideId: identity.slideId,
        elementId: e.id,
        evidence: { overflowPx: e.overflowPx },
      })
    if (e.minFontSizePx < 10)
      add({
        code: 'tiny_text',
        severity: 'warning',
        slideId: identity.slideId,
        elementId: e.id,
        evidence: { fontSizePx: e.minFontSizePx },
      })
    if (e.emptyPlaceholder)
      add({
        code: 'empty_placeholder',
        severity: 'warning',
        slideId: identity.slideId,
        elementId: e.id,
        evidence: {},
      })
    if (e.contrastRatio !== null && e.contrastRatio < 4.5)
      add({
        code: 'low_contrast',
        severity: 'important',
        slideId: identity.slideId,
        elementId: e.id,
        evidence: { contrastRatioMilli: Math.round(e.contrastRatio * 1000) },
      })
  }
  const content = entries.filter(
    (e) => isContent(e) && e.w * e.h < slide.widthPx * slide.heightPx * BACKGROUND_AREA_RATIO,
  )
  const ordered = [...content].sort(
    (left, right) => left.x - right.x || left.id.localeCompare(right.id),
  )
  let capacityExceeded = false
  let candidatePairs = 0
  for (let i = 0; i < ordered.length && findings.length < MAX_QUALITY_FINDINGS; i++) {
    const a = ordered[i]!
    for (let j = i + 1; j < ordered.length && findings.length < MAX_QUALITY_FINDINGS; j++) {
      const b = ordered[j]!
      if (b.x >= a.x + a.w) break
      candidatePairs += 1
      if (identity.stats)
        identity.stats.candidatePairs = Math.min(candidatePairs, MAX_COLLISION_CANDIDATES)
      if (candidatePairs > MAX_COLLISION_CANDIDATES) {
        capacityExceeded = true
        break
      }
      if (!a.hasText && !b.hasText) continue
      const ix = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x)
      const iy = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y)
      if (ix <= 0 || iy <= 0) continue
      const overlapAreaPx = ix * iy
      const ratio = overlapAreaPx / Math.min(a.w * a.h, b.w * b.h)
      if (overlapAreaPx < OVERLAP_MIN_AREA || ratio < OVERLAP_RATIO) continue
      const [elementId, relatedElementId] = [a.id, b.id].sort()
      add({
        code: 'element_collision',
        severity: 'important',
        slideId: identity.slideId,
        elementId,
        relatedElementId,
        evidence: {
          overlapAreaPx: Math.round(overlapAreaPx),
          overlapRatioPermille: Math.round(ratio * 1000),
        },
      })
    }
    if (capacityExceeded) break
  }
  if (capacityExceeded)
    return [
      {
        code: 'quality_capacity_exceeded',
        severity: 'critical',
        slideId: identity.slideId,
        evidence: { candidateLimit: MAX_COLLISION_CANDIDATES },
      },
    ]
  return findings
}

export function createDeterministicQualityReceipt(
  slide: RenderSlide,
  identity: {
    qualityRunId: string
    transactionId: string
    slideId: string
    elementCreationId?: (sourceId: string) => string | undefined
  },
): PresentationQualityReceipt {
  const findings = auditSlideQuality(slide, identity)
  return {
    qualityRunId: identity.qualityRunId,
    transactionId: identity.transactionId,
    slideId: identity.slideId,
    source: 'deterministic',
    status: 'available',
    findings,
    truncated: findings.length === MAX_QUALITY_FINDINGS,
  }
}

export function qualityFindingKey(finding: PresentationQualityFinding): string {
  return [
    finding.slideId,
    finding.elementId ?? '',
    finding.relatedElementId ?? '',
    finding.code,
  ].join(':')
}

export function compareQualityFindings(
  before: readonly PresentationQualityFinding[],
  after: readonly PresentationQualityFinding[],
): { introduced: string[]; resolved: string[] } {
  const beforeKeys = new Set(before.map(qualityFindingKey))
  const afterKeys = new Set(after.map(qualityFindingKey))
  return {
    introduced: [...afterKeys].filter((key) => !beforeKeys.has(key)).sort(),
    resolved: [...beforeKeys].filter((key) => !afterKeys.has(key)).sort(),
  }
}

/**
 * Audit one page's layout and return the list of problems (empty array = pass).
 */
export function auditSlideLayout(slide: RenderSlide): string[] {
  const entries = collectEntries(slide.nodes)
  const issues: string[] = []
  const W = slide.widthPx
  const H = slide.heightPx

  // 1. Out of bounds
  for (const e of entries) {
    const parts: string[] = []
    if (e.x < -EDGE_TOLERANCE_PX) parts.push(`${Math.round(-e.x)}px past the left edge`)
    if (e.y < -EDGE_TOLERANCE_PX) parts.push(`${Math.round(-e.y)}px past the top edge`)
    if (e.x + e.w > W + EDGE_TOLERANCE_PX)
      parts.push(`${Math.round(e.x + e.w - W)}px past the right edge`)
    if (e.y + e.h > H + EDGE_TOLERANCE_PX)
      parts.push(`${Math.round(e.y + e.h - H)}px past the bottom edge`)
    if (parts.length) issues.push(`Out of bounds: ${label(e)} ${parts.join(', ')}`)
  }

  for (const e of entries) {
    if (e.overflowXPx > OVERFLOW_TOLERANCE_PX) {
      issues.push(
        `Text overflow (width): ${label(e)} content exceeds the box width by ${e.overflowXPx}px (widen the box, reduce the font size, or turn on wrapping)`,
      )
    }
  }

  // 2. Text overflow
  for (const e of entries) {
    if (e.overflowPx > OVERFLOW_TOLERANCE_PX) {
      issues.push(
        `Text overflow: ${label(e)} content exceeds the box height by ${e.overflowPx}px (make the box taller or reduce the font size)`,
      )
    }
  }

  // 3. Pairwise overlap of content elements
  const content = entries.filter((e) => isContent(e) && e.w * e.h < W * H * BACKGROUND_AREA_RATIO)
  for (let i = 0; i < content.length; i++) {
    for (let j = i + 1; j < content.length; j++) {
      const a = content[i]!
      const b = content[j]!
      // Only report text<->text and text<->media; media-on-media (e.g. a chart on an image) is often intentional design, don't report
      if (!a.hasText && !b.hasText) continue
      const ix = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x)
      const iy = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y)
      if (ix <= 0 || iy <= 0) continue
      const inter = ix * iy
      const minArea = Math.min(a.w * a.h, b.w * b.h)
      if (inter < OVERLAP_MIN_AREA || inter < minArea * OVERLAP_RATIO) continue
      issues.push(
        `Overlap: ${label(a)} and ${label(b)} intersect by ${Math.round(ix)}×${Math.round(iy)}px`,
      )
      if (issues.length >= MAX_ISSUES) break
    }
    if (issues.length >= MAX_ISSUES) break
  }

  return issues.slice(0, MAX_ISSUES)
}

/** Format the audit result as trailing text for a tool's return value. */
export function formatAudit(issues: string[], round?: string): string {
  if (issues.length === 0)
    return '\n<layout-audit>✅ Passed: no overlap/out-of-bounds/text overflow.</layout-audit>'
  const head = `\n<layout-audit>⚠️ Found ${issues.length} issue(s):\n`
  const body = issues.map((s) => `- ${s}`).join('\n')
  const tail = round
    ? `\n${round}\n</layout-audit>`
    : "\n→ Immediately write another execute_slide_script to fix these issues (don't stop, don't ask the user, don't declare completion). els reflects the new positions after the last apply; compute from it directly. At most 2 fix rounds; only if still unresolved tell the user honestly.\n</layout-audit>"
  return head + body + tail
}
