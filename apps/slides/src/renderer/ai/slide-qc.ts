/**
 * Read-only post-transaction quality review. A fresh AgentLoop receives one bounded
 * screenshot plus geometry-only context and cannot invoke mutation tools.
 */
import {
  AgentLoop,
  type AgentImage,
  type AgentSkill,
  type AgentTransport,
} from '@wiswork/agent-core'
import type { PresentationQualityReceipt } from '@wiswork/presentation-ops'
import { auditSlideQuality } from './layout-audit'
import type { DeckAccess } from './slides-skill'

/** Kill switch: localStorage 'ai-slides-qc' = '0' disables the automatic pass */
export function isQcEnabled(): boolean {
  return localStorage.getItem('ai-slides-qc') !== '0'
}

/** Cost ceiling per generation run — beyond this the tail pages are skipped (reported to the user) */
export const QC_MAX_PAGES = 20

export const VISUAL_QC_LIMITS = Object.freeze({
  maxScreenshots: QC_MAX_PAGES,
  maxScreenshotBytes: 2_000_000,
  maxSummaryElements: 100,
  maxPromptChars: 12_000,
})

export function shouldRunVisualQc(opts: {
  receiptStatus: string
  requested: boolean
  transactionId: string
  sessionId: string
  completedKeys: ReadonlySet<string>
}): boolean {
  return (
    opts.requested &&
    opts.receiptStatus === 'applied' &&
    !opts.completedKeys.has(`${opts.sessionId}:${opts.transactionId}`)
  )
}

/**
 * Pages produced by a HTML conversion call, as 0-based indexes.
 * replace lands a whole new deck; append starts at appendedFrom; insert_at/replace_at touch one page.
 */
export function generatedPageRange(
  mode: 'replace' | 'append' | 'insert_at' | 'replace_at',
  r: { pages?: number; appendedFrom?: number; insertedIndex?: number },
): number[] {
  const total = r.pages ?? 0
  switch (mode) {
    case 'replace':
      return Array.from({ length: total }, (_, i) => i)
    case 'append': {
      const from = r.appendedFrom ?? 0
      return Array.from({ length: Math.max(0, total - from) }, (_, i) => from + i)
    }
    case 'insert_at':
    case 'replace_at':
      return typeof r.insertedIndex === 'number' ? [r.insertedIndex] : []
  }
}

/**
 * Fold one landing's pages into the run's pending-QC set.
 * replace discards earlier pendings (whole new deck); insert_at shifts pendings at/after
 * the insertion point before adding it, keeping indexes valid.
 */
export function mergeQcPages(
  prev: number[],
  mode: 'replace' | 'append' | 'insert_at' | 'replace_at',
  r: { pages?: number; appendedFrom?: number; insertedIndex?: number },
): number[] {
  const range = generatedPageRange(mode, r)
  const sortDedupe = (pages: number[]) => [...new Set(pages)].sort((a, b) => a - b)
  switch (mode) {
    case 'replace':
      return range
    case 'append':
    case 'replace_at':
      return sortDedupe([...prev, ...range])
    case 'insert_at': {
      const at = r.insertedIndex
      if (typeof at !== 'number') return prev
      return sortDedupe([...prev.map((p) => (p >= at ? p + 1 : p)), at])
    }
  }
}

/** Only the two tools the QC pass needs: fresh geometry reads + atomic layout scripts */
const QC_SYSTEM_PROMPT = `You are a read-only slide quality reviewer. Each request gives you ONE bounded screenshot and a geometry-only element summary.

Look at the screenshot for OBJECTIVE layout defects only:
- text overflowing its box, colliding with a neighbor, or clipped by the canvas edge
- elements overlapping unintentionally (a text block over another text block; content under an image)
- unreadable contrast (text color too close to what it sits on)
- obviously ragged alignment or wildly uneven spacing among sibling items (cards, bullets, columns)
- distorted or badly cropped images

Do not edit, invoke tools, quote slide text, or propose an automatic follow-up edit.

Final reply: one short line (under 15 words) stating the quality warning, or exactly "OK" if clean.`

export interface QcPageResult {
  /** page still exists and the pass ran */
  ok: boolean
  /** at least one mutating tool call was applied */
  edited: boolean
  /** model's final one-liner ('OK' when clean) */
  reply: string
  /** deterministic audit issue counts before/after (rollback signal: after > before) */
  preIssues: number
  postIssues: number
  error?: string
}

export function toVisualQualityReceipt(
  qualityRunId: string,
  slideId: string,
  result: QcPageResult,
): PresentationQualityReceipt {
  if (result.error) {
    if (result.error === 'cancelled' || result.error === 'stale_session') {
      return { qualityRunId, source: 'visual', status: 'cancelled', code: result.error }
    }
    return {
      qualityRunId,
      source: 'visual',
      status: 'unavailable',
      code:
        result.error === 'screenshot_unavailable'
          ? 'screenshot_unavailable'
          : 'transport_unavailable',
    }
  }
  return {
    qualityRunId,
    source: 'visual',
    status: 'available',
    findings:
      result.reply && result.reply.toUpperCase() !== 'OK'
        ? [{ code: 'visual_quality', severity: 'warning', slideId, evidence: {} }]
        : [],
    truncated: false,
  }
}

export interface QcPageOptions {
  access: DeckAccess
  transport: AgentTransport
  pageIndex: number
  /** pixelRatio-1 PNG of the page's current rendering; null runs a geometry-only pass */
  screenshot: AgentImage | null
  systemSuffix?: () => string
  signal?: AbortSignal
  /** Session/deck identity guard checked again when the asynchronous review settles. */
  isCurrent?: () => boolean
}

export async function captureCurrentQcShot<T>(opts: {
  capture: () => Promise<T>
  signal: AbortSignal
  isCurrent: () => boolean
}): Promise<{ value: T } | null> {
  const shot = await opts.capture()
  opts.signal.throwIfAborted()
  return opts.isCurrent() ? { value: shot } : null
}

/** Read-only skill: visual QC cannot enter the mutation executor boundary. */
export function createSlideFixSkill(access: DeckAccess): AgentSkill {
  void access
  return {
    id: 'slides-qc',
    systemPrompt: QC_SYSTEM_PROMPT,
    tools: [],
    executeTool: () => ({
      output: 'quality_read_only',
      summary: 'Quality review is read-only',
      isError: true,
      mutated: false,
    }),
  }
}

function buildQcInstruction(
  pageIndex: number,
  slide: ReturnType<DeckAccess['getSlides']>[number],
  issues: ReturnType<typeof auditSlideQuality>,
): string {
  const auditStr = issues.length
    ? `Deterministic geometry audit codes:\n${JSON.stringify(issues)}\n(These are bounded geometry hints.)`
    : 'The deterministic geometry audit found nothing — trust the screenshot for visual defects it cannot measure (contrast, alignment, crowding).'
  const entries = slide.nodes.slice(0, VISUAL_QC_LIMITS.maxSummaryElements).map((node) => ({
    id: node.sourceId,
    type: node.type,
    x: Math.round(node.box.x),
    y: Math.round(node.box.y),
    w: Math.round(node.box.w),
    h: Math.round(node.box.h),
  }))
  return `Slide ${pageIndex + 1} (slideIndex ${pageIndex}) was just applied. The attached image is its current rendering.

Geometry-only inventory (${entries.length} elements; capped):
${JSON.stringify(entries)}

${auditStr}

Inspect the screenshot and report objective layout defects. Do not edit.`.slice(
    0,
    VISUAL_QC_LIMITS.maxPromptChars,
  )
}

/**
 * One page, one focused read-only QC run. Quality failure never changes write status.
 */
export function qcSlidePage(opts: QcPageOptions): Promise<QcPageResult> {
  const { access, transport, pageIndex, screenshot, systemSuffix, signal, isCurrent } = opts
  signal?.throwIfAborted()
  const slide = access.getSlides()[pageIndex]
  if (!slide) {
    return Promise.resolve({
      ok: false,
      edited: false,
      reply: '',
      preIssues: 0,
      postIssues: 0,
      error: `slideIndex ${pageIndex} out of range`,
    })
  }
  const preIssues = auditSlideQuality(slide, { slideId: `slide-${pageIndex + 1}` })
  if (
    screenshot &&
    Math.floor((screenshot.base64.length * 3) / 4) > VISUAL_QC_LIMITS.maxScreenshotBytes
  ) {
    return Promise.resolve({
      ok: false,
      edited: false,
      reply: '',
      preIssues: preIssues.length,
      postIssues: preIssues.length,
      error: 'screenshot_unavailable',
    })
  }
  const instruction = buildQcInstruction(pageIndex, slide, preIssues)

  return new Promise((resolve) => {
    let settled = false
    const finish = (r: { reply: string; error?: string }) => {
      if (settled) return
      settled = true
      signal?.removeEventListener('abort', onAbort)
      if (isCurrent && !isCurrent()) {
        resolve({
          ok: false,
          edited: false,
          reply: '',
          preIssues: preIssues.length,
          postIssues: preIssues.length,
          error: 'stale_session',
        })
        return
      }
      const after = access.getSlides()[pageIndex]
      resolve({
        ok: true,
        edited: false,
        reply: r.reply.trim().replace(/\s+/g, ' ').slice(0, 240),
        preIssues: preIssues.length,
        postIssues: after
          ? auditSlideQuality(after, { slideId: `slide-${pageIndex + 1}` }).length
          : 0,
        ...(r.error !== undefined ? { error: r.error } : {}),
      })
    }
    const loop = new AgentLoop({
      transport,
      skill: createSlideFixSkill(access),
      maxTurns: 1,
      ...(systemSuffix ? { systemSuffix } : {}),
      events: {
        onToolExecuted: () => {},
        onDone: ({ text, cancelled }) =>
          finish({ reply: text, ...(cancelled ? { error: 'cancelled' } : {}) }),
        onError: (error) => finish({ reply: '', error }),
      },
    })
    const onAbort = () => loop.cancel()
    signal?.addEventListener('abort', onAbort, { once: true })
    signal?.throwIfAborted()
    loop.run(instruction, screenshot ? [screenshot] : [])
  })
}
