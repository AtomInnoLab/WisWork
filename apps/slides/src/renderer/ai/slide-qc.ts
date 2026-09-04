/**
 * Read-only post-transaction quality review. A fresh AgentLoop receives one bounded
 * screenshot plus geometry-only context and cannot invoke mutation tools.
 */
import {
  AgentLoop,
  type AgentImage,
  type AgentSkill,
  type AgentStreamRequest,
  type AgentTransport,
} from '@wiswork/agent-core'
import type { PresentationQualityReceipt } from '@wiswork/presentation-ops'
import { auditSlideQuality, createDeterministicQualityReceipt } from './layout-audit'
import type { DeckAccess } from './slides-skill'
import { geometryPxToPoints } from './presentation-geometry-transactions'

/** Kill switch: localStorage 'ai-slides-qc' = '0' disables the automatic pass */
export function isQcEnabled(): boolean {
  return localStorage.getItem('ai-slides-qc') !== '0'
}

/** Cost ceiling per generation run — beyond this the tail pages are skipped (reported to the user) */
export const QC_MAX_PAGES = 20
export const DETERMINISTIC_QC_MAX_PAGES = 50

export const VISUAL_QC_LIMITS = Object.freeze({
  maxScreenshots: QC_MAX_PAGES,
  maxTransportRequestBytes: 2 * 1024 * 1024,
  /** Leaves 50 KB for the bounded prompt, system message, and provider JSON envelope. */
  maxScreenshotRequestBytes: 1_950_000,
  maxSummaryElements: 100,
  maxPromptChars: 12_000,
})

export interface QcGeometryFix {
  sourceId: string
  x: number
  y: number
  width: number
  height: number
}

export interface QcReview {
  status: 'pass' | 'needs_fix' | 'cannot_verify'
  summary: string
  fixes: QcGeometryFix[]
}

const QC_MAX_FIXES = 8

export function parseQcReview(
  text: string,
  slide: ReturnType<DeckAccess['getSlides']>[number],
): QcReview {
  if (text.trim().toUpperCase() === 'OK') return { status: 'pass', summary: 'OK', fixes: [] }
  const start = text.indexOf('{')
  const end = text.lastIndexOf('}')
  if (start < 0 || end <= start) throw new Error('quality_review_invalid')
  const value = JSON.parse(text.slice(start, end + 1)) as unknown
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('quality_review_invalid')
  const record = value as Record<string, unknown>
  if (
    Object.keys(record).some((key) => !['status', 'summary', 'fixes'].includes(key)) ||
    !['pass', 'needs_fix', 'cannot_verify'].includes(String(record.status)) ||
    typeof record.summary !== 'string' ||
    record.summary.length > 240 ||
    !Array.isArray(record.fixes) ||
    record.fixes.length > QC_MAX_FIXES
  )
    throw new Error('quality_review_invalid')
  const allowed = new Set(
    slide.nodes.filter((node) => !node.decoration).map((node) => node.sourceId),
  )
  const fixes = record.fixes.map((raw) => {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw))
      throw new Error('quality_review_invalid')
    const fix = raw as Record<string, unknown>
    if (
      Object.keys(fix).some((key) => !['sourceId', 'x', 'y', 'width', 'height'].includes(key)) ||
      typeof fix.sourceId !== 'string' ||
      !allowed.has(fix.sourceId)
    )
      throw new Error('quality_review_invalid')
    const numbers = [fix.x, fix.y, fix.width, fix.height]
    if (numbers.some((item) => typeof item !== 'number' || !Number.isFinite(item)))
      throw new Error('quality_review_invalid')
    const [x, y, width, height] = numbers as number[]
    if (
      x < 0 ||
      y < 0 ||
      width < 8 ||
      height < 8 ||
      x + width > slide.widthPx ||
      y + height > slide.heightPx
    )
      throw new Error('quality_review_invalid')
    return { sourceId: fix.sourceId, x, y, width, height }
  })
  if ((record.status === 'pass' || record.status === 'cannot_verify') && fixes.length)
    throw new Error('quality_review_invalid')
  if (record.status === 'needs_fix' && fixes.length === 0) throw new Error('quality_review_invalid')
  return {
    status: record.status as QcReview['status'],
    summary: record.summary.trim().replace(/\s+/g, ' '),
    fixes,
  }
}

export async function applyQcGeometryFixes(
  access: DeckAccess,
  pageIndex: number,
  fixes: readonly QcGeometryFix[],
  signal?: AbortSignal,
): Promise<boolean> {
  if (!access.executePresentationOperation || fixes.length === 0 || fixes.length > QC_MAX_FIXES)
    return false
  const slide = access.getSlides()[pageIndex]
  if (!slide) return false
  const scale = slide.scale || access.fitWidthPx / slide.widthPx
  if (!Number.isFinite(scale) || scale <= 0) return false
  const allowed = new Set(
    slide.nodes.filter((node) => !node.decoration).map((node) => node.sourceId),
  )
  if (
    fixes.some(
      (fix) =>
        !allowed.has(fix.sourceId) ||
        ![fix.x, fix.y, fix.width, fix.height].every(Number.isFinite) ||
        fix.x < 0 ||
        fix.y < 0 ||
        fix.width < 8 ||
        fix.height < 8 ||
        fix.x + fix.width > slide.widthPx ||
        fix.y + fix.height > slide.heightPx,
    )
  )
    return false
  const execution = await access.executePresentationOperation(
    {
      transactionId: `slides-qc-fix-${crypto.randomUUID().replaceAll('-', '')}`,
      slideIndex: pageIndex,
      operations: fixes.map((fix) => ({
        kind: 'set_geometry' as const,
        sourceId: fix.sourceId,
        geometry: {
          x: geometryPxToPoints(fix.x, scale),
          y: geometryPxToPoints(fix.y, scale),
          width: geometryPxToPoints(fix.width, scale),
          height: geometryPxToPoints(fix.height, scale),
        },
      })),
    },
    signal,
  )
  return execution.receipt.status === 'applied' || execution.receipt.status === 'unchanged'
}

export function visualQcRequestBytes(request: AgentStreamRequest): number {
  return new TextEncoder().encode(JSON.stringify(request)).byteLength
}

export function createBoundedVisualQcTransport(delegate: AgentTransport): AgentTransport {
  return {
    stream(request, callbacks) {
      let bytes = Infinity
      try {
        bytes = visualQcRequestBytes(request)
      } catch {
        // A non-serializable request is unavailable at the same boundary.
      }
      if (bytes > VISUAL_QC_LIMITS.maxTransportRequestBytes) {
        queueMicrotask(() => callbacks.onError('quality_request_too_large'))
        return { cancel: () => {} }
      }
      return delegate.stream(request, callbacks)
    },
  }
}

export function visualQcImageRequestBytes(image: AgentImage): number {
  return new TextEncoder().encode(JSON.stringify({ images: [image] })).byteLength
}

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

export async function publishAppliedDeterministicQuality(opts: {
  transactionId: string
  receiptStatus: string
  sessionId: string
  pageIndexes: readonly number[]
  completedKeys: Set<string>
  access: DeckAccess
  prepareSlide: (pageIndex: number) => Promise<{
    status: string
    slideId?: string
    elementIds?: Readonly<Record<string, string>>
  }>
  publish: (receipt: PresentationQualityReceipt) => void
  signal?: AbortSignal
  isCurrent: () => boolean
}): Promise<number[]> {
  if (opts.receiptStatus !== 'applied') return []
  const key = `${opts.sessionId}:${opts.transactionId}`
  if (opts.completedKeys.has(key)) return []
  opts.completedKeys.add(key)
  while (opts.completedKeys.size > 100) {
    const oldest = opts.completedKeys.values().next().value
    if (oldest === undefined) break
    opts.completedKeys.delete(oldest)
  }
  const published: number[] = []
  for (const pageIndex of [...new Set(opts.pageIndexes)].slice(0, DETERMINISTIC_QC_MAX_PAGES)) {
    opts.signal?.throwIfAborted()
    if (!opts.isCurrent()) return published
    const prepared = await opts.prepareSlide(pageIndex)
    opts.signal?.throwIfAborted()
    if (!opts.isCurrent()) return published
    const slide = opts.access.getSlides()[pageIndex]
    if (prepared.status !== 'prepared' || !prepared.slideId || !slide) continue
    opts.publish(
      createDeterministicQualityReceipt(slide, {
        qualityRunId: `qc-det-${pageIndex}-${opts.transactionId.slice(0, 100)}`,
        transactionId: opts.transactionId,
        slideId: prepared.slideId,
        ...(prepared.elementIds
          ? { elementCreationId: (sourceId) => prepared.elementIds?.[sourceId] }
          : {}),
      }),
    )
    published.push(pageIndex)
  }
  return published
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

Do not invoke tools, quote slide text, change wording, add/delete elements, or move anything across slides.

Return ONLY strict JSON: {"status":"pass|needs_fix|cannot_verify","summary":"under 15 words","fixes":[{"sourceId":"existing id","x":0,"y":0,"width":100,"height":100}]}. Use fixes only for objective geometry defects, at most 8, inside the canvas. For clean slides use status pass and an empty fixes array. For defects that geometry alone cannot safely fix use cannot_verify and an empty fixes array.`

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
  fixes?: QcGeometryFix[]
  error?: string
}

export function toVisualQualityReceipt(
  qualityRunId: string,
  transactionId: string,
  slideId: string,
  result: QcPageResult,
): PresentationQualityReceipt {
  if (result.error) {
    if (result.error === 'cancelled' || result.error === 'stale_session') {
      return {
        qualityRunId,
        transactionId,
        slideId,
        source: 'visual',
        status: 'cancelled',
        code: result.error,
      }
    }
    return {
      qualityRunId,
      transactionId,
      slideId,
      source: 'visual',
      status: 'unavailable',
      code:
        result.error === 'screenshot_unavailable'
          ? 'screenshot_unavailable'
          : result.error === 'visual_capacity_exceeded'
            ? 'visual_capacity_exceeded'
            : 'transport_unavailable',
    }
  }
  return {
    qualityRunId,
    transactionId,
    slideId,
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

Inspect the screenshot and return the strict quality JSON. Geometry fixes may reference only the listed ids.`.slice(
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
    visualQcImageRequestBytes(screenshot) > VISUAL_QC_LIMITS.maxScreenshotRequestBytes
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
      let review: QcReview | undefined
      if (!r.error && after) {
        try {
          review = parseQcReview(r.reply, after)
        } catch {
          r = { reply: '', error: 'quality_review_invalid' }
        }
      }
      const postFindings = after
        ? auditSlideQuality(after, { slideId: `slide-${pageIndex + 1}` })
        : []
      const blockingFinding = postFindings.find((finding) => finding.severity === 'critical')
      resolve({
        ok: true,
        edited: false,
        reply:
          review?.status === 'pass' && blockingFinding
            ? blockingFinding.code
            : review?.status === 'pass'
              ? 'OK'
              : (review?.summary ?? r.reply.trim().replace(/\s+/g, ' ').slice(0, 240)),
        preIssues: preIssues.length,
        postIssues: postFindings.length,
        ...(review?.fixes.length ? { fixes: review.fixes } : {}),
        ...(r.error !== undefined ? { error: r.error } : {}),
      })
    }
    const loop = new AgentLoop({
      transport: createBoundedVisualQcTransport(transport),
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
