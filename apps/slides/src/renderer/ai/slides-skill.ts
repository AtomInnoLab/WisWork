import type { AgentSkill, FinalResponseReviewContext, ToolDisplay } from '@wiswork/agent-core'
import type {
  GroupRenderNode,
  PictureRenderNode,
  RenderNode,
  RenderSlide,
  ShapeRenderNode,
} from '@wiswork/pptx-render'
import type {
  AddSmartArtOp,
  AgentToolCall,
  AgentToolDef,
  EditParagraph,
  SlidesAcceptanceAuthorityRequest,
  SlidesAcceptanceAuthorityLease,
  SlidesAcceptanceTextProofRequest,
} from '../../shared/ipc'
import { auditSlideLayout, formatAudit } from './layout-audit'
import { runLayoutScript, type LayoutScriptElement, type SlideStylePatch } from './layout-script'
import { t } from '../i18n/locale'
import {
  textFamilyReceiptOutcome,
  textToolTransactionId,
  type TextFamilyExecutionResult,
  type TextFamilyTransactionRequest,
} from './presentation-text-transactions'
import type { SpeakerNotesDraftPreparation } from '../notes-draft'
import { selectionScopeSummary, type SelectionScope } from './edit-queue'
import {
  geometryPxToPoints,
  geometryToolTransactionId,
  normalizePresentationRotation,
  type GeometryFamilyTransactionRequest,
} from './presentation-geometry-transactions'
import {
  backgroundToolTransactionId,
  type BackgroundFamilyTransactionRequest,
} from './presentation-background-transactions'
import type { SlidesAcceptanceAuthority } from './task-acceptance'
import { compileCanonicalSlidesCalls, createSlidesTaskController } from './task-controller'
import type { SlidesTaskReviewAdapter } from './task-review'
import type { PresentationCompletionReceipt } from '@wiswork/presentation-verification'

/**
 * Slides capability as an AgentSkill: deck outline context + three tools (read structure /
 * read one slide / edit element text). Changes go through the existing slides:edit-text IPC;
 * the main process applies them and returns the new RenderSlide, which applySlide writes
 * back into React state — the same pipeline as manual editing.
 */

/** Panel/skill access point to the currently open deck (refs provided by App, stay fresh across renders). */
export interface DeckAccess {
  getSlides(): RenderSlide[]
  getCurrent(): number
  getSelectedIds(): string[]
  getAcceptanceAuthorityLease?(): Promise<SlidesAcceptanceAuthorityLease>
  verifyAcceptanceTextProof?(request: SlidesAcceptanceTextProofRequest): Promise<boolean>
  /** Read-only, revision-bound durable facts used to compile and verify a frozen task contract. */
  inspectAcceptanceAuthority?(
    request: SlidesAcceptanceAuthorityRequest,
  ): Promise<SlidesAcceptanceAuthority>
  taskReviewAdapter?: SlidesTaskReviewAdapter
  onTaskReviewComplete?(receipt: PresentationCompletionReceipt): void | Promise<void>
  beginTaskCorrectionHistory?(): Promise<boolean>
  finishTaskCorrectionHistory?(): Promise<string | undefined>
  /** Immutable durable scope for the active queued run, if any. */
  getSelectionScope?(): SelectionScope | undefined
  applySlide(slideIndex: number, updated: RenderSlide): void
  /** Replace the whole deck (after adding/removing slides) and jump to the goTo slide */
  applyDeck(slides: RenderSlide[], goTo?: number): void
  /** Execute one canonical, durable presentation transaction for an enrolled tool family. */
  executePresentationOperation?(
    request:
      | TextFamilyTransactionRequest
      | GeometryFamilyTransactionRequest
      | BackgroundFamilyTransactionRequest,
    signal?: AbortSignal,
  ): Promise<TextFamilyExecutionResult>
  /** Mirror an already-applied notes transaction into the visible notes editor. */
  prepareSpeakerNotesWrite?(slideIndex: number): Promise<SpeakerNotesDraftPreparation>
  applySpeakerNotes?(slideIndex: number, text: string, expectedDraftVersion: number): void
  /** Survey: shows a card with options and waits for the user's choices, returning an answer summary. */
  askClarification?(questions: ClarifyQuestion[]): Promise<{ answers: string; cancelled?: boolean }>
  /**
   * In-tool image search (embedded in the tool):
   * given English keywords, returns an array of real image URLs (at most N).
   * On search failure returns an empty array (fail-open; doesn't block the main generation path).
   */
  searchImages?(query: string, maxResults: number): Promise<string[]>
  /**
   * In-tool Style Skill generation:
   * a dedicated LLM call focused on producing a complete structured visual style guide
   * (color rules/fonts/layout variants per page type/overall style). Promotes style from an
   * "outline side-product" to a "dedicated deliverable" — less AI-looking, consistent across pages.
   */
  generateStyleSkill?(args: {
    topic: string
    questionnaire?: string
    styleHint?: string
    signal?: AbortSignal
  }): Promise<{ ok: boolean; styleSkill?: string; error?: string }>
  /**
   * In-tool planning: given topic + page count, the LLM produces a structured outline.
   * Fixes "missing pages at the input side" at the root — the main agent doesn't hand-write dozens of pages of pages JSON (avoids the argument being truncated by max_tokens).
   * style is already produced by generateStyleSkill; this function only outputs core_hook + per-page outlines (styleSkill serves as a reference for consistency).
   * Batched recursion is scheduled by the skill (continueFrom keeps the narrative coherent across batches).
   */
  planDeckOutline?(args: {
    topic: string
    count: number
    startPage: number
    context?: string
    styleSkill?: string
    continueFrom?: { coreHook: string }
    signal?: AbortSignal
  }): Promise<{
    ok: boolean
    // Same loose shape as OutlineJson (outline-json.ts): the LLM output is
    // only validated field-by-field at the point of use.
    outline?: { core_hook?: unknown; pages?: unknown }
    error?: string
  }>
  /**
   * Persist the current draft's Style Skill as a sidecar file (same directory and name as the draft, .styleskill.json).
   * fail-open: failure doesn't block the main path.
   */
  saveSidecar?(data: { topic: string; styleSkill: string; createdAt: string }): Promise<void>
  /**
   * Save styleSkill into userData/style-templates/<name>.json for later reuse.
   */
  saveStyleTemplate?(
    name: string,
    data: { topic: string; styleSkill: string; createdAt: string },
  ): Promise<{ ok: boolean; error?: string }>
  /**
   * List saved Style templates (name + topic + createdAt).
   */
  listStyleTemplates?(): Promise<Array<{ name: string; topic: string; createdAt: string }>>
  /**
   * Load the content of a given Style template.
   */
  loadStyleTemplate?(
    name: string,
  ): Promise<{ ok: boolean; styleSkill?: string; topic?: string; error?: string }>
  fitWidthPx: number
  /** Base retry backoff in ms for single-page generation failures (default 2000; tests pass 0 to disable backoff) */
  retryBackoffMs?: number
  /**
   * Names of text attachments in the current conversation that were never read with
   * read_attachment. When non-empty, deck construction must wait until they are read
   * (decks must be built from attachment content, not generic filler).
   */
  unreadTextAttachments?(): string[]
}

/** Single survey question structure (with options). */
export interface ClarifyQuestion {
  id: string
  label: string
  description?: string
  /** Option text array (≤5 per question); the frontend automatically appends "Other (fill in)" */
  options: string[]
  /** Multi-select (single-select by default) */
  multi?: boolean
}

const AGENT_SYSTEM_PROMPT = `You are the AI assistant inside WisWork Slides. Help users create, edit, and verify presentations with the available local tools.

## Workflow
- Start with get_deck_context, and use read_slide when exact text, colors, or element details matter.
- For a new presentation, use ask_clarification when requirements are ambiguous, then plan_deck. Build every page with add_slide and the local add_text_box, add_shape, add_chart, add_table, add_smartart, and insert_web_image tools. Empty decks are valid and may be built directly with these tools.
- Supported editing capability map:
  - Change existing text font, size, color, emphasis, or alignment with set_element_style; for coordinated edits use execute_slide_script and setStyle.
  - Move, resize, rotate, or align titles and other elements with set_element_transform; for coordinated edits use execute_slide_script and setBox/moveBy/resizeBy.
  - For a multi-page request, inspect and edit each target page (one tool call per target slideIndex), then verify the affected pages. Do not infer that changing one page changes the others.
- An edit request must not finish with inspection or advice only. Apply the requested supported edits and verify them before responding.
- Claim a capability limitation only after the relevant tool explicitly returns unsupported or fail-closed. Do not infer limitations from a read result or from unfamiliarity with a tool.
- Use execute_slide_script for coordinated edits to existing elements. Use the individual set_element_* tools for focused changes.
- Use set_speaker_notes to add, replace, or clear presenter notes without changing canvas content.
- Use web_search for current facts and image_search for real imagery. Never invent precise figures; declare figure provenance through dataSource.
- Keep layouts readable, consistent, within canvas bounds, and free of accidental overlaps. Verify the deck outline after substantial edits.
- Read all text attachments before using their content. Prefer user-provided material over generic filler.
- If a requested capability is unavailable, report that limitation clearly and continue with the supported local tools when possible.`

const SLIDES_CAPABILITY_CORRECTION =
  '[System correction] This requested edit is supported. Use the available Slides editing tools to apply it, verify the result, and only then report completion. Do not stop at inspection or advice.'

const DENIAL_ASSERTION =
  /(?:无法|不能|不支持|不可用|没有(?:办法|能力)|只能|\bcannot\b|\bcan't\b|\bunable to\b|\bdoes not support\b|\bnot supported\b|\bunavailable\b|\bonly (?:can|allows?)\b)/i
const TEXT_TARGET = /(?:字体|文字|文本|标题|重点文字|\bfont\b|\btext\b|\btitle\b)/i
const TEXT_STYLE =
  /(?:颜色|色彩|字号|字体|加粗|粗体|斜体|下划线|样式|格式|\bcolou?r\b|\bsize\b|\bfamily\b|\bstyle\b|\bformat\b|\bbold\b|\bitalic\b|\bunderline\b|\balign(?:ment)?\b)/i
const GEOMETRY_TARGET =
  /(?:标题(?:栏)?|元素|文本框|形状|图片|\btitle\b|\belement\b|\btext box\b|\bshape\b|\bimage\b)/i
const GEOMETRY_CHANGE =
  /(?:位置|尺寸|大小|移动|调整|对齐|旋转|坐标|布局|\bposition\b|\bsize\b|\bgeometry\b|\balign(?:ment|ed)?\b|\brotat(?:e|ed|ion)\b|\bmov(?:e|ed|ing)\b|\bresiz(?:e|ed|ing)\b)/i
const HOST_READ_FAILURE =
  /(?:(?:PowerPoint|幻灯片|PPT).{0,32}(?:read failed|read error|读取失败|操作失败|未能应用)|(?:read failed|read error|读取失败).{0,32}(?:PowerPoint|幻灯片|PPT))/i
const FAILURE_STATUS =
  /(?:unsupported|fail(?:ed)?[- ]closed|target_stale|proposal_stale|office_state_uncertain|failed|error|不支持|失败|报错)/i
const STYLE_TOOL = /(?:set_element_style|(?:style|format) (?:tool|operation)|样式编辑工具)/i
const GEOMETRY_TOOL =
  /(?:set_element_transform|(?:transform|geometry) (?:tool|operation)|几何编辑工具)/i
const SCRIPT_TOOL = /(?:execute_slide_script|脚本编辑工具)/i
const REAL_UNSUPPORTED_FAMILY =
  /(?:嵌入字体|字体打包|平滑切换|变体过渡|embedded fonts?|font packaging|morph transitions?)/i
const QUALIFIED_UNCERTAINTY = /(?:不能保证|cannot guarantee|can't guarantee)/i
const TARGET_CONSTRAINT =
  /(?:锁定|只读|版式装饰|嵌套(?:元素|目标|组)|\blocked\b|\bread-only\b|\blayout decoration\b|\bnested (?:target|element|group)\b)/gi
const CONCRETE_TARGET =
  /(?:标题(?:栏)?|元素|文本框|形状|图片|图表|目标|对象|字体|文字|文本|\btitle\b|\belement\b|\btext box\b|\bshape\b|\bimage\b|\bchart\b|\btarget\b|\bobject\b|\bfont\b|\btext\b)/gi
const DENIAL_ANCHOR = new RegExp(DENIAL_ASSERTION.source, 'gi')

const MAX_REVIEW_TEXT_CHARS = 4096
const MAX_REVIEW_CLAUSES = 64
const MAX_REVIEW_CLAUSE_CHARS = 512
const PROTECTED_PERIOD = '\u0000'

function protectCommonPeriods(text: string): string {
  return text
    .replace(/https?:\/\/[^\s。！？!?;；]+/gi, (url) =>
      url.replaceAll('.', (match, offset: number) =>
        offset === url.length - 1 ? match : PROTECTED_PERIOD,
      ),
    )
    .replace(/\b(?:e\.g\.|i\.e\.)/gi, (abbreviation) =>
      abbreviation.replaceAll('.', PROTECTED_PERIOD),
    )
    .replace(/(?<=\d)\.(?=\d)/g, PROTECTED_PERIOD)
}

function boundClause(clause: string): string {
  if (clause.length <= MAX_REVIEW_CLAUSE_CHARS) return clause
  const half = MAX_REVIEW_CLAUSE_CHARS / 2
  return `${clause.slice(0, half)}${clause.slice(-half)}`
}

function boundedClauses(text: string): string[] {
  const clauses =
    protectCommonPeriods(text.slice(0, MAX_REVIEW_TEXT_CHARS)).match(
      /[^.。！？!?;；\r\n]+[.。！？!?;；]?/g,
    ) ?? []
  const selected =
    clauses.length <= MAX_REVIEW_CLAUSES
      ? clauses
      : [...clauses.slice(0, MAX_REVIEW_CLAUSES / 2), ...clauses.slice(-MAX_REVIEW_CLAUSES / 2)]
  return selected
    .map((clause) => boundClause(clause.trim()).replaceAll(PROTECTED_PERIOD, '.'))
    .filter(Boolean)
}

type CapabilityFamily = 'text-style' | 'geometry'

function supportedCapabilityFamilies(clause: string): CapabilityFamily[] {
  const families: CapabilityFamily[] = []
  if (TEXT_TARGET.test(clause) && TEXT_STYLE.test(clause)) families.push('text-style')
  if (GEOMETRY_TARGET.test(clause) && GEOMETRY_CHANGE.test(clause)) families.push('geometry')
  return families
}

function hasRelatedToolFailure(clause: string, family: CapabilityFamily): boolean {
  if (!FAILURE_STATUS.test(clause)) return false
  if (SCRIPT_TOOL.test(clause)) return true
  return family === 'text-style' ? STYLE_TOOL.test(clause) : GEOMETRY_TOOL.test(clause)
}

type TargetKind = 'title' | 'element' | 'text-box' | 'shape' | 'image' | 'chart' | 'target' | 'text'

function targetKind(token: string): TargetKind {
  const normalized = token.toLowerCase()
  if (/标题|title/.test(normalized)) return 'title'
  if (/文本框|text box/.test(normalized)) return 'text-box'
  if (/形状|shape/.test(normalized)) return 'shape'
  if (/图片|image/.test(normalized)) return 'image'
  if (/图表|chart/.test(normalized)) return 'chart'
  if (/目标|对象|target|object/.test(normalized)) return 'target'
  if (/字体|文字|文本|font|text/.test(normalized)) return 'text'
  return 'element'
}

interface TargetMention {
  kind: TargetKind
  index: number
  identity?: string
  coreference: boolean
}

function targetMention(clause: string, match: RegExpMatchArray): TargetMention {
  const index = match.index ?? 0
  const before = clause.slice(Math.max(0, index - 24), index)
  const after = clause.slice(index + match[0].length, index + match[0].length + 40)
  const chineseOrdinal = before.match(/第([一二三四五六七八九十百\d]+)个?\s*$/)
  const englishOrdinal = before.match(/\b(first|second|third|fourth|fifth|\d+(?:st|nd|rd|th))\s+$/i)
  const explicitId = after.match(/^\s+id\s*[:=#]\s*([A-Za-z0-9_-]+)/i)
  const shortSuffix = after.match(/^\s+([A-Z]|\d+)\b/)
  const bareId = after.match(/^\s+([A-Za-z0-9]+(?:[-_][A-Za-z0-9_-]+))(?=\s|[,.;:!?，。；！？]|$)/)
  const identity = explicitId?.[1]
    ? `id:${explicitId[1].toLowerCase()}`
    : bareId?.[1]
      ? `id:${bareId[1].toLowerCase()}`
      : shortSuffix?.[1]
        ? `label:${shortSuffix[1].toLowerCase()}`
        : chineseOrdinal?.[1]
          ? `ordinal:${chineseOrdinal[1]}`
          : englishOrdinal?.[1]
            ? `ordinal:${englishOrdinal[1].toLowerCase()}`
            : undefined
  return {
    kind: targetKind(match[0]),
    index,
    ...(identity ? { identity } : {}),
    coreference: /(?:该|此|同一)\s*$/.test(before) || /\b(?:this|that|same)\s+$/i.test(before),
  }
}

function targetMentions(clause: string): TargetMention[] {
  return [...clause.matchAll(CONCRETE_TARGET)].map((match) => targetMention(clause, match))
}

function nearestTarget(clause: string, index: number): TargetMention | undefined {
  const targets = targetMentions(clause)
  let nearest: TargetMention | undefined
  let distance = Number.POSITIVE_INFINITY
  for (const target of targets) {
    const nextDistance = Math.abs(target.index - index)
    if (nextDistance < distance) {
      nearest = target
      distance = nextDistance
    }
  }
  return nearest
}

function constrainedTarget(clause: string, index: number): TargetMention | undefined {
  const targets = targetMentions(clause)
  const immediateFollowing = targets.find((target) => {
    return target.index >= index && /^[\s的]{0,5}$/.test(clause.slice(index, target.index))
  })
  if (immediateFollowing) return immediateFollowing
  return targets.filter((target) => target.index < index).at(-1) ?? targets[0]
}

function sameTarget(constrained: TargetMention, denied: TargetMention): boolean {
  const compatibleKind =
    constrained.kind === denied.kind || constrained.kind === 'target' || denied.kind === 'target'
  if (!compatibleKind) return false
  if (constrained.identity && denied.identity) return constrained.identity === denied.identity
  if (constrained.identity) return denied.coreference
  if (denied.identity) return constrained.coreference
  if (constrained.coreference || denied.coreference) return true
  return true
}

function hasSameTargetConstraint(clause: string): boolean {
  const deniedTargets = [...clause.matchAll(DENIAL_ANCHOR)]
    .map((match) => nearestTarget(clause, match.index ?? 0))
    .filter((target): target is TargetMention => Boolean(target))
  if (deniedTargets.length === 0) return false
  const constrainedTargets = [...clause.matchAll(TARGET_CONSTRAINT)]
    .map((match) => constrainedTarget(clause, (match.index ?? 0) + match[0].length))
    .filter((target): target is TargetMention => Boolean(target))
  return constrainedTargets.some((constrained) =>
    deniedTargets.some((denied) => sameTarget(constrained, denied)),
  )
}

function isLocallyJustified(clause: string, family: CapabilityFamily): boolean {
  return (
    HOST_READ_FAILURE.test(clause) ||
    REAL_UNSUPPORTED_FAMILY.test(clause) ||
    hasRelatedToolFailure(clause, family) ||
    hasSameTargetConstraint(clause) ||
    QUALIFIED_UNCERTAINTY.test(clause)
  )
}

/**
 * Rejects only a narrow, tool-free false denial of an editing family that
 * Slides already exposes. The correction is constant and cannot echo deck or
 * response content into a later model turn.
 */
export function reviewSlidesFinalResponse({
  text,
  mutated,
}: FinalResponseReviewContext): string | undefined {
  if (mutated || !text.trim()) return undefined
  for (const clause of boundedClauses(text)) {
    if (/[?？]\s*$/.test(clause)) continue
    if (!DENIAL_ASSERTION.test(clause)) continue
    const families = supportedCapabilityFamilies(clause)
    if (families.some((family) => !isLocallyJustified(clause, family))) {
      return SLIDES_CAPABILITY_CORRECTION
    }
  }
  return undefined
}

/** Paragraph schema (shared by set_element_text / add_text_box / add_shape) */
const PARAGRAPHS_DEF = {
  paragraphs: {
    type: 'array',
    description: 'Complete paragraph list, one object per paragraph',
    items: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'Paragraph plain text' },
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        underline: { type: 'boolean' },
        fontSize: { type: 'number', description: 'Font size (pt)' },
        fontFamily: {
          type: 'string',
          description: 'Font name; omit to inherit the theme font (recommended)',
        },
        color: { type: 'string', description: '#RRGGBB' },
        align: { type: 'string', enum: ['left', 'center', 'right'] },
      },
      required: ['text'],
    },
  },
} as const

interface ToolParagraph {
  text?: unknown
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontSize?: number
  fontFamily?: string
  color?: string
  align?: 'left' | 'center' | 'right'
}

function toEditParagraphs(raw: unknown): EditParagraph[] | null {
  if (!Array.isArray(raw) || raw.length === 0) return null
  return raw.map((p) => {
    const para = p as ToolParagraph
    return {
      runs: [
        {
          text: String(para.text ?? ''),
          ...(para.bold ? { bold: true } : {}),
          ...(para.italic ? { italic: true } : {}),
          ...(para.underline ? { underline: true } : {}),
          ...(typeof para.fontSize === 'number' ? { fontSize: para.fontSize } : {}),
          ...(para.fontFamily ? { fontFamily: para.fontFamily } : {}),
          ...(para.color ? { color: para.color } : {}),
        },
      ],
      ...(para.align ? { align: para.align } : {}),
    }
  })
}

const ALL_TOOLS: AgentToolDef[] = [
  {
    name: 'get_deck_context',
    description:
      "Get the deck's latest outline: per-page list of text elements (element id | type | text preview). Call to confirm global state after edits.",
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'read_slide',
    description:
      'Read all elements of a page with full text (untruncated) and current colors (fill/text/stroke, hex). Call before rewriting a page.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based)' },
      },
      required: ['slideIndex'],
    },
  },
  {
    name: 'set_element_text',
    description:
      "Replace a text element's entire content. paragraphs is the complete post-replacement paragraph array, one object per paragraph; whole-paragraph bold/italic etc. use the boolean fields on the paragraph object.",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based)' },
        sourceId: { type: 'string', description: 'Element id (from the outline/read_slide)' },
        paragraphs: { $ref: '#/definitions/paragraphs' },
      },
      required: ['slideIndex', 'sourceId', 'paragraphs'],
      definitions: PARAGRAPHS_DEF,
    },
  },
  {
    name: 'set_element_style',
    description:
      "Change an element's text formatting without changing the text: font size/color/bold/italic/underline/alignment/font. " +
      "Pass only the fields to change; others stay as-is. Applies to the element's entire text.",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based)' },
        sourceId: { type: 'string', description: 'Element id' },
        fontSize: { type: 'number', description: 'Font size (pt)' },
        color: { type: 'string', description: '#RRGGBB' },
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        underline: { type: 'boolean' },
        fontFamily: { type: 'string', description: 'Font name; usually omit to inherit the theme' },
        align: { type: 'string', enum: ['left', 'center', 'right'] },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
  {
    name: 'set_element_transform',
    description:
      'Move/resize/rotate an element (pixel coordinates, origin top-left, canvas 1280 wide). Pass only the fields to change; others keep their values.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string' },
        x: { type: 'number', description: 'Top-left x (px)' },
        y: { type: 'number', description: 'Top-left y (px)' },
        w: { type: 'number', description: 'Width (px)' },
        h: { type: 'number', description: 'Height (px)' },
        rotationDeg: { type: 'number', description: 'Rotation angle (degrees, clockwise)' },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
  {
    name: 'execute_slide_script',
    description:
      '[Preferred tool for editing a slide] Runs your JS edit script against one page; a single script covers: add/delete/position/size/alignment/distribution/relative nudges/text/style/fill/stroke.' +
      ' At run time the script automatically receives the real geometry and text of every element on the page (els) — **no read_slide needed first**; read-write combined, compute from els inside the script.' +
      ' Geometry changes are applied atomically in one batch (undoable as a whole), the rest in script order, and a layout audit (overlap/out-of-bounds/text overflow) is returned at the end.' +
      ' Far more reliable than individual set_element_* calls — coordinate math happens at execution site, not from memory. If the audit reports problems, call this tool again immediately to fix.\n' +
      'Script environment (constrained synchronous JS-like DSL; no external APIs or ambient globals):\n' +
      '- els: array, each item {id,type,text,x,y,w,h,rotation,fontSizePt?,fill?,textColor?,strokeColor?,inGroup?,groupId?,locked?} (pixels, origin top-left; fill/textColor/strokeColor are current colors in #RRGGBB, read-only — write via setFill/setStyle/setStroke; inGroup+groupId=directly editable group child (all primitives work, coordinates absolute as shown); inGroup without groupId=nested in a sub-group, read-only — ungroup_element the outer group first; locked=layout decoration, read-only)\n' +
      '- canvas: {w,h} canvas size (px)\n' +
      "- setBox(id, {x?,y?,w?,h?,rotation?}): set an element's target box, pass only fields to change\n" +
      '- moveBy(id, dx, dy): relative move (left = negative dx, up = negative dy)\n' +
      '- resizeBy(id, dw, dh): relative resize\n' +
      "- setText(id, textOrParagraphs): replace text entirely; pass a string (split into paragraphs by \\n) or a paragraph array (same format as set_element_text's paragraphs)\n" +
      '- setStyle(id, {fontSize?,color?,bold?,italic?,underline?,align?,fontFamily?}): change style without changing text, pass only fields to change\n' +
      "- setFill(id, colorOrNone): solid fill '#RRGGBB' or 'none'\n" +
      '- setStroke(id, {color?,widthPt?} | null): stroke; pass null to remove\n' +
      '- addText(clientId, textOrParagraphs, {x,y,w,h,rotation?}): add a text box and return its transaction-local id; later primitives may use that id\n' +
      '- delete(id): delete a top-level existing element or a text box added earlier in this script; later use of that id fails closed\n' +
      '- log(...): debug output (echoed back to you); the return value is echoed back to you (put a summary there)\n' +
      '- Supported computation: const/let, arithmetic, if/for/for...of/while, functions/arrows, JSON object/array literals, Math, regex.test, and safe array/string methods. No classes, async, modules, constructors, prototypes, or dynamic code.\n' +
      'Example 1 — three cards equal width, equal spacing:\n' +
      'const cards = els.filter(e => /card/.test(e.id));\n' +
      'const gap = 32, w = (canvas.w - 2*80 - (cards.length-1)*gap) / cards.length;\n' +
      'cards.forEach((c, i) => setBox(c.id, { x: 80 + i*(w+gap), y: 200, w, h: 320 }));\n' +
      "Example 2 — move the title left a bit: moveBy('title', -30, 0);\n" +
      "Example 3 — make the title blue and bold: setStyle('t1', { color: '#1a73e8', bold: true });",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based)' },
        code: {
          type: 'string',
          description:
            'JS script body (synchronous code; may use els/canvas plus addText/delete/setBox/moveBy/resizeBy/setText/setStyle/setFill/setStroke/log; may return a summary)',
        },
        explanation: {
          type: 'string',
          description:
            'One sentence describing what this script does (≤60 chars, shown to the user)',
        },
      },
      required: ['slideIndex', 'code'],
    },
  },
  {
    name: 'set_element_fill',
    description: 'Set an element\'s solid fill. fill=#RRGGBB; pass "none" for no fill.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string' },
        fill: { type: 'string', description: '#RRGGBB or none' },
      },
      required: ['slideIndex', 'sourceId', 'fill'],
    },
  },
  {
    name: 'set_element_stroke',
    description:
      "Set an element's stroke. Pass color (#RRGGBB) + widthPt (points); to remove the stroke pass remove=true.",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string' },
        color: { type: 'string', description: '#RRGGBB' },
        widthPt: { type: 'number', description: 'Line width (points), default 1' },
        remove: { type: 'boolean', description: 'true = remove stroke' },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
  {
    name: 'web_search',
    description:
      'Web search for text information (material/data/facts). Use when you need current information or are unsure about facts. Returns title/link/snippet.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search keywords' },
        maxResults: { type: 'integer', description: 'Max results, default 6' },
      },
      required: ['query'],
    },
  },
  {
    name: 'image_search',
    description:
      'Search image assets (for slide imagery). Returns a list of imageUrl; after choosing, insert with insert_web_image.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Image search keywords (English works better)' },
        maxResults: { type: 'integer', description: 'Max results, default 8' },
      },
      required: ['query'],
    },
  },
  {
    name: 'insert_web_image',
    description:
      'Download an image URL obtained from image_search and insert it into a page (pixel coordinates).',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        url: { type: 'string', description: 'Direct image link (imageUrl from image_search)' },
        x: { type: 'number' },
        y: { type: 'number' },
        w: { type: 'number' },
        h: { type: 'number' },
      },
      required: ['slideIndex', 'url', 'x', 'y', 'w', 'h'],
    },
  },
  {
    name: 'crop_image',
    description:
      'Crop a picture non-destructively (srcRect): l/t/r/b are fractions (0..1) cut from each edge of the source image. The element frame stays where it is; the remaining region stretches to fill it. Pass all zeros to remove an existing crop.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Picture element id' },
        l: { type: 'number', description: 'Fraction cut from the left edge (0..1)' },
        t: { type: 'number', description: 'Fraction cut from the top edge (0..1)' },
        r: { type: 'number', description: 'Fraction cut from the right edge (0..1)' },
        b: { type: 'number', description: 'Fraction cut from the bottom edge (0..1)' },
      },
      required: ['slideIndex', 'sourceId', 'l', 't', 'r', 'b'],
    },
  },
  {
    name: 'set_picture_opacity',
    description:
      "Set a picture's whole-image opacity. opacity 0..1; 1 = fully opaque (removes the effect).",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Picture element id' },
        opacity: { type: 'number', description: '0 (invisible) .. 1 (opaque)' },
      },
      required: ['slideIndex', 'sourceId', 'opacity'],
    },
  },
  {
    name: 'replace_image',
    description:
      "Swap a picture's source image for a direct URL (for example, from image_search) in place so position, size, z-order, border, and effects survive. keepCrop preserves the existing crop window and is only correct when the new image has the same pixel geometry as the old one.",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Picture element id' },
        url: { type: 'string', description: 'Direct image link' },
        keepCrop: { type: 'boolean', description: 'Keep the existing crop window (default false)' },
      },
      required: ['slideIndex', 'sourceId', 'url'],
    },
  },
  {
    name: 'ask_clarification',
    description:
      "[Call before creating a whole new deck] Shows a questionnaire card with options, letting the user make key choices for this deck (audience/scenario/tone/focus etc.); the user's choices directly determine the deck's Core Hook and style. Questions must target the specific topic, each being a real trade-off (options represent different directions). Ask 2–4 questions, ≤5 options each. After calling, wait for the user to finish choosing in the card and generate once you have the answers. Don't repeat the questions in your reply text.",
    inputSchema: {
      type: 'object',
      properties: {
        questions: {
          type: 'array',
          description: 'Question list (2–4 questions)',
          items: {
            type: 'object',
            properties: {
              id: { type: 'string', description: 'Unique question id (short English/pinyin)' },
              label: { type: 'string', description: 'Question text' },
              description: { type: 'string', description: 'Optional one-line note' },
              options: {
                type: 'array',
                items: { type: 'string' },
                description:
                  'Options (≤5); the frontend automatically appends "Decide for me" and "Other"',
              },
              multi: { type: 'boolean', description: 'Multi-select (single-select by default)' },
            },
            required: ['id', 'label', 'options'],
          },
        },
      },
      required: ['questions'],
    },
  },
  {
    name: 'plan_deck',
    description:
      "[When creating a whole new deck, call after researching material and images] Outputs a structured plan: the Core Hook + unified style scheme + each page's title/content brief/layout/image keywords. Think the whole deck through first, to avoid starting strong and fizzling out. The plan is echoed to the user.",
    inputSchema: {
      type: 'object',
      properties: {
        core_hook: {
          type: 'string',
          description:
            "The deck's narrative anchor (one sentence, with tension, ideally containing a number or counter-intuitive contrast)",
        },
        style: {
          type: 'string',
          description:
            'Unified design system: primary/secondary colors, font tone, content margins, card/corner style (e.g. "dark blue primary + gold accents, data-dashboard look"); every page follows it',
        },
        pages: {
          type: 'array',
          description: 'Per-page plan',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string', description: 'Page title' },
              type: { type: 'string', description: 'cover|content|data|closing' },
              brief: {
                type: 'string',
                description:
                  'Page content description (use real data/facts; say what goes in each region)',
              },
              layout: {
                type: 'string',
                description:
                  'Layout (e.g. three_column_cards/hero_big_number/two_column/timeline/left_text_right_image); content pages must not repeat',
              },
              image_queries: {
                type: 'array',
                items: { type: 'string' },
                description:
                  "English image-search keywords for this page's image slots (one per slot; [] for no images)",
              },
            },
            required: ['title', 'brief', 'layout'],
          },
        },
      },
      required: ['core_hook', 'style', 'pages'],
    },
  },
  {
    name: 'delete_slide',
    description:
      "Delete an entire page (not allowed when only one page remains). After deletion, later pages' slideIndex shifts down.",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based)' },
      },
      required: ['slideIndex'],
    },
  },
  {
    name: 'save_style_template',
    description:
      '[Save the current deck\'s style as a reusable template] Saves the current presentation\'s Style Skill (visual style guide) under the given name; next time you generate a deck, pass the style_template argument to reuse it directly and skip style generation. Call when the user says "save this style" / "save as template".',
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'Template name (short, e.g. "minimal-blue" or "tech-dark")',
        },
      },
      required: ['name'],
    },
  },
  {
    name: 'list_style_templates',
    description:
      'List all saved style templates (name + topic + createdAt). When the user says "use last time\'s style" or "use some template", call this first to see what exists, then load the selected template for local reuse.',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'add_slide',
    description:
      "Create a new page by cloning the layout (including background) of page sourceIndex, inserted right after it (new page number = sourceIndex+1; pages after it shift back); clearText=true (default) clears text to get a layout-preserving blank page. When building page by page, use the CURRENT LAST page as sourceIndex so new pages append at the end. The return value gives the new page's slideIndex; subsequent content fills MUST use that returned page number.",
    inputSchema: {
      type: 'object',
      properties: {
        sourceIndex: {
          type: 'integer',
          description: 'Page to use as the layout template (0-based)',
        },
        clearText: {
          type: 'boolean',
          description: "Default true; false keeps the template page's text",
        },
      },
      required: ['sourceIndex'],
    },
  },
  {
    name: 'add_text_box',
    description: 'Create a new text box on a page (pixel coordinates). Returns the new element id.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        x: { type: 'number' },
        y: { type: 'number' },
        w: { type: 'number' },
        h: { type: 'number' },
        paragraphs: { $ref: '#/definitions/paragraphs' },
      },
      required: ['slideIndex', 'x', 'y', 'w', 'h', 'paragraphs'],
      definitions: PARAGRAPHS_DEF,
    },
  },
  {
    name: 'add_shape',
    description:
      'Create a new shape on a page (optionally with solid fill and text). kind uses OOXML preset geometry names, common ones: rect/roundRect/ellipse/triangle/diamond/rightArrow/leftArrow/chevron/star5/heart/pie/donut/cloud/wedgeRoundRectCallout. Returns the new element id.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        kind: {
          type: 'string',
          description:
            'OOXML preset geometry name, e.g. rect / roundRect / ellipse / rightArrow / star5',
        },
        x: { type: 'number' },
        y: { type: 'number' },
        w: { type: 'number' },
        h: { type: 'number' },
        fillColor: { type: 'string', description: '#RRGGBB' },
        paragraphs: { $ref: '#/definitions/paragraphs' },
      },
      required: ['slideIndex', 'kind', 'x', 'y', 'w', 'h'],
      definitions: PARAGRAPHS_DEF,
    },
  },
  {
    name: 'add_chart',
    description:
      "Insert a chart on a page (native pptx chart, still editable in PowerPoint). categories are the x-axis categories; series is each series' name and values (length must match categories). Omit x/y/w/h to center it. dataSource declares where the numbers came from and is enforced — never present invented numbers as real data.",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        kind: { type: 'string', enum: ['bar', 'barStacked', 'line', 'area', 'pie', 'doughnut'] },
        title: { type: 'string', description: 'Chart title (optional)' },
        categories: { type: 'array', items: { type: 'string' } },
        series: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              name: { type: 'string' },
              values: { type: 'array', items: { type: 'number' } },
            },
            required: ['name', 'values'],
          },
        },
        dataSource: {
          type: 'string',
          enum: ['user', 'document', 'search', 'sample'],
          description:
            "Provenance of the values: 'user' = supplied by the user/attachments, 'document' = read from this deck, 'search' = from web_search results in this conversation (run it first), 'sample' = illustrative placeholders you must disclose to the user",
        },
        x: { type: 'number' },
        y: { type: 'number' },
        w: { type: 'number' },
        h: { type: 'number' },
      },
      required: ['slideIndex', 'kind', 'categories', 'series', 'dataSource'],
    },
  },
  {
    name: 'add_smartart',
    description:
      'Insert a SmartArt-style diagram (shape composition) on a page: list=vertical list, process=process arrows, cycle=cycle, hierarchy=org structure, pyramid=stacked pyramid levels, matrix=2x2 quadrant grid, venn=overlapping circles. items are the node texts. Omit x/y/w/h to center it.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        layout: {
          type: 'string',
          enum: ['list', 'process', 'cycle', 'hierarchy', 'pyramid', 'matrix', 'venn'],
        },
        items: {
          type: 'array',
          items: { type: 'string' },
          description: 'Node texts (2-8 recommended)',
        },
        x: { type: 'number' },
        y: { type: 'number' },
        w: { type: 'number' },
        h: { type: 'number' },
      },
      required: ['slideIndex', 'layout', 'items'],
    },
  },
  {
    name: 'add_table',
    description:
      'Insert a native pptx table on a page (with built-in styling, still editable in PowerPoint). cells gives text row by row (optional; ' +
      'missing rows/columns stay empty). Omit x/y/w/h to center it.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        rows: { type: 'integer', description: 'Row count (including header)' },
        cols: { type: 'integer', description: 'Column count' },
        cells: {
          type: 'array',
          items: { type: 'array', items: { type: 'string' } },
          description: 'Cell texts, row by row, e.g. [["Name","Qty"],["A","1"]]',
        },
        x: { type: 'number' },
        y: { type: 'number' },
        w: { type: 'number' },
        h: { type: 'number' },
      },
      required: ['slideIndex', 'rows', 'cols'],
    },
  },
  {
    name: 'edit_table_cell',
    description:
      "Replace one table cell's text entirely. The table element id comes from the outline/read_slide (type=table); row/col are 0-based.",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Table element id' },
        row: { type: 'integer', description: 'Row number (0-based)' },
        col: { type: 'integer', description: 'Column number (0-based)' },
        paragraphs: { $ref: '#/definitions/paragraphs' },
      },
      required: ['slideIndex', 'sourceId', 'row', 'col', 'paragraphs'],
      definitions: PARAGRAPHS_DEF,
    },
  },
  {
    name: 'edit_table_structure',
    description:
      'Add/remove table rows/columns: kind=insert-row/delete-row/insert-col/delete-col; index is the row/column number (0-based), insert defaults to after it, before=true inserts before it.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Table element id' },
        kind: { type: 'string', enum: ['insert-row', 'delete-row', 'insert-col', 'delete-col'] },
        index: { type: 'integer', description: 'Row/column number (0-based)' },
        before: { type: 'boolean', description: 'For insert, set true to insert before index' },
      },
      required: ['slideIndex', 'sourceId', 'kind', 'index'],
    },
  },
  {
    name: 'edit_table_style',
    description:
      'Modify table styling: apply a preset (styleName) or individually change header row/banding/shading/borders. styleName options: none/lightGrid/zebraBlue/zebraGray/headerDarkBlue/headerOrange/noBorder/fullBorder.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Table element id' },
        styleName: {
          type: 'string',
          description: 'Preset style name (see description), highest priority',
        },
        firstRow: { type: 'boolean', description: 'Enable header row (first-row emphasis)' },
        bandRow: { type: 'boolean', description: 'Enable banded rows' },
        shadingColor: {
          type: 'string',
          description: 'Shading color #RRGGBB, "none" clears shading',
        },
        borderColor: { type: 'string', description: 'Border color #RRGGBB' },
        borderWidthPt: { type: 'number', description: 'Border width (pt)' },
        borderPreset: {
          type: 'string',
          enum: ['all', 'none'],
          description: '"all" = full borders, "none" = clear borders',
        },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
  {
    name: 'edit_chart',
    description:
      'Modify a chart (including charts from imported files; first edit converts it to editable automatically): change type/data/colors/chart elements. kind options: bar/barStacked/line/area/pie/doughnut. colorScheme: default/colorful/colorful2/mono-accent1..6 (theme-derived); legacy keys blue/warm/cool/mono still work.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'Chart element id (type=chart)' },
        kind: {
          type: 'string',
          enum: ['bar', 'barStacked', 'line', 'area', 'pie', 'doughnut'],
          description: 'Change chart type (optional)',
        },
        categories: {
          type: 'array',
          items: { type: 'string' },
          description: 'X-axis/category labels (optional)',
        },
        series: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              name: { type: 'string' },
              values: { type: 'array', items: { type: 'number' } },
            },
            required: ['name', 'values'],
          },
          description: 'Data series (optional)',
        },
        dataSource: {
          type: 'string',
          enum: ['user', 'document', 'search', 'sample'],
          description:
            "Required when passing series: provenance of the values ('user'/'document'/'search'/'sample'; 'search' needs a prior web_search, 'sample' must be disclosed to the user)",
        },
        colorScheme: {
          type: 'string',
          description: 'Color scheme (optional): default/colorful/colorful2/mono-accent1..6',
        },
        title: { type: 'string', description: 'Chart title (optional)' },
        legendPos: {
          type: 'string',
          enum: ['b', 't', 'r', 'l', 'none'],
          description: 'Legend position (optional)',
        },
        dataLabels: { type: 'boolean', description: 'Data labels toggle (optional)' },
        gridlines: { type: 'boolean', description: 'Value-axis gridlines toggle (optional)' },
        switchRowCol: {
          type: 'boolean',
          description: 'Switch rows/columns: categories ↔ series (optional)',
        },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
  {
    name: 'set_slide_background',
    description: 'Set a solid page background color. slideIndex=-1 applies to all pages.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based); -1 = all pages' },
        color: { type: 'string', description: '#RRGGBB' },
      },
      required: ['slideIndex', 'color'],
    },
  },
  {
    name: 'set_speaker_notes',
    description:
      'Replace the complete speaker notes for one slide. Newlines separate paragraphs; an empty string clears the notes. This changes presenter notes only, not canvas content.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based)' },
        text: {
          type: 'string',
          description: 'Complete speaker notes text, at most 12000 characters',
        },
      },
      required: ['slideIndex', 'text'],
    },
  },
  {
    name: 'delete_element',
    description: 'Delete one element from a page.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string' },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
  {
    name: 'ungroup_element',
    description:
      'Ungroup a group element: promote its direct children to top-level page elements (positions/sizes preserved). Use when group members must be edited/deleted independently (e.g. elements nested in a sub-group, or deleting a single member). Note: ungrouping rewrites the page, so all element ids on it change — use the fresh ids returned in the result.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Page number (0-based)' },
        sourceId: { type: 'string', description: 'Group element id' },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
]

const UNSUPPORTED_CLOUD_TOOLS = new Set([
  'generate_image',
  'analyze_media',
  'regenerate_slide',
  'generate_deck',
])
const TOOLS = ALL_TOOLS.filter((tool) => !UNSUPPORTED_CLOUD_TOOLS.has(tool.name))

/** Collect readable text of nodes (including nested group children); returns a list of [sourceId, type, text] */
/** Find one node by id in the node tree (including groups). */
function findNodeById(nodes: RenderNode[], id: string): RenderNode | undefined {
  for (const n of nodes) {
    if (n.sourceId === id) return n
    if (n.type === 'group') {
      const hit = findNodeById(n.children, id)
      if (hit) return hit
    }
  }
  return undefined
}

/**
 * Editable context of an element: a top-level node, or a direct child of a top-level group
 * (with groupId + the group's absolute origin for abs↔group-local px conversion — child render
 * boxes are group-local, matching the in-group edit IPCs). Deeper nesting returns {nested:true}:
 * the main process patches one level only, so those stay read-only until ungrouped.
 */
type EditTarget =
  { node: RenderNode; groupId?: string; groupOrigin?: { x: number; y: number } } | { nested: true }
function resolveEditTarget(slide: RenderSlide, sourceId: string): EditTarget | null {
  for (const n of slide.nodes) {
    if (n.sourceId === sourceId) return { node: n }
    if (n.type === 'group') {
      const g = n as GroupRenderNode
      const child = g.children.find((c) => c.sourceId === sourceId)
      if (child)
        return {
          node: child,
          groupId: n.sourceId,
          groupOrigin: { x: Math.round(n.box.x), y: Math.round(n.box.y) },
        }
      if (findNodeById(g.children, sourceId)) return { nested: true }
    }
  }
  return null
}

/** Shared not-found / nested-in-subgroup error text for element-targeting tools. */
function targetError(target: EditTarget | null, sourceId: string, pageNo: number): string | null {
  if (!target)
    return `Element ${sourceId} not found on page ${pageNo} (ids change after regenerate/ungroup/save; call read_slide for fresh ids)`
  if ('nested' in target)
    return `Element ${sourceId} is nested inside a sub-group; call ungroup_element on the outer group first, or edit the sub-group as a whole`
  return null
}

/**
 * Restore a render node's current text into EditParagraph[] (aggregate runs by line, keeping
 * each run's existing formatting). Used by set_element_style: change formatting while keeping
 * the text. fontSize is converted back from px to pt.
 */
function nodeToParagraphs(node: ShapeRenderNode): EditParagraph[] {
  const lines = node.text?.lines ?? []
  return lines.map((line) => ({
    runs: line.runs.map((r) => ({
      text: r.text,
      ...(r.bold ? { bold: true } : {}),
      ...(r.italic ? { italic: true } : {}),
      ...(r.underline ? { underline: true } : {}),
      ...(r.fontSizePx ? { fontSize: Math.round((r.fontSizePx * 72) / 96) } : {}),
      ...(r.fontFamily ? { fontFamily: r.fontFamily } : {}),
      ...(r.color ? { color: r.color } : {}),
    })),
  }))
}

/**
 * Merge style-override fields into existing paragraphs: bold/italic/font size/color/font are
 * overridden per run, align is set on the paragraph, fields not passed stay unchanged. Shared
 * by the set_element_style tool and execute_slide_script's setStyle dispatch.
 */
function mergeStyleIntoParagraphs(cur: EditParagraph[], ov: SlideStylePatch): EditParagraph[] {
  return cur.map((p) => ({
    runs: p.runs.map((r) => ({
      text: r.text,
      bold: ov.bold ?? r.bold,
      italic: ov.italic ?? r.italic,
      underline: ov.underline ?? r.underline,
      fontSize: typeof ov.fontSize === 'number' ? ov.fontSize : r.fontSize,
      fontFamily: ov.fontFamily ?? r.fontFamily,
      color: ov.color ?? r.color,
    })),
    align: ov.align ?? p.align,
  }))
}

/** Element info shared by outline/read_slide/edit scripts (includes absolute geometry; locked = layout decoration, read-only). */
type NodeInfo = LayoutScriptElement

function nodeText(n: RenderNode): string {
  if (n.type === 'shape' || n.type === 'text') {
    return ((n as ShapeRenderNode).text?.lines ?? [])
      .map((line) => line.runs.map((r) => r.text).join(''))
      .join('\n')
  }
  if (n.type === 'table') {
    // Tables join cell text row by row (tab-separated) so the AI can read table content
    const byRow = new Map<number, string[]>()
    for (const c of n.cells) {
      const t = (c.text?.lines ?? []).map((l) => l.runs.map((r) => r.text).join('')).join(' ')
      const row = byRow.get(c.y) ?? []
      row.push(t)
      byRow.set(c.y, row)
    }
    return [...byRow.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([, r]) => r.join('\t'))
      .join('\n')
  }
  return ''
}

/** Max font size of the text (pt, converted back from px); returns undefined when there is no text. */
function nodeMaxFontPt(n: RenderNode): number | undefined {
  if (n.type !== 'shape' && n.type !== 'text') return undefined
  let maxPx = 0
  for (const line of (n as ShapeRenderNode).text?.lines ?? []) {
    for (const r of line.runs) if (r.fontSizePx > maxPx) maxPx = r.fontSizePx
  }
  return maxPx > 0 ? Math.round((maxPx * 72) / 96) : undefined
}

/** Normalize a render color to #RRGGBB (strips alpha); undefined when not a hex color. */
function hex6(color: string | undefined): string | undefined {
  if (!color) return undefined
  const m = /^#([0-9a-fA-F]{6})/.exec(color.trim())
  return m ? `#${m[1].toUpperCase()}` : undefined
}

/** Dominant text color = the run color covering the most characters (bullets excluded). */
function dominantTextColor(n: RenderNode): string | undefined {
  if (n.type !== 'shape' && n.type !== 'text') return undefined
  const weight = new Map<string, number>()
  for (const line of (n as ShapeRenderNode).text?.lines ?? []) {
    for (const r of line.runs) {
      if (r.isBullet) continue
      const c = hex6(r.color)
      if (c) weight.set(c, (weight.get(c) ?? 0) + r.text.length)
    }
  }
  let best: string | undefined
  let max = 0
  for (const [c, w] of weight) {
    if (w > max) {
      best = c
      max = w
    }
  }
  return best
}

/** Readable colors of a node (solid fill / dominant text color / stroke); pictures only expose stroke. */
function nodeColors(n: RenderNode): Pick<NodeInfo, 'fill' | 'textColor' | 'strokeColor'> {
  const out: Pick<NodeInfo, 'fill' | 'textColor' | 'strokeColor'> = {}
  if (n.type === 'shape' || n.type === 'text') {
    const s = n as ShapeRenderNode
    if (s.fill.kind === 'solid') {
      const c = hex6(s.fill.color)
      if (c) out.fill = c
    }
    const stroke = hex6(s.stroke?.color)
    if (stroke) out.strokeColor = stroke
    const text = dominantTextColor(n)
    if (text) out.textColor = text
  } else if (n.type === 'picture') {
    const stroke = hex6((n as PictureRenderNode).stroke?.color)
    if (stroke) out.strokeColor = stroke
  }
  return out
}

/**
 * Collect node info (including nested group children). A child's box is in group-local
 * coordinates (ext/chExt scaling already baked into geometry at build time); here we add the
 * group offset to convert to absolute coordinates and set the inGroup flag. Direct children of a
 * top-level group also carry groupId (editable via the in-group pipeline); deeper nesting stays
 * read-only (the main process patches one level only).
 */
function collectNodeInfos(
  nodes: RenderNode[],
  ox = 0,
  oy = 0,
  parent?: { id: string; topLevel: boolean },
): NodeInfo[] {
  const out: NodeInfo[] = []
  for (const n of nodes) {
    const b = n.box
    const abs = {
      x: Math.round(ox + b.x),
      y: Math.round(oy + b.y),
      w: Math.round(b.w),
      h: Math.round(b.h),
    }
    const base: NodeInfo = {
      id: n.sourceId,
      type: n.type,
      text: nodeText(n),
      ...abs,
      rotation: b.rotationDeg,
      ...(parent ? { inGroup: true } : {}),
      ...(parent?.topLevel ? { groupId: parent.id } : {}),
      ...(n.decoration ? { locked: true } : {}),
      ...nodeColors(n),
    }
    const fontPt = nodeMaxFontPt(n)
    if (fontPt !== undefined) base.fontSizePt = fontPt
    out.push(base)
    if (n.type === 'group') {
      const g = n as GroupRenderNode
      out.push(...collectNodeInfos(g.children, abs.x, abs.y, { id: n.sourceId, topLevel: !parent }))
    }
  }
  return out
}

function preview(text: string, max = 50): string {
  const flat = text.replace(/\n/g, ' / ')
  return flat.length > max ? `${flat.slice(0, max)}…` : flat
}

function buildDeckOutline(slides: RenderSlide[], current: number, selectedIds: string[]): string {
  const canvas = slides[0] ? `Canvas ${slides[0].widthPx}×${slides[0].heightPx}px.` : ''
  const lines: string[] = [
    `The presentation has ${slides.length} pages; page ${current + 1} is currently shown. ${canvas}`,
    `(Page order is the current actual order and may differ from generation time or earlier conversation; the user's "page N" refers to this outline)`,
  ]
  if (selectedIds.length > 0) lines.push(`User selected elements: ${selectedIds.join(', ')}`)
  slides.forEach((slide, i) => {
    lines.push(`Page ${i + 1} (slideIndex=${i}):`)
    const infos = collectNodeInfos(slide.nodes)
    const fillCount = new Map<string, number>()
    for (const n of infos) {
      if (n.fill) fillCount.set(n.fill, (fillCount.get(n.fill) ?? 0) + 1)
    }
    const mainFills = [...fillCount.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([c, count]) => (count > 1 ? `${c}×${count}` : c))
    if (mainFills.length > 0) lines.push(`  main fills: ${mainFills.join(' ')}`)
    for (const n of infos) {
      lines.push(`  - ${n.id} | ${n.type}${n.text ? ` | "${preview(n.text)}"` : ''}`)
    }
  })
  lines.push('(Use read_slide to see element positions/sizes/colors)')
  return lines.join('\n')
}

/**
 * Element inventory of one slide as read_slide reports it (ids + geometry + colors + text).
 * Shared by the read_slide tool and the post-generation layout QC pass (slide-qc.ts), so the
 * QC model maps screenshot pixels back to the same ids/coordinates the edit tools accept.
 */
export function formatSlideDump(slide: RenderSlide): string {
  const infos = collectNodeInfos(slide.nodes)
  const parts = infos.map((n) => {
    const flags = [
      n.groupId
        ? `in group ${n.groupId} (directly editable)`
        : n.inGroup
          ? 'nested in a sub-group (read-only; ungroup_element the outer group to edit)'
          : '',
      n.locked ? 'layout decoration (read-only)' : '',
    ]
      .filter(Boolean)
      .join(' ')
    const rot = n.rotation ? ` rotation ${Math.round(n.rotation)}°` : ''
    const font = n.fontSizePt ? ` font ${n.fontSizePt}pt` : ''
    const colors = [
      n.fill ? `fill${n.fill}` : '',
      n.textColor ? `text${n.textColor}` : '',
      n.strokeColor ? `stroke${n.strokeColor}` : '',
    ]
      .filter(Boolean)
      .join(' ')
    const head = `${n.id} | ${n.type}${flags ? ` | ${flags}` : ''} | pos(${n.x},${n.y}) size ${n.w}×${n.h}${rot}${font}${colors ? ` | ${colors}` : ''}`
    return n.text ? `${head}\n${n.text}` : `${head} | (no text)`
  })
  const colorlessTypes = [
    ...new Set(
      infos
        .filter((n) => !n.fill && !n.textColor && !n.strokeColor)
        .filter((n) => n.type === 'picture' || n.type === 'chart')
        .map((n) => n.type),
    ),
  ]
  const colorNote = colorlessTypes.length
    ? `\n(${colorlessTypes.join('/')} colors not available)`
    : ''
  return `Canvas ${slide.widthPx}×${slide.heightPx}px\n${parts.join('\n---\n') || '(no elements on this page)'}${colorNote}`
}

export function createSlidesSkill(
  access: DeckAccess,
  presentationFlags = {
    planning: true,
    verifiedCompletion: true,
    visualReview: true,
    autoCorrection: true,
  },
): AgentSkill {
  // The HTML pipeline was already used in this conversation → later calls without an explicit mode default to append.
  // Safety net for when the AI ignores the "pass all pages at once" constraint: separate calls no longer overwrite each other (P0-1).
  const state: SkillState = {}
  const scopedTools = new Set([
    'set_element_text',
    'set_element_style',
    'set_element_transform',
    'execute_layout_script',
    'execute_slide_script',
    'set_element_fill',
    'set_element_stroke',
    'add_text_box',
    'set_slide_background',
    'set_speaker_notes',
    'delete_element',
  ])
  const enrolledTargets = new Map<string, Map<string, string>>()
  const controller =
    access.taskReviewAdapter && presentationFlags.verifiedCompletion
      ? createSlidesTaskController({
          enroll: async (calls, currentContract, signal) => {
            signal?.throwIfAborted()
            const canonical = new Set([
              'set_element_style',
              'set_element_transform',
              'set_element_fill',
              'set_element_stroke',
              'set_slide_background',
            ])
            if (!calls.length || calls.some(({ name }) => !canonical.has(name)))
              return { kind: 'bypass' }
            if (!access.getAcceptanceAuthorityLease || !access.inspectAcceptanceAuthority)
              return { kind: 'bypass' }
            const lease = await access.getAcceptanceAuthorityLease()
            const affectedSlides = [
              ...new Set(calls.map((call) => Number(call.input.slideIndex) + 1)),
            ]
            if (affectedSlides.some((slide) => !Number.isSafeInteger(slide) || slide < 1))
              return { kind: 'clarify', question: 'presentation_scope_required' }
            const sourceTargets = calls.flatMap((call) =>
              typeof call.input.sourceId === 'string'
                ? [{ slide: Number(call.input.slideIndex) + 1, sourceId: call.input.sourceId }]
                : [],
            )
            const raw = (await access.inspectAcceptanceAuthority({
              affectedSlides,
              referenceSlides: [],
              expectedDocumentToken: lease.documentToken,
              expectedSessionToken: lease.sessionToken,
              expectedRevision: lease.revision,
              leaseToken: lease.leaseToken,
              sourceTargets,
            })) as SlidesAcceptanceAuthority & { sourceTargetTokens?: Record<string, string> }
            const sourceTargetTokens = raw.sourceTargetTokens ?? {}
            const { sourceTargetTokens: _resolved, ...authorityValue } = raw
            const authority = authorityValue as SlidesAcceptanceAuthority
            const taskId = `task-${crypto.randomUUID().replaceAll('-', '')}`
            enrolledTargets.set(taskId, new Map(Object.entries(sourceTargetTokens)))
            while (enrolledTargets.size > 8) {
              const oldest = enrolledTargets.keys().next().value as string | undefined
              if (!oldest) break
              enrolledTargets.delete(oldest)
            }
            const compiled = compileCanonicalSlidesCalls({
              calls,
              authority,
              sourceTargetTokens,
              taskId,
            })
            if (!('kind' in compiled) && !presentationFlags.planning)
              return { ...compiled, plan: undefined }
            return compiled
          },
          reviewAdapter: {
            ...access.taskReviewAdapter,
            correct: async (intents, authority, signal) => {
              if (
                !access.executePresentationOperation ||
                !access.taskReviewAdapter?.isCurrent(authority)
              )
                throw new Error('stale_authority')
              if (intents.length !== 1) throw new Error('unsupported_correction')
              const intent = intents[0]!
              if (intent.roleOrTarget.kind !== 'target') throw new Error('unsupported_correction')
              const targetToken = intent.roleOrTarget.targetToken
              const resolved = [...(enrolledTargets.get(authority.taskId ?? '') ?? [])].find(
                ([, token]) => token === targetToken,
              )
              if (!resolved) throw new Error('unsupported_correction')
              const [scope] = resolved
              const separator = scope.indexOf(':')
              const slideIndex = Number(scope.slice(0, separator)) - 1
              const sourceId = scope.slice(separator + 1)
              const transactionId = `slides-correction-${crypto.randomUUID().replaceAll('-', '')}`
              const request: GeometryFamilyTransactionRequest =
                intent.property === 'fill_color'
                  ? {
                      transactionId,
                      slideIndex,
                      operations: [
                        {
                          kind: 'set_fill',
                          sourceId,
                          fill: { kind: 'solid', color: String(intent.value) },
                        },
                      ],
                    }
                  : (() => {
                      // Geometry needs the complete authoritative box and text style needs the
                      // complete paragraph structure; acceptance facts intentionally omit both.
                      throw new Error('unsupported_correction')
                    })()
              const opened = await access.beginTaskCorrectionHistory?.()
              let execution: TextFamilyExecutionResult
              let correctionRollbackId: string | undefined
              try {
                execution = await access.executePresentationOperation(request, signal)
              } finally {
                if (opened) correctionRollbackId = await access.finishTaskCorrectionHistory?.()
              }
              if (
                execution.receipt.status !== 'applied' &&
                execution.receipt.status !== 'unchanged'
              )
                throw new Error('correction_not_applied')
              if (
                execution.receipt.status === 'applied' &&
                (execution.receipt.mutatedTargets?.length !== 1 ||
                  execution.receipt.mutatedTargets[0] !== targetToken)
              )
                throw new Error('correction_scope_mismatch')
              return {
                mutationReceiptId: execution.receipt.transactionId,
                applied: execution.receipt.status === 'applied',
                ...(correctionRollbackId ? { rollbackId: correctionRollbackId } : {}),
                correctedCheckIds: [intent.checkId],
              }
            },
          },
          afterTaskReview: access.onTaskReviewComplete,
          flags: presentationFlags,
        })
      : undefined
  const executionAccess: DeckAccess = controller
    ? {
        ...access,
        executePresentationOperation: async (request, signal) => {
          if (!access.executePresentationOperation)
            throw new Error('Canonical presentation transactions are unavailable')
          const result = await access.executePresentationOperation(request, signal)
          controller.recordMutation({
            transactionId: result.receipt.transactionId,
            status: result.receipt.status,
            mutatedTargetTokens:
              result.receipt.status === 'applied' ? [...(result.receipt.mutatedTargets ?? [])] : [],
          })
          return result
        },
      }
    : access
  return {
    id: 'slides',
    systemPrompt: AGENT_SYSTEM_PROMPT,
    get tools() {
      return access.getSelectionScope?.()
        ? TOOLS.filter((tool) => scopedTools.has(tool.name))
        : TOOLS
    },
    buildContext: () =>
      access.getSelectionScope?.()
        ? `<selection scope>\n${selectionScopeSummary(access.getSelectionScope()!)}. This scope is immutable and enforced by the host.\n</selection scope>`
        : `<deck outline>\n${buildDeckOutline(access.getSlides(), access.getCurrent(), access.getSelectedIds())}\n</deck outline>`,
    reviewFinalResponse: reviewSlidesFinalResponse,
    ...(controller ? { presentation: controller.hooks } : {}),
    executeTool: (call, signal) => executeTool(executionAccess, call, state, signal),
  }
}

interface SkillState {
  fallbackInvocationIds?: WeakMap<AgentToolCall, string>
  /** A web_search ran in this conversation — unlocks dataSource:'search' in the figure gate */
  webSearched?: boolean
  /** Most recently available Style Skill (used by save_style_template) */
  lastStyleSkill?: string
  /** Topic associated with the most recently available Style Skill */
  lastTopic?: string
}

const fail = (summary: string, output: string) => ({
  output,
  isError: true,
  mutated: false,
  summary,
})

// ── Figure-provenance gate ────────────────────────────────────
// Prompt rules ("search before writing data") did not stop invented numbers being
// delivered as fact, so provenance is enforced at the tool layer: chart data and
// data-dense briefs must declare a dataSource, 'search' is only accepted after a
// real web_search in this conversation, and 'sample' figures must be disclosed.

/**
 * Returns an error message when the declared dataSource does not justify the figures
 * this call carries, null when the call may proceed.
 */
function dataSourceGateError(call: AgentToolCall, state: SkillState | undefined): string | null {
  const src = String(call.input.dataSource ?? '')
  if (src === 'user' || src === 'document' || src === 'sample') return null
  if (src === 'search') {
    if (state?.webSearched) return null
    return (
      "dataSource is 'search' but no web_search has run in this conversation. " +
      'Run web_search first and build the figures from the results (cite them to the user), ' +
      "or declare 'user'/'document' if the figures actually came from the user or this deck."
    )
  }
  return (
    'This call carries specific figures, so dataSource is required: ' +
    "'user' (figures supplied by the user or attachments), 'document' (read from this deck), " +
    "'search' (from web_search results — run it first), or 'sample' (illustrative placeholders; " +
    'you must tell the user they are NOT real data). Never present invented numbers as facts.'
  )
}

/** Appended to a successful tool output when the model declared the figures illustrative. */
const SAMPLE_DATA_NOTE =
  '\nNOTE: dataSource is "sample" — you MUST tell the user these figures are illustrative placeholders, not real data, and offer to research real numbers with web_search.'

/** Template placeholder strings that must never survive into a delivered page */
const PLACEHOLDER_PATTERNS: RegExp[] = [
  /copy paste fonts\. choose the only option to retain text/i,
  /click to edit master title style/i,
  /\btext here\b/i,
  /supporting text here/i,
  /lorem ipsum/i,
]

/**
 * Cheap post-generation audit of one page's HTML: flags leftover template
 * placeholders and near-empty pages so the model is forced to redo them instead of
 * delivering filler as success. Returns a short reason, or null when the page looks fine.
 */
export function auditPageHtml(html: string): string | null {
  const text = html
    .replace(/<(style|script)\b[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
  for (const re of PLACEHOLDER_PATTERNS) {
    const m = re.exec(text)
    if (m) return `contains template placeholder text "${m[0]}"`
  }
  if (text.length < 10) return 'has almost no text content'
  return null
}

async function executeTool(
  access: DeckAccess,
  call: AgentToolCall,
  state?: SkillState,
  signal?: AbortSignal,
) {
  signal?.throwIfAborted()
  const activeScope = access.getSelectionScope?.()
  if (activeScope) {
    const allowed = new Set([
      'set_element_text',
      'set_element_style',
      'set_element_transform',
      'execute_layout_script',
      'execute_slide_script',
      'set_element_fill',
      'set_element_stroke',
      'add_text_box',
      'set_slide_background',
      'set_speaker_notes',
      'delete_element',
    ])
    if (!allowed.has(call.name))
      return fail(call.name, 'selection_scope_conflict: tool is unavailable for a scoped edit')
  }
  if (!('invocationId' in call) || !call.invocationId) {
    state ??= {}
    state.fallbackInvocationIds ??= new WeakMap()
    let invocationId = state.fallbackInvocationIds.get(call)
    if (!invocationId) {
      invocationId = globalThis.crypto.randomUUID()
      state.fallbackInvocationIds.set(call, invocationId)
    }
    Object.defineProperty(call, 'invocationId', {
      configurable: false,
      enumerable: false,
      writable: false,
      value: invocationId,
    })
  }
  const unguardedAccess = access
  access = {
    ...unguardedAccess,
    applySlide: (slideIndex, updated) => {
      signal?.throwIfAborted()
      unguardedAccess.applySlide(slideIndex, updated)
    },
    applyDeck: (updatedSlides, goTo) => {
      signal?.throwIfAborted()
      unguardedAccess.applyDeck(updatedSlides, goTo)
    },
  }
  const slides = access.getSlides()
  if (UNSUPPORTED_CLOUD_TOOLS.has(call.name)) {
    return fail(
      call.name,
      'unsupported_feature: This feature is not available in the current WisWork development version.',
    )
  }
  switch (call.name) {
    case 'get_deck_context':
      return {
        output: buildDeckOutline(slides, access.getCurrent(), access.getSelectedIds()),
        mutated: false,
        summary: t('aiSumDeckContext'),
      }

    case 'read_slide': {
      const idx = Number(call.input.slideIndex)
      const slide = slides[idx]
      if (!slide)
        return fail(t('aiFailReadSlide'), `slideIndex out of range (0-${slides.length - 1})`)
      return {
        output: formatSlideDump(slide),
        mutated: false,
        summary: t('aiSumReadSlide', { n: idx + 1 }),
      }
    }

    case 'set_element_text': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      const slide = slides[idx]
      if (!slide)
        return fail(t('aiFailEditText'), `slideIndex out of range (0-${slides.length - 1})`)
      const paragraphs = toEditParagraphs(call.input.paragraphs)
      if (!paragraphs) return fail(t('aiFailEditText'), 'paragraphs must be a non-empty array')
      const target = resolveEditTarget(slide, sourceId)
      const terr = targetError(target, sourceId, idx + 1)
      if (terr || !target || 'nested' in target) return fail(t('aiFailEditText'), terr!)
      if (!access.executePresentationOperation)
        return fail(t('aiFailEditText'), 'Canonical presentation transactions are unavailable')
      const transactionId = await textToolTransactionId(call)
      const execution = await access.executePresentationOperation(
        {
          transactionId,
          slideIndex: idx,
          sourceId,
          operation: { kind: 'set_text', paragraphs },
        },
        signal,
      )
      const outcome = textFamilyReceiptOutcome(execution.receipt)
      if (execution.authoritativeState === 'reload_required') {
        return {
          output:
            'The text change was applied, but the editor could not refresh authoritative state. Reloading is required before more edits.',
          mutated: true,
          stopToolBatch: true,
          summary: t('aiSumEditText', { n: idx + 1 }),
        }
      }
      if (!outcome.ok) {
        return {
          output:
            outcome.detail === 'write_state_uncertain'
              ? 'The text change may be partially applied. Inspect the slide before retrying.'
              : `The text change was not applied (${outcome.detail ?? 'write_not_applied'}).`,
          isError: true,
          mutated: outcome.mutated,
          summary: t('aiFailEditText'),
        }
      }
      return {
        output: outcome.mutated
          ? `Replaced the text of element ${sourceId} on page ${idx + 1} (${paragraphs.length} paragraphs).`
          : `Element ${sourceId} on page ${idx + 1} already had the requested text.`,
        mutated: outcome.mutated,
        summary: t('aiSumEditText', { n: idx + 1 }),
      }
    }

    case 'set_element_style': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      const slide = slides[idx]
      if (!slide) return fail(t('aiFailStyle'), `slideIndex out of range (0-${slides.length - 1})`)
      const target = resolveEditTarget(slide, sourceId)
      const terr = targetError(target, sourceId, idx + 1)
      if (terr || !target || 'nested' in target) return fail(t('aiFailStyle'), terr!)
      const node = target.node
      if (!(node.type === 'text' || node.type === 'shape')) {
        return fail(t('aiFailStyle'), `Element ${sourceId} (${node.type}) has no editable text`)
      }
      const cur = nodeToParagraphs(node as ShapeRenderNode)
      if (!cur.length) return fail(t('aiFailStyle'), 'This element has no text to format')
      const ov = call.input as SlideStylePatch
      const paragraphs = mergeStyleIntoParagraphs(cur, ov)
      if (!access.executePresentationOperation)
        return fail(t('aiFailStyle'), 'Canonical presentation transactions are unavailable')
      const execution = await access.executePresentationOperation(
        {
          transactionId: await textToolTransactionId(call),
          slideIndex: idx,
          sourceId,
          operation: { kind: 'set_text', paragraphs },
        },
        signal,
      )
      const outcome = textFamilyReceiptOutcome(execution.receipt)
      if (execution.authoritativeState === 'reload_required') {
        return {
          output:
            'The formatting change was applied, but the editor could not refresh authoritative state. Reloading is required before more edits.',
          mutated: true,
          stopToolBatch: true,
          summary: t('aiSumStyle', { n: idx + 1 }),
        }
      }
      if (!outcome.ok) {
        return {
          output: outcome.mutated
            ? 'The formatting change may be partially applied. Inspect the slide before retrying.'
            : `The formatting change was not applied (${outcome.detail ?? 'write_not_applied'}).`,
          isError: true,
          mutated: outcome.mutated,
          summary: t('aiFailStyle'),
        }
      }
      return {
        output: outcome.mutated
          ? `Updated the formatting of element ${sourceId} on page ${idx + 1}.`
          : `Element ${sourceId} on page ${idx + 1} already had the requested formatting.`,
        mutated: outcome.mutated,
        summary: t('aiSumStyle', { n: idx + 1 }),
      }
    }

    case 'set_element_transform': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      const slide = slides[idx]
      if (!slide)
        return fail(t('aiFailTransform'), `slideIndex out of range (0-${slides.length - 1})`)
      const target = resolveEditTarget(slide, sourceId)
      const terr = targetError(target, sourceId, idx + 1)
      if (terr || !target || 'nested' in target) return fail(t('aiFailTransform'), terr!)
      const b = target.node.box
      // Group-child render boxes are group-local; the tool takes absolute px, so convert both ways via the group origin
      const origin = target.groupOrigin ?? { x: 0, y: 0 }
      const inp = call.input as {
        x?: number
        y?: number
        w?: number
        h?: number
        rotationDeg?: number
      }
      const x = typeof inp.x === 'number' ? inp.x : origin.x + b.x
      const y = typeof inp.y === 'number' ? inp.y : origin.y + b.y
      const w = typeof inp.w === 'number' ? inp.w : b.w
      const h = typeof inp.h === 'number' ? inp.h : b.h
      const rotation = normalizePresentationRotation(
        typeof inp.rotationDeg === 'number' ? inp.rotationDeg : b.rotationDeg,
      )
      if (![x, y, w, h].every(Number.isFinite) || w <= 0 || h <= 0)
        return fail(t('aiFailTransform'), 'Transform geometry is invalid')
      if (!access.executePresentationOperation)
        return fail(t('aiFailTransform'), 'Canonical presentation transactions are unavailable')
      const execution = await access.executePresentationOperation(
        {
          transactionId: await geometryToolTransactionId(call),
          slideIndex: idx,
          operations: [
            {
              sourceId,
              geometry: {
                x: geometryPxToPoints(x, slide.scale),
                y: geometryPxToPoints(y, slide.scale),
                width: geometryPxToPoints(w, slide.scale),
                height: geometryPxToPoints(h, slide.scale),
                rotation,
              },
            },
          ],
        },
        signal,
      )
      const outcome = textFamilyReceiptOutcome(execution.receipt)
      if (execution.authoritativeState === 'reload_required') {
        return {
          output:
            'The transform was applied, but the editor could not refresh authoritative state. Reloading is required before more edits.',
          mutated: true,
          stopToolBatch: true,
          summary: t('aiSumTransform', { n: idx + 1 }),
        }
      }
      if (!outcome.ok) {
        return {
          output: outcome.mutated
            ? 'The transform may be partially applied. Inspect the slide before retrying.'
            : `The transform was not applied (${outcome.detail ?? 'write_not_applied'}).`,
          isError: true,
          mutated: outcome.mutated,
          summary: t('aiFailTransform'),
        }
      }
      const updated = access.getSlides()[idx] ?? slide
      const afterTarget = resolveEditTarget(updated, sourceId)
      const after = afterTarget && !('nested' in afterTarget) ? afterTarget : null
      const nb = after
        ? {
            ...after.node.box,
            x: (after.groupOrigin?.x ?? 0) + after.node.box.x,
            y: (after.groupOrigin?.y ?? 0) + after.node.box.y,
          }
        : undefined
      const boxStr = nb
        ? `New geometry: pos(${Math.round(nb.x)},${Math.round(nb.y)}) size ${Math.round(nb.w)}×${Math.round(nb.h)}.`
        : ''
      const issues = auditSlideLayout(updated)
      const auditStr = issues.length
        ? `\n⚠️ The layout audit found ${issues.length} issue(s) on this page:\n${issues.map((s) => `- ${s}`).join('\n')}\nFor multi-element layout adjustments switch to execute_slide_script (it reads every element's real geometry and applies atomically).`
        : ''
      return {
        output: outcome.mutated
          ? `Adjusted the position/size of element ${sourceId} on page ${idx + 1}. ${boxStr}${auditStr}`
          : `Element ${sourceId} on page ${idx + 1} already had the requested geometry.${auditStr}`,
        mutated: outcome.mutated,
        summary: t('aiSumTransform', { n: idx + 1 }),
      }
    }

    // execute_layout_script is a legacy alias (avoids breaking existing sessions/prompts); both share the same logic
    case 'execute_layout_script':
    case 'execute_slide_script': {
      const idx = Number(call.input.slideIndex)
      const slide = slides[idx]
      if (!slide) return fail(t('aiFailScript'), `slideIndex out of range (0-${slides.length - 1})`)
      const code = String(call.input.code ?? '').trim()
      if (!code) return fail(t('aiFailScript'), 'code must not be empty')
      const infos = collectNodeInfos(slide.nodes)
      const r = runLayoutScript(code, infos, { w: slide.widthPx, h: slide.heightPx })
      const logsStr = r.logs.length ? `\nlog output:\n${r.logs.join('\n')}` : ''
      if (r.error) {
        return fail(
          t('aiFailScript'),
          `Script execution error: ${r.error}${logsStr}\n(You can fix the script and retry; see els for the element list and geometry)`,
        )
      }
      const returnedStr = r.returned !== undefined ? `\nScript returned: ${r.returned}` : ''
      if (r.ops.length === 0 && r.edits.length === 0) {
        return {
          output: `Script finished but called no edit primitives (addText/delete/setBox/moveBy/setText/setStyle/setFill/setStroke); the page was not modified.${returnedStr}${logsStr}`,
          mutated: false,
          summary: t('aiSumScriptNoop', { n: idx + 1 }),
        }
      }
      // Every primitive exposed by this DSL now belongs to an enrolled canonical family. Compile
      // the complete script before preparing any target so an unsupported paint/text target fails
      // closed without a proposal, mutation, or partially executed prefix.
      if (r.ops.length > 0 || r.edits.length > 0) {
        if (!access.executePresentationOperation)
          return fail(t('aiFailScript'), 'Canonical presentation transactions are unavailable')
        const operations: GeometryFamilyTransactionRequest['operations'][number][] = r.ops.map(
          (op) => {
            const target = resolveEditTarget(slide, op.id)
            if (!target || 'nested' in target)
              throw new TypeError(`Geometry target ${op.id} is no longer editable`)
            return {
              sourceId: op.id,
              geometry: {
                x: geometryPxToPoints(op.x, slide.scale),
                y: geometryPxToPoints(op.y, slide.scale),
                width: geometryPxToPoints(op.w, slide.scale),
                height: geometryPxToPoints(op.h, slide.scale),
                rotation: normalizePresentationRotation(op.rotation),
              },
            }
          },
        )
        const paragraphState = new Map<string, EditParagraph[]>()
        const canonicalParagraphs = (paragraphs: readonly EditParagraph[]) => {
          const paragraphKeys = new Set(['runs', 'align'])
          const runKeys = new Set([
            'text',
            'bold',
            'italic',
            'underline',
            'fontSize',
            'fontFamily',
            'color',
          ])
          if (
            paragraphs.some(
              (paragraph) =>
                Object.keys(paragraph).some((key) => !paragraphKeys.has(key)) ||
                paragraph.align === 'justify' ||
                paragraph.runs.some((run) => Object.keys(run).some((key) => !runKeys.has(key))),
            )
          )
            return null
          return paragraphs.map((paragraph) => ({
            runs: paragraph.runs.map((run) => ({ ...run })),
            ...(paragraph.align === 'left' ||
            paragraph.align === 'center' ||
            paragraph.align === 'right'
              ? { align: paragraph.align }
              : {}),
          }))
        }
        const generatedIds = new Set<string>()
        for (const edit of r.edits) {
          if (edit.kind === 'add_text') {
            const paragraphs = canonicalParagraphs(edit.paragraphs)
            if (!paragraphs)
              return fail(t('aiFailScript'), `addText("${edit.id}"): unsupported rich formatting`)
            const text = paragraphs
              .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
              .join('\n')
            operations.push({
              kind: 'add_text_box',
              clientId: edit.id,
              text,
              geometry: {
                x: geometryPxToPoints(edit.geometry.x, slide.scale),
                y: geometryPxToPoints(edit.geometry.y, slide.scale),
                width: geometryPxToPoints(edit.geometry.w, slide.scale),
                height: geometryPxToPoints(edit.geometry.h, slide.scale),
                rotation: normalizePresentationRotation(edit.geometry.rotation),
              },
            })
            generatedIds.add(edit.id)
            paragraphState.set(edit.id, edit.paragraphs)
            operations.push({
              createdByClientId: edit.id,
              kind: 'set_text',
              paragraphs,
            })
            continue
          }
          if (edit.kind === 'delete') {
            if (generatedIds.has(edit.id)) {
              operations.push({ createdByClientId: edit.id, kind: 'delete_element' })
              continue
            }
            const target = resolveEditTarget(slide, edit.id)
            if (!target || 'nested' in target || target.groupId)
              return fail(t('aiFailScript'), `Delete target ${edit.id} is no longer editable`)
            operations.push({ sourceId: edit.id, kind: 'delete_element' })
            continue
          }
          const generated = generatedIds.has(edit.id)
          const target = generated ? undefined : resolveEditTarget(slide, edit.id)
          const existingTarget = target && !('nested' in target) ? target : undefined
          if (!generated && !existingTarget)
            return fail(t('aiFailScript'), `Target ${edit.id} is no longer editable`)
          const targetReference = generated ? { createdByClientId: edit.id } : { sourceId: edit.id }
          if (edit.kind === 'fill') {
            if (
              !generated &&
              !(existingTarget!.node.type === 'text' || existingTarget!.node.type === 'shape')
            )
              return fail(
                t('aiFailScript'),
                `setFill("${edit.id}"): element does not support a durable fill`,
              )
            operations.push({
              ...targetReference,
              kind: 'set_fill',
              fill:
                edit.fill === 'none'
                  ? { kind: 'none' }
                  : { kind: 'solid', color: edit.fill.toUpperCase() },
            })
          } else if (edit.kind === 'stroke') {
            if (
              !generated &&
              !(
                existingTarget!.node.type === 'text' ||
                existingTarget!.node.type === 'shape' ||
                existingTarget!.node.type === 'picture'
              )
            )
              return fail(
                t('aiFailScript'),
                `setStroke("${edit.id}"): element does not support a durable stroke`,
              )
            operations.push({
              ...targetReference,
              kind: 'set_stroke',
              stroke: edit.stroke && {
                color: edit.stroke.color.toUpperCase(),
                width: edit.stroke.widthPt,
              },
            })
          } else if (edit.kind === 'text') {
            if (
              !generated &&
              !(existingTarget!.node.type === 'text' || existingTarget!.node.type === 'shape')
            )
              return fail(
                t('aiFailScript'),
                `setText("${edit.id}"): element does not support text editing`,
              )
            paragraphState.set(
              edit.id,
              edit.paragraphs.map((paragraph) => ({
                ...paragraph,
                runs: paragraph.runs.map((run) => ({ ...run })),
              })),
            )
            const paragraphs = canonicalParagraphs(paragraphState.get(edit.id)!)
            if (!paragraphs)
              return fail(
                t('aiFailScript'),
                `setText("${edit.id}"): rich paragraph formatting is not supported by atomic scripts`,
              )
            operations.push({
              ...targetReference,
              kind: 'set_text',
              paragraphs,
            })
          } else if (edit.kind === 'style') {
            if (
              !generated &&
              !(existingTarget!.node.type === 'text' || existingTarget!.node.type === 'shape')
            )
              return fail(t('aiFailScript'), `setStyle("${edit.id}"): text element not found`)
            const current =
              paragraphState.get(edit.id) ??
              nodeToParagraphs(existingTarget!.node as ShapeRenderNode)
            if (!current.length)
              return fail(
                t('aiFailScript'),
                `setStyle("${edit.id}"): this element has no text to format`,
              )
            const paragraphs = mergeStyleIntoParagraphs(current, edit.style)
            paragraphState.set(edit.id, paragraphs)
            const canonical = canonicalParagraphs(paragraphs)
            if (!canonical)
              return fail(
                t('aiFailScript'),
                `setStyle("${edit.id}"): existing rich paragraph formatting cannot be preserved atomically`,
              )
            operations.push({
              ...targetReference,
              kind: 'set_text',
              paragraphs: canonical,
            })
          } else {
            return fail(t('aiFailScript'), 'Unsupported script operation')
          }
        }
        const execution = await access.executePresentationOperation(
          {
            transactionId: await geometryToolTransactionId(call),
            slideIndex: idx,
            operations,
          },
          signal,
        )
        const outcome = textFamilyReceiptOutcome(execution.receipt)
        if (execution.authoritativeState === 'reload_required') {
          return {
            output:
              'The layout changes were applied, but authoritative state could not be refreshed. Reloading is required before more edits.',
            mutated: true,
            stopToolBatch: true,
            summary: t('aiSumScript', { n: idx + 1 }),
          }
        }
        if (!outcome.ok) {
          return {
            output: outcome.mutated
              ? 'The layout changes may be partially applied. Inspect the slide before retrying.'
              : `The layout changes were not applied (${outcome.detail ?? 'write_not_applied'}).`,
            isError: true,
            mutated: outcome.mutated,
            summary: t('aiFailScript'),
          }
        }
        const refreshed = access.getSlides()[idx] ?? slide
        const audit = formatAudit(auditSlideLayout(refreshed))
        const counts = {
          add: r.edits.filter((edit) => edit.kind === 'add_text').length,
          delete: r.edits.filter((edit) => edit.kind === 'delete').length,
          text: r.edits.filter((edit) => edit.kind === 'text').length,
          style: r.edits.filter((edit) => edit.kind === 'style').length,
          fill: r.edits.filter((edit) => edit.kind === 'fill').length,
          stroke: r.edits.filter((edit) => edit.kind === 'stroke').length,
        }
        const parts: string[] = []
        if (counts.add) parts.push(`add ${counts.add} text box(es)`)
        if (counts.delete) parts.push(`delete ${counts.delete} element(s)`)
        if (r.ops.length) parts.push(`layout ${r.ops.length} element(s)`)
        if (counts.text) parts.push(`text ${counts.text} item(s)`)
        if (counts.style) parts.push(`style ${counts.style} item(s)`)
        if (counts.fill) parts.push(`fill ${counts.fill} item(s)`)
        if (counts.stroke) parts.push(`stroke ${counts.stroke} item(s)`)
        const createdTargetSummary =
          execution.receipt.status === 'applied' && execution.receipt.createdTargets?.length
            ? r.edits
                .filter((edit) => edit.kind === 'add_text')
                .map((edit) => {
                  const operationIndex = operations.findIndex(
                    (operation) =>
                      operation.kind === 'add_text_box' && operation.clientId === edit.id,
                  )
                  const target =
                    execution.receipt.status === 'applied'
                      ? execution.receipt.createdTargets?.find(
                          (created) => created.clientId === `op-${operationIndex + 1}`,
                        )
                      : undefined
                  return target ? `${edit.id}=${target.elementId}` : null
                })
                .filter((value): value is string => value !== null)
                .join(', ')
            : ''
        return {
          output: outcome.mutated
            ? `Script applied: ${parts.join(', ')}.${createdTargetSummary ? ` Created targets: ${createdTargetSummary}.` : ''}${returnedStr}${logsStr}${audit ? `\n${audit}` : ''}`
            : `Script already matched the requested state.${returnedStr}${logsStr}${audit ? `\n${audit}` : ''}`,
          mutated: outcome.mutated,
          summary: t('aiSumScript', { n: idx + 1 }),
        }
      }
      // ── Dispatch: geometry applied atomically once via batchEditTransform, the rest serially in script order,
      //   each step using the returned new slide as the next step's current state. One failure doesn't crash the whole run; report faithfully.
      // Balanced with the end in the finally below; an unbalanced pair would leave
      // the session mid-batch, where undo/redo refuse to run
      const batchOpened = (await window.slidesApi.beginHistoryBatch?.()) === true
      try {
        signal?.throwIfAborted()
        let current = slide
        const failures: string[] = []
        let boxApplied = 0
        const topOps = r.ops.filter((op) => !op.groupId)
        const grpOps = r.ops.filter((op) => op.groupId)
        if (topOps.length > 0) {
          const updated = await window.slidesApi.batchEditTransform({
            slideIndex: idx,
            fitWidthPx: access.fitWidthPx,
            items: topOps.map((op) => ({
              sourceId: op.id,
              xPx: op.x,
              yPx: op.y,
              wPx: op.w,
              hPx: op.h,
              rotationDeg: op.rotation,
            })),
          })
          signal?.throwIfAborted()
          if (updated) {
            current = updated
            access.applySlide(idx, updated)
            boxApplied = topOps.length
          } else {
            failures.push(
              `Batch geometry apply failed (${topOps.length} items; some elements may no longer exist on this page)`,
            )
          }
        }
        // Group children go through the per-element in-group transform IPC (abs px → group-local px via the group origin)
        for (const op of grpOps) {
          const gnode = findNodeById(current.nodes, op.groupId!)
          const updated = gnode
            ? await window.slidesApi.editTransform({
                slideIndex: idx,
                sourceId: op.id,
                groupId: op.groupId!,
                xPx: op.x - Math.round(gnode.box.x),
                yPx: op.y - Math.round(gnode.box.y),
                wPx: op.w,
                hPx: op.h,
                rotationDeg: op.rotation,
                fitWidthPx: access.fitWidthPx,
              })
            : null
          signal?.throwIfAborted()
          if (!updated) {
            failures.push(`setBox("${op.id}"): geometry apply inside group ${op.groupId} failed`)
            continue
          }
          current = updated
          access.applySlide(idx, updated)
          boxApplied += 1
        }
        const counts = { text: 0, style: 0, fill: 0, stroke: 0 }
        for (const e of r.edits) {
          const grp = e.groupId ? { groupId: e.groupId } : {}
          let updated: RenderSlide | null
          if (e.kind === 'text') {
            updated = await window.slidesApi.editText({
              slideIndex: idx,
              sourceId: e.id,
              paragraphs: e.paragraphs,
              ...grp,
            })
            signal?.throwIfAborted()
            if (!updated) {
              failures.push(
                `setText("${e.id}"): element not found or does not support text editing`,
              )
              continue
            }
          } else if (e.kind === 'style') {
            // Changing style = read current paragraphs (including earlier edit results), merge override fields, write back whole
            const node = findNodeById(current.nodes, e.id)
            if (!node || !(node.type === 'text' || node.type === 'shape')) {
              failures.push(`setStyle("${e.id}"): text element not found`)
              continue
            }
            const cur = nodeToParagraphs(node as ShapeRenderNode)
            if (!cur.length) {
              failures.push(`setStyle("${e.id}"): this element has no text to format`)
              continue
            }
            updated = await window.slidesApi.editText({
              slideIndex: idx,
              sourceId: e.id,
              paragraphs: mergeStyleIntoParagraphs(cur, e.style),
              ...grp,
            })
            signal?.throwIfAborted()
            if (!updated) {
              failures.push(`setStyle("${e.id}"): element does not support format editing`)
              continue
            }
          } else if (e.kind === 'fill') {
            updated = await window.slidesApi.editFill({
              slideIndex: idx,
              sourceId: e.id,
              fill: e.fill,
              ...grp,
            })
            signal?.throwIfAborted()
            if (!updated) {
              failures.push(`setFill("${e.id}"): element does not support fill`)
              continue
            }
          } else if (e.kind === 'stroke') {
            updated = await window.slidesApi.editStroke({
              slideIndex: idx,
              sourceId: e.id,
              stroke: e.stroke,
              ...grp,
            })
            signal?.throwIfAborted()
            if (!updated) {
              failures.push(`setStroke("${e.id}"): element does not support stroke`)
              continue
            }
          } else continue
          current = updated
          access.applySlide(idx, updated)
          counts[e.kind] += 1
        }
        const totalApplied = boxApplied + counts.text + counts.style + counts.fill + counts.stroke
        if (totalApplied === 0) {
          return fail(
            t('aiFailScript'),
            `All operations collected by the script failed to apply:\n${failures.map((f) => `- ${f}`).join('\n')}${returnedStr}${logsStr}`,
          )
        }
        const parts: string[] = []
        if (boxApplied > 0) parts.push(`layout ${boxApplied} element(s)`)
        if (counts.text > 0) parts.push(`text ${counts.text} item(s)`)
        if (counts.style > 0) parts.push(`style ${counts.style} item(s)`)
        if (counts.fill > 0) parts.push(`fill ${counts.fill} item(s)`)
        if (counts.stroke > 0) parts.push(`stroke ${counts.stroke} item(s)`)
        const failStr = failures.length
          ? `\n⚠️ ${failures.length} operation(s) failed (the rest took effect):\n${failures.map((f) => `- ${f}`).join('\n')}`
          : ''
        const issues = auditSlideLayout(current)
        return {
          output: `Applied the edit script to page ${idx + 1}: ${parts.join(', ')}.${returnedStr}${logsStr}${failStr}${formatAudit(issues)}`,
          mutated: true,
          summary: t('aiSumScript', { n: idx + 1, count: totalApplied }),
        }
      } finally {
        if (batchOpened) await window.slidesApi.endHistoryBatch?.()
      }
    }

    case 'set_element_fill': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailFill'), `slideIndex out of range (0-${slides.length - 1})`)
      const target = resolveEditTarget(slides[idx]!, sourceId)
      const terr = targetError(target, sourceId, idx + 1)
      if (terr || !target || 'nested' in target) return fail(t('aiFailFill'), terr!)
      if (!(target.node.type === 'text' || target.node.type === 'shape'))
        return fail(t('aiFailFill'), `Element ${sourceId} does not support a durable fill`)
      const rawFill = String(call.input.fill)
      const fill =
        rawFill === 'none'
          ? ({ kind: 'none' } as const)
          : /^#?[0-9a-fA-F]{6}$/.test(rawFill)
            ? ({ kind: 'solid', color: `#${rawFill.replace(/^#/, '').toUpperCase()}` } as const)
            : null
      if (!fill) return fail(t('aiFailFill'), 'fill must be #RRGGBB or none')
      if (!access.executePresentationOperation)
        return fail(t('aiFailFill'), 'Canonical presentation transactions are unavailable')
      const execution = await access.executePresentationOperation(
        {
          transactionId: await geometryToolTransactionId(call),
          slideIndex: idx,
          operations: [{ sourceId, kind: 'set_fill', fill }],
        },
        signal,
      )
      const outcome = textFamilyReceiptOutcome(execution.receipt)
      if (execution.authoritativeState === 'reload_required')
        return {
          output: 'The fill was applied, but reloading is required before more edits.',
          mutated: true,
          stopToolBatch: true,
          summary: t('aiSumFill', { n: idx + 1 }),
        }
      if (!outcome.ok)
        return {
          output: outcome.mutated
            ? 'The fill change may be partially applied. Inspect the slide before retrying.'
            : `The fill was not applied (${outcome.detail ?? 'write_not_applied'}).`,
          isError: true,
          mutated: outcome.mutated,
          summary: t('aiFailFill'),
        }
      return {
        output: outcome.mutated
          ? `Set the fill of element ${sourceId} on page ${idx + 1}.`
          : `Element ${sourceId} already had the requested fill.`,
        mutated: outcome.mutated,
        summary: t('aiSumFill', { n: idx + 1 }),
      }
    }

    case 'set_element_stroke': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailStroke'), `slideIndex out of range (0-${slides.length - 1})`)
      const remove = call.input.remove === true
      const stroke = remove
        ? null
        : { color: String(call.input.color ?? '#000000'), widthPt: Number(call.input.widthPt ?? 1) }
      const target = resolveEditTarget(slides[idx]!, sourceId)
      const terr = targetError(target, sourceId, idx + 1)
      if (terr || !target || 'nested' in target) return fail(t('aiFailStroke'), terr!)
      if (!(
        target.node.type === 'text' ||
        target.node.type === 'shape' ||
        target.node.type === 'picture'
      ))
        return fail(t('aiFailStroke'), `Element ${sourceId} does not support a durable stroke`)
      if (
        stroke &&
        (!/^#[0-9a-fA-F]{6}$/.test(stroke.color) ||
          !Number.isFinite(stroke.widthPt) ||
          stroke.widthPt < 0)
      )
        return fail(t('aiFailStroke'), 'stroke color/width is invalid')
      if (!access.executePresentationOperation)
        return fail(t('aiFailStroke'), 'Canonical presentation transactions are unavailable')
      const execution = await access.executePresentationOperation(
        {
          transactionId: await geometryToolTransactionId(call),
          slideIndex: idx,
          operations: [
            {
              sourceId,
              kind: 'set_stroke',
              stroke: stroke && { color: stroke.color.toUpperCase(), width: stroke.widthPt },
            },
          ],
        },
        signal,
      )
      const outcome = textFamilyReceiptOutcome(execution.receipt)
      if (execution.authoritativeState === 'reload_required')
        return {
          output: 'The stroke was applied, but reloading is required before more edits.',
          mutated: true,
          stopToolBatch: true,
          summary: t('aiSumStroke', { n: idx + 1 }),
        }
      if (!outcome.ok)
        return {
          output: outcome.mutated
            ? 'The stroke change may be partially applied. Inspect the slide before retrying.'
            : `The stroke was not applied (${outcome.detail ?? 'write_not_applied'}).`,
          isError: true,
          mutated: outcome.mutated,
          summary: t('aiFailStroke'),
        }
      return {
        output: `${remove ? 'Removed' : 'Set'} the stroke of element ${sourceId} on page ${idx + 1}.`,
        mutated: outcome.mutated,
        summary: t('aiSumStroke', { n: idx + 1 }),
      }
    }

    case 'web_search': {
      const query = String(call.input.query ?? '').trim()
      if (!query) return fail(t('aiFailWebSearch'), 'query must not be empty')
      const r = await window.slidesApi.webSearch(query, Number(call.input.maxResults) || 6)
      signal?.throwIfAborted()
      if (state) state.webSearched = true
      // output for the LLM: title+URL+summary (each summary truncated to 120 chars to stay lean)
      const SNIPPET_MAX = 120
      const lines: string[] = []
      if (r.answer) lines.push(`Direct answer: ${r.answer.slice(0, 300)}\n`)
      r.results.forEach((it, i) => {
        const snip =
          it.snippet.length > SNIPPET_MAX ? it.snippet.slice(0, SNIPPET_MAX) + '…' : it.snippet
        lines.push(`${i + 1}. ${it.title}\n   ${it.url}\n   ${snip}`)
      })
      // display side channel: link list (full title+URL for the UI, not in LLM context)
      const display: ToolDisplay = {
        kind: 'links',
        items: r.results.map((it) => ({ url: it.url, title: it.title })),
      }
      return {
        output: lines.join('\n') || '(no results)',
        mutated: false,
        summary: t('aiSumWebSearch', { query, count: r.results.length }),
        display,
      }
    }

    case 'image_search': {
      const query = String(call.input.query ?? '').trim()
      if (!query) return fail(t('aiFailImageSearch'), 'query must not be empty')
      const r = await window.slidesApi.imageSearch(query, Number(call.input.maxResults) || 8)
      signal?.throwIfAborted()
      // output for the LLM: keep the existing format (the LLM needs to read URLs into image_queries; format unchanged)
      const lines = r.images.map(
        (im, i) =>
          `${i + 1}. ${im.title || '(untitled)'} [${im.width ?? '?'}x${im.height ?? '?'}]\n   ${im.imageUrl}`,
      )
      // display side channel: image list (for UI thumbnails, not in LLM context)
      const display: ToolDisplay = {
        kind: 'images',
        items: r.images.map((im) => ({ url: im.imageUrl, title: im.title || undefined })),
      }
      return {
        output: lines.join('\n') || '(no images)',
        mutated: false,
        summary: t('aiSumImageSearch', { query, count: r.images.length }),
        display,
      }
    }

    case 'insert_web_image': {
      const idx = Number(call.input.slideIndex)
      if (!slides[idx])
        return fail(t('aiFailInsertImage'), `slideIndex out of range (0-${slides.length - 1})`)
      const url = String(call.input.url ?? '')
      if (!/^https?:\/\//.test(url)) return fail(t('aiFailInsertImage'), 'Invalid url')
      const r = await window.slidesApi.insertImageUrl({
        slideIndex: idx,
        url,
        xPx: Number(call.input.x),
        yPx: Number(call.input.y),
        wPx: Number(call.input.w),
        hPx: Number(call.input.h),
        fitWidthPx: access.fitWidthPx,
      })
      signal?.throwIfAborted()
      if (!r)
        return fail(
          t('aiFailInsertImage'),
          'Download or insertion failed (the image may be inaccessible)',
        )
      access.applySlide(idx, r.slide)
      return {
        output: `Inserted the image on page ${idx + 1}, element id=${r.sourceId}.`,
        mutated: true,
        summary: t('aiSumInsertImage', { n: idx + 1 }),
      }
    }

    case 'crop_image':
    case 'set_picture_opacity':
    case 'replace_image': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      const failKey =
        call.name === 'crop_image'
          ? ('aiFailCropImage' as const)
          : call.name === 'set_picture_opacity'
            ? ('aiFailPictureOpacity' as const)
            : ('aiFailReplaceImage' as const)
      const slide = slides[idx]
      if (!slide) return fail(t(failKey), `slideIndex out of range (0-${slides.length - 1})`)
      const target = resolveEditTarget(slide, sourceId)
      const terr = targetError(target, sourceId, idx + 1)
      if (terr || !target || 'nested' in target) return fail(t(failKey), terr!)
      if (target.node.type !== 'picture')
        return fail(t(failKey), `Element ${sourceId} is not a picture (type: ${target.node.type})`)
      if (target.groupId)
        return fail(
          t(failKey),
          `Element ${sourceId} is inside a group; this tool only supports top-level pictures — ungroup_element first`,
        )

      if (call.name === 'crop_image') {
        const frac = (v: unknown) => Math.min(1, Math.max(0, Number(v) || 0))
        const cl = frac(call.input.l)
        const ct = frac(call.input.t)
        const cr = frac(call.input.r)
        const cb = frac(call.input.b)
        if (cl + cr >= 0.99 || ct + cb >= 0.99)
          return fail(t(failKey), 'Crop removes the whole image (l+r and t+b must be < 1)')
        const srcRect = cl || ct || cr || cb ? { l: cl, t: ct, r: cr, b: cb } : null
        const updated = await window.slidesApi.editPictureSrcRect({
          slideIndex: idx,
          sourceId,
          srcRect,
        })
        signal?.throwIfAborted()
        if (!updated) return fail(t(failKey), 'Crop failed')
        access.applySlide(idx, updated)
        return {
          output: srcRect
            ? `Cropped picture ${sourceId} on page ${idx + 1} (l=${cl} t=${ct} r=${cr} b=${cb}).`
            : `Removed the crop of picture ${sourceId} on page ${idx + 1}.`,
          mutated: true,
          summary: t('aiSumCropImage', { n: idx + 1 }),
        }
      }

      if (call.name === 'set_picture_opacity') {
        const opacity = Number(call.input.opacity)
        if (!Number.isFinite(opacity) || opacity < 0 || opacity > 1)
          return fail(t(failKey), 'opacity must be between 0 and 1')
        const updated = await window.slidesApi.editPictureOpacity({
          slideIndex: idx,
          sourceId,
          opacity,
        })
        signal?.throwIfAborted()
        if (!updated) return fail(t(failKey), 'Opacity change failed')
        access.applySlide(idx, updated)
        return {
          output: `Set the opacity of picture ${sourceId} on page ${idx + 1} to ${opacity}.`,
          mutated: true,
          summary: t('aiSumPictureOpacity', { n: idx + 1 }),
        }
      }

      const url = String(call.input.url ?? '')
      if (!/^https?:\/\//.test(url)) return fail(t(failKey), 'Invalid url')
      const updated = await window.slidesApi.replacePictureUrl({
        slideIndex: idx,
        sourceId,
        url,
        ...(call.input.keepCrop ? { keepSrcRect: true } : {}),
      })
      signal?.throwIfAborted()
      if (!updated)
        return fail(
          t(failKey),
          'Replacement failed (the image may be inaccessible, or the element is not a replaceable picture)',
        )
      access.applySlide(idx, updated)
      return {
        output: `Replaced the image of picture ${sourceId} on page ${idx + 1} in place (frame/z-order/effects kept).`,
        mutated: true,
        summary: t('aiSumReplaceImage', { n: idx + 1 }),
      }
    }

    case 'ask_clarification': {
      if (!access.askClarification)
        return fail(
          t('aiFailClarify'),
          'The current environment does not support questionnaire cards',
        )
      const raw = Array.isArray(call.input.questions) ? call.input.questions : []
      const questions: ClarifyQuestion[] = raw
        .map((q: Record<string, unknown>, i: number) => ({
          id: String(q.id ?? `q${i + 1}`),
          label: String(q.label ?? ''),
          description: q.description ? String(q.description) : undefined,
          options: Array.isArray(q.options)
            ? q.options.map((o: unknown) => String(o)).slice(0, 5)
            : [],
          multi: !!q.multi,
        }))
        .filter((q) => q.label && q.options.length > 0)
      if (questions.length === 0)
        return fail(
          t('aiFailClarify'),
          'questions must be non-empty and every question needs options',
        )
      const r = await access.askClarification(questions)
      signal?.throwIfAborted()
      if (r.cancelled) {
        return {
          output:
            'The user skipped the questionnaire. Decide the Core Hook and style yourself based on professional judgment and generate directly.',
          mutated: false,
          summary: t('aiSumClarifySkipped'),
        }
      }
      return {
        output: `User questionnaire answers:\n${r.answers}\nDecide the Core Hook and style accordingly, then build the slides with the available local tools.`,
        mutated: false,
        summary: t('aiSumClarifyDone'),
      }
    }

    case 'plan_deck': {
      const coreHook = String(call.input.core_hook ?? '').trim()
      const style = String(call.input.style ?? '').trim()
      const pages = Array.isArray(call.input.pages) ? call.input.pages : []
      if (!coreHook || !style || pages.length === 0) {
        return fail(t('aiFailPlan'), 'plan_deck requires core_hook + style + non-empty pages')
      }
      // Planning summary echoed back to the user
      const lines = pages.map((p: Record<string, unknown>, i: number) => {
        const q =
          Array.isArray(p.image_queries) && p.image_queries.length
            ? ` [images: ${p.image_queries.length}]`
            : ''
        return `Page ${i + 1} [${String(p.layout ?? '')}] ${String(p.title ?? '')} — ${String(p.brief ?? '').slice(0, 40)}${q}`
      })
      const summary = t('aiSumPlan', { count: pages.length, hook: coreHook })
      return {
        output: `Plan confirmed:\nCore Hook: ${coreHook}\nStyle: ${style}\n${lines.join('\n')}\nNow follow this plan and build all ${pages.length} pages with the available local slide and element tools. Use the latest deck outline and each tool result to track actual completion.`,
        mutated: false,
        summary,
      }
    }

    case 'delete_slide': {
      const idx = Number(call.input.slideIndex)
      if (!slides[idx])
        return fail(t('aiFailDeleteSlide'), `slideIndex out of range (0-${slides.length - 1})`)
      if (slides.length <= 1)
        return fail(t('aiFailDeleteSlide'), 'Only one page remains; cannot delete')
      const r = await window.slidesApi.deleteSlide(idx)
      signal?.throwIfAborted()
      if (!r) return fail(t('aiFailDeleteSlide'), 'Deletion failed')
      access.applyDeck(r, Math.max(0, Math.min(idx, r.length - 1)))
      return {
        output: `Deleted page ${idx + 1}; the deck now has ${r.length} pages. Note that slideIndex of pages after it shifted down by 1.`,
        mutated: true,
        summary: t('aiSumDeleteSlide', { n: idx + 1 }),
      }
    }

    case 'add_slide': {
      const src = Number(call.input.sourceIndex)
      if (!slides[src])
        return fail(t('aiFailNewSlide'), `sourceIndex out of range (0-${slides.length - 1})`)
      const r = await window.slidesApi.addSlide({
        sourceIndex: src,
        clearText: call.input.clearText !== false,
        fitWidthPx: access.fitWidthPx,
      })
      signal?.throwIfAborted()
      if (!r) return fail(t('aiFailNewSlide'), 'Creation failed')
      access.applyDeck(r.slides, r.index)
      return {
        output: `Created page ${r.index + 1} (${r.slides.length} pages total). ✅ Use slideIndex=${r.index} when filling content into this new page (not 1, unless it happens to be 1). To add another page after it, use sourceIndex=${r.slides.length - 1} (current last page) so it appends at the end.`,
        mutated: true,
        summary: t('aiSumNewSlide', { n: r.index + 1 }),
      }
    }

    case 'add_text_box': {
      const idx = Number(call.input.slideIndex)
      if (!slides[idx])
        return fail(t('aiFailNewElement'), `slideIndex out of range (0-${slides.length - 1})`)
      const paragraphs = toEditParagraphs(call.input.paragraphs)
      if (!paragraphs) return fail(t('aiFailNewTextbox'), 'paragraphs must be a non-empty array')
      if (!access.executePresentationOperation)
        return fail(t('aiFailNewTextbox'), 'Canonical presentation transactions are unavailable')
      const clientId = 'created-text-box'
      const canonical = paragraphs.map((paragraph) => ({
        runs: paragraph.runs.map((run) => ({
          ...run,
          ...(run.color ? { color: run.color.toUpperCase() } : {}),
        })),
        ...(paragraph.align === 'left' ||
        paragraph.align === 'center' ||
        paragraph.align === 'right'
          ? { align: paragraph.align }
          : {}),
      }))
      const text = canonical.map((p) => p.runs.map((run) => run.text).join('')).join('\n')
      const execution = await access.executePresentationOperation(
        {
          transactionId: await geometryToolTransactionId(call),
          slideIndex: idx,
          operations: [
            {
              kind: 'add_text_box',
              clientId,
              text,
              geometry: {
                x: geometryPxToPoints(
                  Number(call.input.x),
                  slides[idx]!.scale || access.fitWidthPx / slides[idx]!.widthPx,
                ),
                y: geometryPxToPoints(
                  Number(call.input.y),
                  slides[idx]!.scale || access.fitWidthPx / slides[idx]!.widthPx,
                ),
                width: geometryPxToPoints(
                  Number(call.input.w),
                  slides[idx]!.scale || access.fitWidthPx / slides[idx]!.widthPx,
                ),
                height: geometryPxToPoints(
                  Number(call.input.h),
                  slides[idx]!.scale || access.fitWidthPx / slides[idx]!.widthPx,
                ),
              },
            },
            { kind: 'set_text', createdByClientId: clientId, paragraphs: canonical },
          ],
        },
        signal,
      )
      const outcome = textFamilyReceiptOutcome(execution.receipt)
      if (!outcome.ok)
        return fail(t('aiFailNewTextbox'), outcome.detail ?? 'Insertion was not applied')
      const createdId =
        execution.receipt.status === 'applied'
          ? (execution.receipt.createdTargets?.find((target) => target.clientId === 'op-1')
              ?.elementId ?? execution.receipt.createdIds?.[0])
          : undefined
      return {
        output: outcome.mutated
          ? `Created a new text box on page ${idx + 1}${createdId ? `, durable element id=${createdId}` : ''}.`
          : `The requested text box already exists from this tool invocation.`,
        mutated: outcome.mutated,
        ...(execution.authoritativeState === 'reload_required' ? { stopToolBatch: true } : {}),
        summary: t('aiSumNewTextbox', { n: idx + 1 }),
      }
    }

    case 'add_shape': {
      const idx = Number(call.input.slideIndex)
      if (!slides[idx])
        return fail(t('aiFailNewElement'), `slideIndex out of range (0-${slides.length - 1})`)
      const paragraphs = toEditParagraphs(call.input.paragraphs)
      const kind = String(call.input.kind)
      if (!/^[a-zA-Z][a-zA-Z0-9]*$/.test(kind))
        return fail(t('aiFailNewShape'), `Invalid shape name: ${kind}`)
      const r = await window.slidesApi.addElement({
        slideIndex: idx,
        kind,
        xPx: Number(call.input.x),
        yPx: Number(call.input.y),
        wPx: Number(call.input.w),
        hPx: Number(call.input.h),
        fitWidthPx: access.fitWidthPx,
        ...(paragraphs ? { paragraphs } : {}),
        ...(call.input.fillColor ? { fillColor: String(call.input.fillColor) } : {}),
      })
      signal?.throwIfAborted()
      if (!r) return fail(t('aiFailNewElement'), 'Insertion failed')
      access.applySlide(idx, r.slide)
      return {
        output: `Created a new shape on page ${idx + 1}, element id=${r.sourceId}.`,
        mutated: true,
        summary: t('aiSumNewShape', { n: idx + 1 }),
      }
    }

    case 'add_chart': {
      const idx = Number(call.input.slideIndex)
      const slide = slides[idx]
      if (!slide) return fail(t('aiFailChart'), `slideIndex out of range (0-${slides.length - 1})`)
      const categories = Array.isArray(call.input.categories)
        ? call.input.categories.map(String)
        : []
      const seriesRaw = Array.isArray(call.input.series) ? call.input.series : []
      const series = seriesRaw
        .map((s) => ({
          name: String((s as { name?: unknown }).name ?? ''),
          values: Array.isArray((s as { values?: unknown }).values)
            ? ((s as { values: unknown[] }).values.map(Number) as number[])
            : [],
        }))
        .filter((s) => s.values.length > 0)
      if (categories.length === 0 || series.length === 0) {
        return fail(t('aiFailChart'), 'Neither categories nor series may be empty')
      }
      const gateErr = dataSourceGateError(call, state)
      if (gateErr) return fail(t('aiFailChart'), gateErr)
      const defW = Math.round(slide.widthPx * 0.62)
      const defH = Math.round(slide.heightPx * 0.62)
      const w = Number(call.input.w) || defW
      const h = Number(call.input.h) || defH
      const r = await window.slidesApi.addChart({
        slideIndex: idx,
        kind: String(call.input.kind) as
          'bar' | 'barStacked' | 'line' | 'area' | 'pie' | 'doughnut',
        ...(call.input.title ? { title: String(call.input.title) } : {}),
        categories,
        series,
        xPx:
          Number.isFinite(Number(call.input.x)) && call.input.x != null
            ? Number(call.input.x)
            : Math.round((slide.widthPx - w) / 2),
        yPx:
          Number.isFinite(Number(call.input.y)) && call.input.y != null
            ? Number(call.input.y)
            : Math.round((slide.heightPx - h) / 2),
        wPx: w,
        hPx: h,
        fitWidthPx: access.fitWidthPx,
      })
      signal?.throwIfAborted()
      if (!r) return fail(t('aiFailChart'), 'Insertion failed (check kind and data)')
      access.applySlide(idx, r.slide)
      const sampleNote = call.input.dataSource === 'sample' ? SAMPLE_DATA_NOTE : ''
      return {
        output: `Inserted a ${String(call.input.kind)} chart on page ${idx + 1}, element id=${r.sourceId}.${sampleNote}`,
        mutated: true,
        summary: t('aiSumChart', { n: idx + 1 }),
      }
    }

    case 'add_smartart': {
      const idx = Number(call.input.slideIndex)
      const slide = slides[idx]
      if (!slide)
        return fail(t('aiFailSmartart'), `slideIndex out of range (0-${slides.length - 1})`)
      const items = Array.isArray(call.input.items)
        ? call.input.items.map(String).filter(Boolean)
        : []
      if (items.length < 2) return fail(t('aiFailSmartart'), 'items requires at least 2 entries')
      const defW = Math.round(slide.widthPx * 0.7)
      const defH = Math.round(slide.heightPx * 0.5)
      const w = Number(call.input.w) || defW
      const h = Number(call.input.h) || defH
      const r = await window.slidesApi.addSmartArt({
        slideIndex: idx,
        layout: String(call.input.layout) as AddSmartArtOp['layout'],
        items,
        xPx:
          Number.isFinite(Number(call.input.x)) && call.input.x != null
            ? Number(call.input.x)
            : Math.round((slide.widthPx - w) / 2),
        yPx:
          Number.isFinite(Number(call.input.y)) && call.input.y != null
            ? Number(call.input.y)
            : Math.round((slide.heightPx - h) / 2),
        wPx: w,
        hPx: h,
        fitWidthPx: access.fitWidthPx,
      })
      signal?.throwIfAborted()
      if (!r) return fail(t('aiFailSmartart'), 'Insertion failed (check layout)')
      access.applySlide(idx, r.slide)
      return {
        output: `Inserted a ${String(call.input.layout)} diagram (${items.length} nodes) on page ${idx + 1}, element id=${r.sourceId}.`,
        mutated: true,
        summary: t('aiSumSmartart', { n: idx + 1 }),
      }
    }

    case 'add_table': {
      const idx = Number(call.input.slideIndex)
      const slide = slides[idx]
      if (!slide) return fail(t('aiFailTable'), `slideIndex out of range (0-${slides.length - 1})`)
      const rows = Number(call.input.rows)
      const cols = Number(call.input.cols)
      if (
        !Number.isInteger(rows) ||
        !Number.isInteger(cols) ||
        rows < 1 ||
        cols < 1 ||
        rows > 30 ||
        cols > 12
      ) {
        return fail(t('aiFailTable'), 'Invalid rows (1-30) / cols (1-12)')
      }
      const defW = Math.round(slide.widthPx * 0.7)
      const defH = Math.round(Math.min(slide.heightPx * 0.6, rows * 40 + 20))
      const w = Number(call.input.w) || defW
      const h = Number(call.input.h) || defH
      const r = await window.slidesApi.addTable({
        slideIndex: idx,
        rows,
        cols,
        xPx:
          Number.isFinite(Number(call.input.x)) && call.input.x != null
            ? Number(call.input.x)
            : Math.round((slide.widthPx - w) / 2),
        yPx:
          Number.isFinite(Number(call.input.y)) && call.input.y != null
            ? Number(call.input.y)
            : Math.round((slide.heightPx - h) / 2),
        wPx: w,
        hPx: h,
        fitWidthPx: access.fitWidthPx,
      })
      signal?.throwIfAborted()
      if (!r) return fail(t('aiFailTable'), 'Insertion failed')
      let updated = r.slide
      // Fill cells one by one (cells optional; out-of-range parts ignored)
      const cells = Array.isArray(call.input.cells) ? (call.input.cells as unknown[][]) : []
      let filled = 0
      for (let ri = 0; ri < Math.min(cells.length, rows); ri++) {
        const rowCells = Array.isArray(cells[ri]) ? cells[ri]! : []
        for (let ci = 0; ci < Math.min(rowCells.length, cols); ci++) {
          const text = String(rowCells[ci] ?? '')
          if (!text) continue
          const u = await window.slidesApi.editTableCell({
            slideIndex: idx,
            sourceId: r.sourceId,
            row: ri,
            col: ci,
            paragraphs: [{ runs: [{ text }] }],
          })
          signal?.throwIfAborted()
          if (u) {
            updated = u
            filled++
          }
        }
      }
      access.applySlide(idx, updated)
      return {
        output: `Inserted a ${rows}×${cols} table on page ${idx + 1}, element id=${r.sourceId}${filled ? `, filled ${filled} cell(s) with text` : ''}.`,
        mutated: true,
        summary: t('aiSumTable', { n: idx + 1 }),
      }
    }

    case 'edit_table_cell': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailEditTable'), `slideIndex out of range (0-${slides.length - 1})`)
      const paragraphs = toEditParagraphs(call.input.paragraphs)
      if (!paragraphs) return fail(t('aiFailEditTable'), 'paragraphs must be a non-empty array')
      const row = Number(call.input.row)
      const col = Number(call.input.col)
      const updated = await window.slidesApi.editTableCell({
        slideIndex: idx,
        sourceId,
        row,
        col,
        paragraphs,
      })
      signal?.throwIfAborted()
      if (!updated)
        return fail(
          t('aiFailEditTable'),
          `Table ${sourceId} not found or cell (${row},${col}) out of range`,
        )
      access.applySlide(idx, updated)
      return {
        output: `Replaced the text of cell (${row},${col}) in table ${sourceId} on page ${idx + 1}.`,
        mutated: true,
        summary: t('aiSumTableCell', { n: idx + 1 }),
      }
    }

    case 'edit_table_structure': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailTableStructure'), `slideIndex out of range (0-${slides.length - 1})`)
      const kind = String(call.input.kind) as
        'insert-row' | 'delete-row' | 'insert-col' | 'delete-col'
      if (!['insert-row', 'delete-row', 'insert-col', 'delete-col'].includes(kind)) {
        return fail(t('aiFailTableStructure'), 'Invalid kind')
      }
      const r = await window.slidesApi.tableStructure({
        slideIndex: idx,
        sourceId,
        kind,
        index: Number(call.input.index),
        ...(call.input.before ? { before: true } : {}),
      })
      signal?.throwIfAborted()
      if (!r)
        return fail(
          t('aiFailTableStructure'),
          `Operation failed (table ${sourceId} does not exist, index out of range, or the last row/column cannot be deleted)`,
        )
      access.applySlide(idx, r.slide)
      return {
        output: `Applied ${kind} (index=${Number(call.input.index)}) to table ${sourceId} on page ${idx + 1}. The table id may have been updated to ${r.sourceId}.`,
        mutated: true,
        summary: t('aiSumTableStructure', {
          n: idx + 1,
          op: t(
            kind.startsWith('insert')
              ? kind.endsWith('row')
                ? 'aiOpInsertRow'
                : 'aiOpInsertCol'
              : kind.endsWith('row')
                ? 'aiOpDeleteRow'
                : 'aiOpDeleteCol',
          ),
        }),
      }
    }

    case 'edit_table_style': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailTableStyle'), `slideIndex out of range (0-${slides.length - 1})`)
      const op: import('../../shared/ipc').EditTableStyleOp = { slideIndex: idx, sourceId }
      if (call.input.styleName != null) op.styleName = String(call.input.styleName)
      if (call.input.firstRow != null) op.firstRow = Boolean(call.input.firstRow)
      if (call.input.bandRow != null) op.bandRow = Boolean(call.input.bandRow)
      if (call.input.shadingColor != null) op.shadingColor = String(call.input.shadingColor)
      if (call.input.borderColor != null) op.borderColor = String(call.input.borderColor)
      if (call.input.borderWidthPt != null) op.borderWidthPt = Number(call.input.borderWidthPt)
      if (call.input.borderPreset != null)
        op.borderPreset = String(call.input.borderPreset) as 'all' | 'none'
      const updated = await window.slidesApi.editTableStyle(op)
      signal?.throwIfAborted()
      if (!updated)
        return fail(
          t('aiFailTableStyle'),
          `Operation failed (table ${sourceId} does not exist or is not of type table)`,
        )
      access.applySlide(idx, updated.slide)
      return {
        output: `Updated the style of table ${sourceId} on page ${idx + 1}.`,
        mutated: true,
        summary: t('aiSumTableStyle', { n: idx + 1 }),
      }
    }

    case 'edit_chart': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailChartEdit'), `slideIndex out of range (0-${slides.length - 1})`)
      const op: import('../../shared/ipc').EditChartOp = { slideIndex: idx, sourceId }
      if (call.input.kind != null)
        op.kind = String(call.input.kind) as import('../../shared/ipc').EditChartOp['kind']
      if (Array.isArray(call.input.categories))
        op.categories = (call.input.categories as unknown[]).map(String)
      if (Array.isArray(call.input.series)) {
        const gateErr = dataSourceGateError(call, state)
        if (gateErr) return fail(t('aiFailChartEdit'), gateErr)
        op.series = (call.input.series as Array<{ name: unknown; values: unknown[] }>).map((s) => ({
          name: String(s.name ?? ''),
          values: (Array.isArray(s.values) ? s.values : []).map(Number),
        }))
      }
      if (call.input.colorScheme != null) op.colorScheme = String(call.input.colorScheme)
      if (call.input.title != null) op.title = String(call.input.title)
      if (call.input.legendPos != null)
        op.legendPos = String(
          call.input.legendPos,
        ) as import('../../shared/ipc').EditChartOp['legendPos']
      if (typeof call.input.dataLabels === 'boolean') op.dataLabels = call.input.dataLabels
      if (typeof call.input.gridlines === 'boolean') op.gridlines = call.input.gridlines
      if (call.input.switchRowCol === true) op.switchRowCol = true
      const updated = await window.slidesApi.editChart(op)
      signal?.throwIfAborted()
      if (!updated)
        return fail(
          t('aiFailChartEdit'),
          `Operation failed (element ${sourceId} does not exist or is not a chart)`,
        )
      access.applySlide(idx, updated.slide)
      const sampleNote = op.series && call.input.dataSource === 'sample' ? SAMPLE_DATA_NOTE : ''
      return {
        output: `Updated chart ${sourceId} on page ${idx + 1}.${sampleNote}`,
        mutated: true,
        summary: t('aiSumChartEdit', { n: idx + 1 }),
      }
    }

    case 'set_slide_background': {
      const idx = Number(call.input.slideIndex)
      const color = String(call.input.color ?? '')
      if (idx !== -1 && !slides[idx])
        return fail(t('aiFailBackground'), `slideIndex out of range (0-${slides.length - 1} or -1)`)
      if (!/^#?[0-9a-fA-F]{6}$/.test(color))
        return fail(t('aiFailBackground'), 'color must be #RRGGBB')
      if (!access.executePresentationOperation)
        return fail(t('aiFailBackground'), 'Canonical presentation transactions are unavailable')
      const normalizedColor = `#${color.replace(/^#/, '').toUpperCase()}`
      const backgrounds = (idx === -1 ? slides.map((_slide, slideIndex) => slideIndex) : [idx]).map(
        (slideIndex) => ({ slideIndex, color: normalizedColor }),
      )
      const execution = await access.executePresentationOperation(
        {
          transactionId: await backgroundToolTransactionId(call),
          backgrounds,
        },
        signal,
      )
      const outcome = textFamilyReceiptOutcome(execution.receipt)
      if (!outcome.ok)
        return fail(t('aiFailBackground'), outcome.detail ?? 'Setting was not applied')
      return {
        output:
          idx === -1
            ? `${outcome.mutated ? 'Set' : 'Kept'} the background of all ${slides.length} pages at ${normalizedColor}.`
            : `${outcome.mutated ? 'Set' : 'Kept'} the background of page ${idx + 1} at ${normalizedColor}.`,
        mutated: outcome.mutated,
        ...(execution.authoritativeState === 'reload_required' ? { stopToolBatch: true } : {}),
        summary: idx === -1 ? t('aiSumBackgroundAll') : t('aiSumBackground', { n: idx + 1 }),
      }
    }

    case 'set_speaker_notes': {
      const idx = call.input.slideIndex
      const text = call.input.text
      if (typeof idx !== 'number' || !Number.isInteger(idx) || !slides[idx])
        return fail(t('aiFailEditText'), `slideIndex out of range (0-${slides.length - 1})`)
      if (typeof text !== 'string')
        return fail(t('aiFailEditText'), 'speaker notes text must be a string')
      if (text.length > 12_000)
        return fail(t('aiFailEditText'), 'speaker notes exceed the 12000 character limit')
      if (!access.executePresentationOperation)
        return fail(t('aiFailEditText'), 'speaker notes are unavailable in this host')
      const draftPreparation = (await access.prepareSpeakerNotesWrite?.(idx)) ?? {
        ready: true,
        expectedDraftVersion: 0,
      }
      if (!draftPreparation.ready)
        return fail(
          t('aiFailEditText'),
          'Speaker notes changed or could not be saved before the agent edit. No change was applied.',
        )
      signal?.throwIfAborted()
      const execution = await access.executePresentationOperation(
        {
          transactionId: await textToolTransactionId(call),
          slideIndex: idx,
          operation: { kind: 'set_speaker_notes', notes: text },
        },
        signal,
      )
      const outcome = textFamilyReceiptOutcome(execution.receipt)
      if (execution.authoritativeState === 'reload_required') {
        return {
          output:
            'The speaker notes were applied, but the editor could not refresh authoritative state. Reloading is required before more edits.',
          mutated: true,
          stopToolBatch: true,
          summary: t('aiSumEditText', { n: idx + 1 }),
        }
      }
      if (!outcome.ok) {
        return {
          output: outcome.mutated
            ? 'The speaker notes write could not be verified. Inspect the slide notes before retrying.'
            : `The speaker notes were not changed (${outcome.detail ?? 'write_not_applied'}).`,
          isError: true,
          mutated: outcome.mutated,
          summary: t('aiFailEditText'),
        }
      }
      if (outcome.mutated)
        access.applySpeakerNotes?.(idx, text, draftPreparation.expectedDraftVersion)
      return {
        output: outcome.mutated
          ? text
            ? `Wrote speaker notes for page ${idx + 1} (${text.length} characters).`
            : `Cleared speaker notes for page ${idx + 1}.`
          : `Speaker notes on page ${idx + 1} already matched.`,
        mutated: outcome.mutated,
        summary: t('aiSumEditText', { n: idx + 1 }),
      }
    }

    case 'delete_element': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailDeleteElement'), `slideIndex out of range (0-${slides.length - 1})`)
      // Deletion is top-level only: for group members guide to ungroup instead of a misleading "not found"
      const target = resolveEditTarget(slides[idx]!, sourceId)
      if (target && ('nested' in target || target.groupId)) {
        const gid = 'nested' in target ? undefined : target.groupId
        return fail(
          t('aiFailDeleteElement'),
          `Element ${sourceId} is inside a group${gid ? ` (${gid})` : ''}; call ungroup_element on the group first and then delete it, or delete the whole group`,
        )
      }
      if (!target)
        return fail(t('aiFailDeleteElement'), `Element ${sourceId} not found on page ${idx + 1}`)
      if (!access.executePresentationOperation)
        return fail(t('aiFailDeleteElement'), 'Canonical presentation transactions are unavailable')
      const execution = await access.executePresentationOperation(
        {
          transactionId: await geometryToolTransactionId(call),
          slideIndex: idx,
          operations: [{ sourceId, kind: 'delete_element' }],
        },
        signal,
      )
      const outcome = textFamilyReceiptOutcome(execution.receipt)
      if (!outcome.ok)
        return fail(t('aiFailDeleteElement'), outcome.detail ?? 'Deletion was not applied')
      return {
        output: outcome.mutated
          ? `Deleted element ${sourceId} from page ${idx + 1}.`
          : `Element ${sourceId} was already deleted by this tool invocation.`,
        mutated: outcome.mutated,
        ...(execution.authoritativeState === 'reload_required' ? { stopToolBatch: true } : {}),
        summary: t('aiSumDeleteElement', { n: idx + 1 }),
      }
    }

    case 'ungroup_element': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      const slide = slides[idx]
      if (!slide)
        return fail(t('aiFailUngroup'), `slideIndex out of range (0-${slides.length - 1})`)
      const node = slide.nodes.find((n) => n.sourceId === sourceId)
      if (!node) {
        return fail(
          t('aiFailUngroup'),
          findNodeById(slide.nodes, sourceId)
            ? `${sourceId} is inside another group; ungroup the outer group first`
            : `Element ${sourceId} not found on page ${idx + 1}`,
        )
      }
      if (node.type !== 'group')
        return fail(t('aiFailUngroup'), `${sourceId} is not a group (type: ${node.type})`)
      if (node.decoration)
        return fail(t('aiFailUngroup'), `${sourceId} is a layout decoration, read-only`)
      const updated = await window.slidesApi.ungroupElement({ slideIndex: idx, sourceId })
      signal?.throwIfAborted()
      if (!updated) return fail(t('aiFailUngroup'), 'Ungroup failed')
      access.applySlide(idx, updated)
      // Ungrouping rewrites the page and re-ids every element; echo the fresh list so no extra read_slide is needed
      const fresh = collectNodeInfos(updated.nodes)
        .map((n) => `${n.id} | ${n.type}${n.text ? ` | ${preview(n.text)}` : ''}`)
        .join('\n')
      return {
        output: `Ungrouped ${sourceId} on page ${idx + 1} into ${node.children.length} top-level elements. All element ids on this page changed; current elements:\n${fresh}`,
        mutated: true,
        summary: t('aiSumUngroup', { n: idx + 1 }),
      }
    }

    case 'save_style_template': {
      const name = String(call.input.name ?? '').trim()
      if (!name) return fail(t('aiFailSaveTemplate'), 'name must not be empty')
      if (!access.saveStyleTemplate)
        return fail(
          t('aiFailSaveTemplate'),
          'The current environment does not support template saving',
        )
      const styleSkillToSave = state?.lastStyleSkill ?? ''
      const topicToSave = state?.lastTopic ?? ''
      if (!styleSkillToSave)
        return fail(t('aiFailSaveTemplate'), 'The current deck has no reusable Style Skill to save')
      const r = await access.saveStyleTemplate(name, {
        topic: topicToSave,
        styleSkill: styleSkillToSave,
        createdAt: new Date().toISOString(),
      })
      signal?.throwIfAborted()
      if (!r.ok) return fail(t('aiFailSaveTemplate'), r.error ?? 'Save failed')
      return {
        output: `Saved the style "${name}" as a template; next time pass style_template:"${name}" to reuse it directly.`,
        mutated: false,
        summary: t('aiSumSaveTemplate', { name }),
      }
    }

    case 'list_style_templates': {
      if (!access.listStyleTemplates)
        return fail(
          t('aiFailListTemplates'),
          'The current environment does not support template listing',
        )
      const templates = await access.listStyleTemplates()
      signal?.throwIfAborted()
      if (templates.length === 0) {
        return {
          output:
            'No saved style templates yet. After generating a deck, call save_style_template(name) to save the current style.',
          mutated: false,
          summary: t('aiSumTemplatesEmpty'),
        }
      }
      const lines = templates.map(
        (t) => `- ${t.name} (topic: ${t.topic || 'unknown'}, saved ${t.createdAt.slice(0, 10)})`,
      )
      return {
        output: `Saved style templates (${templates.length}):\n${lines.join('\n')}\n\nSelect a saved template to reuse its style locally.`,
        mutated: false,
        summary: t('aiSumListTemplates', { count: templates.length }),
      }
    }

    default:
      return fail(call.name, `Unknown tool: ${call.name}`)
  }
}
