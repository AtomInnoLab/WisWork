export const PRESENTATION_DESIGN_WORKFLOW_PROMPT = `
## Shared presentation design workflow
For a whole-deck creation or redesign, work as a presentation director, not a text filler:
1. Brief: establish audience, occasion, desired decision, page count, source constraints, and one narrative conclusion. Ask only for missing choices that materially change the result.
2. DESIGN.md contract: immediately after the brief or questionnaire, draft a user-editable design contract covering the confirmed choices, open questions, research plan, concrete color tokens, typography hierarchy, safe margins/grid, image treatment, density limits, layout families, and composition rules. Create and show this draft immediately, before fact or image research. Keep the first draft compact: record the brief and discovery log, but keep the page plan and asset lists empty until research supplies them. Treat it as binding for every slide. As facts and assets are collected, update the same draft with evidence, provenance, and asset status. Do not mark the final contract ready until required fact research, image search, and asset validation are complete.
3. Deck plan: give every slide one conclusion-led headline, a narrative role, one focal visual, supporting evidence, a layout family, asset needs/provenance, density, and slide-specific acceptance criteria. Do not use the same layout family on adjacent slides unless continuity requires it.
4. Asset plan: search real people, products, places, brands, and current facts; validate selected image sources before finalizing DESIGN.md as ready. Generate only abstract/custom illustration when generation exists. Use native editable charts for data. Never invent precise data or image URLs.
5. Prototype gate: first create or identify the cover, one representative content page, and one complex visual page. Screenshot and review those pages before continuing the remaining production batches. When fewer than three slides are requested, review every slide.
6. Production loop: complete the remaining work in batches of 2–3 slides. After each batch, inspect screenshots plus geometry, repair concrete defects, and re-screenshot changed slides before continuing.
7. Quality bar: each slide needs one clear conclusion and one focal visual; readable hierarchy, deliberate whitespace, aligned geometry, sufficient contrast, relevant imagery, and no accidental overflow, overlap, distortion, placeholder content, or repetitive card grids. Review design-system consistency and rhythm across adjacent slides, then run final whole-deck verification.
Keep all edits native, editable, reversible, and within the host's permission model. If a capability is unavailable, use the declared fallback and report the specific unresolved limitation.
`.trim()

export function buildPresentationDesignDocument(style: string): string {
  const body = style.trim()
  if (!body) throw new Error('empty_presentation_design')
  return `# DESIGN.md\n\n${body}`
}

const hasPresentationDesignBody = (value: string): boolean =>
  Boolean(value.replace(/^\s*#\s*DESIGN\.md\s*/i, '').trim())

/** Extract a design snapshot from either desktop prose or Office JSON tool output. */
export function extractPresentationDesignDocument(output: string): string | undefined {
  try {
    const value = JSON.parse(output) as { designMd?: unknown }
    if (typeof value.designMd === 'string' && hasPresentationDesignBody(value.designMd))
      return value.designMd.trim()
  } catch {
    // Desktop plan output is intentionally readable prose rather than JSON.
  }
  const start = output.indexOf('# DESIGN.md')
  if (start < 0) return undefined
  const rest = output.slice(start)
  const end = rest.search(/\n# [^\n]+\n/)
  const designDocument = (end < 0 ? rest : rest.slice(0, end)).trim()
  return hasPresentationDesignBody(designDocument) ? designDocument : undefined
}

export interface PresentationDesignPagePlan {
  title: string
  type?: string
  brief: string
  layout: string
  purpose: string
  visual: string
  evidence: string[]
  acceptance: string[]
  density: 'low' | 'medium' | 'high'
  image_queries: string[]
}

export interface PresentationDesignPlan {
  core_hook: string
  style: string
  pages: PresentationDesignPagePlan[]
  prototype_pages: number[]
}

export type PresentationDesignStatus = 'draft' | 'ready' | 'producing' | 'verified'
export type PresentationAssetStatus =
  'needed' | 'searching' | 'downloaded' | 'validated' | 'ready' | 'fallback_ready'

export type PresentationDesignInvalidationScope =
  { type: 'slide'; slideNumbers: number[] } | { type: 'global' } | { type: 'narrative' }

export interface PresentationDesignInvalidation {
  previousRevision: number
  revision: number
  reason: string
  scope: PresentationDesignInvalidationScope
}

export interface PresentationDesignAcceptanceRule {
  id: string
  criterion: string
}

export interface PresentationDesignContract {
  schemaVersion: 1
  revision: number
  status: PresentationDesignStatus
  prototypePages: number[]
  discovery?: {
    questionnaire: string[]
    openQuestions: string[]
    researchNotes: string[]
  }
  brief: {
    topic: string
    audience: string
    occasion: string
    desiredOutcome: string
    language: string
    pageCount: number
    aspectRatio: string
    sourceConstraints: string[]
  }
  narrative: {
    coreHook: string
    opening: string
    development: string
    tension: string
    resolution: string
    closingAction: string
  }
  visualSystem: {
    style: string
    colors: Record<string, string>
    typography: Record<string, string>
    safeMargin: string
    grid: string
    imageTreatment: string
    chartTreatment: string
    antiPatterns: string[]
  }
  slides: Array<{
    number: number
    title: string
    role: string
    claim: string
    content: string[]
    evidence: string[]
    visualRoute: string
    layoutFamily: string
    focalVisual: string
    density: 'low' | 'medium' | 'high'
    assetIds: string[]
    acceptance: PresentationDesignAcceptanceRule[]
  }>
  assets: Array<{
    id: string
    slideNumbers: number[]
    type: string
    role: string
    intent: string
    source: string
    crop: string
    placement: string
    status: PresentationAssetStatus
    localReference?: string
    fallback?: string
  }>
  deckAcceptance: PresentationDesignAcceptanceRule[]
}

const boundedString = (maxLength: number) => ({ type: 'string', minLength: 1, maxLength })
const stringArray = (maxItems: number, maxLength: number) => ({
  type: 'array',
  maxItems,
  items: boundedString(maxLength),
})
const acceptanceSchema = () => ({
  type: 'object',
  additionalProperties: true,
  required: ['id', 'criterion'],
  properties: { id: boundedString(80), criterion: boundedString(500) },
})

/** Shared model-facing schema. Parsers remain more tolerant when loading older snapshots. */
export const PRESENTATION_DESIGN_CONTRACT_SCHEMA: Record<string, unknown> = {
  type: 'object',
  additionalProperties: true,
  required: [
    'schemaVersion',
    'revision',
    'status',
    'prototypePages',
    'brief',
    'narrative',
    'visualSystem',
    'slides',
    'assets',
    'deckAcceptance',
  ],
  properties: {
    schemaVersion: { type: 'integer', const: 1 },
    revision: { type: 'integer', minimum: 1 },
    status: { type: 'string', enum: ['draft', 'ready', 'producing', 'verified'] },
    discovery: {
      type: 'object',
      additionalProperties: true,
      properties: {
        questionnaire: stringArray(30, 1_000),
        openQuestions: stringArray(30, 1_000),
        researchNotes: stringArray(100, 2_000),
      },
    },
    prototypePages: {
      type: 'array',
      minItems: 0,
      maxItems: 3,
      uniqueItems: true,
      items: { type: 'integer', minimum: 1, maximum: 12 },
    },
    brief: {
      type: 'object',
      additionalProperties: true,
      required: [
        'topic',
        'audience',
        'occasion',
        'desiredOutcome',
        'language',
        'pageCount',
        'aspectRatio',
        'sourceConstraints',
      ],
      properties: {
        topic: boundedString(500),
        audience: boundedString(500),
        occasion: boundedString(500),
        desiredOutcome: boundedString(1_000),
        language: boundedString(100),
        pageCount: { type: 'integer', minimum: 1, maximum: 12 },
        aspectRatio: boundedString(50),
        sourceConstraints: stringArray(30, 1_000),
      },
    },
    narrative: {
      type: 'object',
      additionalProperties: true,
      required: ['coreHook', 'opening', 'development', 'tension', 'resolution', 'closingAction'],
      properties: {
        coreHook: boundedString(1_000),
        opening: boundedString(1_000),
        development: boundedString(2_000),
        tension: boundedString(1_000),
        resolution: boundedString(1_000),
        closingAction: boundedString(1_000),
      },
    },
    visualSystem: {
      type: 'object',
      additionalProperties: true,
      required: [
        'style',
        'colors',
        'typography',
        'safeMargin',
        'grid',
        'imageTreatment',
        'chartTreatment',
        'antiPatterns',
      ],
      properties: {
        style: boundedString(6_000),
        colors: {
          type: 'object',
          minProperties: 1,
          maxProperties: 30,
          additionalProperties: boundedString(100),
        },
        typography: {
          type: 'object',
          minProperties: 1,
          maxProperties: 30,
          additionalProperties: boundedString(100),
        },
        safeMargin: boundedString(100),
        grid: boundedString(100),
        imageTreatment: boundedString(1_000),
        chartTreatment: boundedString(1_000),
        antiPatterns: stringArray(30, 500),
      },
    },
    slides: {
      type: 'array',
      minItems: 0,
      maxItems: 12,
      items: {
        type: 'object',
        additionalProperties: true,
        required: [
          'number',
          'title',
          'role',
          'claim',
          'content',
          'evidence',
          'visualRoute',
          'layoutFamily',
          'focalVisual',
          'density',
          'assetIds',
          'acceptance',
        ],
        properties: {
          number: { type: 'integer', minimum: 1, maximum: 12 },
          title: boundedString(300),
          role: boundedString(500),
          claim: boundedString(2_000),
          content: stringArray(20, 2_000),
          evidence: stringArray(20, 1_000),
          visualRoute: boundedString(1_000),
          layoutFamily: boundedString(100),
          focalVisual: boundedString(1_000),
          density: { type: 'string', enum: ['low', 'medium', 'high'] },
          assetIds: stringArray(20, 100),
          acceptance: { type: 'array', minItems: 1, maxItems: 20, items: acceptanceSchema() },
        },
      },
    },
    assets: {
      type: 'array',
      maxItems: 120,
      items: {
        type: 'object',
        additionalProperties: true,
        required: [
          'id',
          'slideNumbers',
          'type',
          'role',
          'intent',
          'source',
          'crop',
          'placement',
          'status',
        ],
        properties: {
          id: boundedString(100),
          slideNumbers: {
            type: 'array',
            maxItems: 12,
            uniqueItems: true,
            items: { type: 'integer', minimum: 1, maximum: 12 },
          },
          type: boundedString(100),
          role: boundedString(100),
          intent: boundedString(1_000),
          source: boundedString(2_000),
          crop: boundedString(200),
          placement: boundedString(500),
          status: {
            type: 'string',
            enum: ['needed', 'searching', 'downloaded', 'validated', 'ready', 'fallback_ready'],
          },
          localReference: boundedString(2_000),
          fallback: boundedString(2_000),
        },
      },
    },
    deckAcceptance: { type: 'array', minItems: 0, maxItems: 30, items: acceptanceSchema() },
  },
}

const text = (value: unknown, max: number): string => {
  if (typeof value !== 'string' || !value.trim() || value.length > max)
    throw new Error('invalid_presentation_plan')
  return value.trim()
}

const texts = (
  value: unknown,
  maxItems: number,
  maxLength: number,
  required: boolean,
): string[] => {
  if (!Array.isArray(value) || (required && value.length === 0) || value.length > maxItems)
    throw new Error('invalid_presentation_plan')
  return value.map((item) => text(item, maxLength))
}

export function parsePresentationDesignPlan(value: unknown): PresentationDesignPlan {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('invalid_presentation_plan')
  const plan = value as Record<string, unknown>
  if (!Array.isArray(plan.pages) || plan.pages.length === 0 || plan.pages.length > 12)
    throw new Error('invalid_presentation_plan')
  const pages = plan.pages.map((raw) => {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw))
      throw new Error('invalid_presentation_plan')
    const page = raw as Record<string, unknown>
    if (!['low', 'medium', 'high'].includes(String(page.density)))
      throw new Error('invalid_presentation_plan')
    return {
      title: text(page.title, 300),
      ...(page.type === undefined ? {} : { type: text(page.type, 50) }),
      brief: text(page.brief, 2_000),
      layout: text(page.layout, 100),
      purpose: text(page.purpose, 500),
      visual: text(page.visual, 1_000),
      evidence: texts(page.evidence ?? [], 8, 500, false),
      acceptance: texts(page.acceptance, 8, 300, true),
      density: page.density as PresentationDesignPagePlan['density'],
      image_queries: texts(page.image_queries ?? [], 4, 200, false),
    }
  })
  const prototypes = plan.prototype_pages
  const expected = Math.min(3, pages.length)
  if (
    !Array.isArray(prototypes) ||
    prototypes.length !== expected ||
    new Set(prototypes).size !== prototypes.length ||
    prototypes.some(
      (index) =>
        !Number.isSafeInteger(index) || (index as number) < 0 || (index as number) >= pages.length,
    )
  )
    throw new Error('invalid_presentation_plan')
  return {
    core_hook: text(plan.core_hook, 500),
    style: text(plan.style, 6_000),
    pages,
    prototype_pages: prototypes as number[],
  }
}

const optionalText = (value: unknown, max: number): string =>
  typeof value === 'string' && value.trim() && value.length <= max ? value.trim() : ''

const positiveInteger = (value: unknown, fallback: number): number =>
  Number.isSafeInteger(value) && (value as number) > 0 ? (value as number) : fallback

const stringRecord = (value: unknown): Record<string, string> => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return {}
  return Object.fromEntries(
    Object.entries(value).flatMap(([key, item]) =>
      typeof item === 'string' && item.trim() ? [[key, item.trim()]] : [],
    ),
  )
}

const acceptanceRules = (value: unknown): PresentationDesignAcceptanceRule[] => {
  if (!Array.isArray(value)) return []
  return value.flatMap((raw) => {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return []
    const rule = raw as Record<string, unknown>
    const id = optionalText(rule.id, 80)
    const criterion = optionalText(rule.criterion, 500)
    return id && criterion ? [{ id, criterion }] : []
  })
}

/** Normalize either the current contract or the legacy plan_deck payload. */
export function parsePresentationDesignContract(value: unknown): PresentationDesignContract {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('invalid_presentation_design_contract')
  const raw = value as Record<string, unknown>
  if (!('schemaVersion' in raw)) {
    const legacy = parsePresentationDesignPlan(value)
    const assets = legacy.pages.flatMap((page, pageIndex) =>
      page.image_queries.map((query, queryIndex) => ({
        id: `slide-${pageIndex + 1}-asset-${queryIndex + 1}`,
        slideNumbers: [pageIndex + 1],
        type: 'image',
        role: 'substantive',
        intent: query,
        source: '',
        crop: '',
        placement: '',
        status: 'needed' as const,
      })),
    )
    return {
      schemaVersion: 1,
      revision: 1,
      status: 'draft',
      prototypePages: legacy.prototype_pages.map((index) => index + 1),
      brief: {
        topic: '',
        audience: '',
        occasion: '',
        desiredOutcome: '',
        language: '',
        pageCount: legacy.pages.length,
        aspectRatio: '',
        sourceConstraints: [],
      },
      narrative: {
        coreHook: legacy.core_hook,
        opening: '',
        development: '',
        tension: '',
        resolution: '',
        closingAction: '',
      },
      visualSystem: {
        style: legacy.style,
        colors: {},
        typography: {},
        safeMargin: '',
        grid: '',
        imageTreatment: '',
        chartTreatment: '',
        antiPatterns: [],
      },
      slides: legacy.pages.map((page, index) => ({
        number: index + 1,
        title: page.title,
        role: page.purpose,
        claim: page.brief,
        content: [page.brief],
        evidence: page.evidence,
        visualRoute: page.visual,
        layoutFamily: page.layout,
        focalVisual: page.visual,
        density: page.density,
        assetIds: assets.filter((asset) => asset.slideNumbers.includes(index + 1)).map((a) => a.id),
        acceptance: page.acceptance.map((criterion, acceptanceIndex) => ({
          id: `A${index + 1}.${acceptanceIndex + 1}`,
          criterion,
        })),
      })),
      assets,
      deckAcceptance: [],
    }
  }

  if (raw.schemaVersion !== 1) throw new Error('unsupported_presentation_design_schema')
  const brief = (raw.brief ?? {}) as Record<string, unknown>
  const narrative = (raw.narrative ?? {}) as Record<string, unknown>
  const visual = (raw.visualSystem ?? {}) as Record<string, unknown>
  const discovery = raw.discovery as Record<string, unknown> | undefined
  const statuses: PresentationDesignStatus[] = ['draft', 'ready', 'producing', 'verified']
  const assetStatuses: PresentationAssetStatus[] = [
    'needed',
    'searching',
    'downloaded',
    'validated',
    'ready',
    'fallback_ready',
  ]
  const contractStatus = raw.status ?? 'draft'
  if (!statuses.includes(contractStatus as PresentationDesignStatus))
    throw new Error('invalid_presentation_design_contract')
  const rawSlides = raw.slides ?? []
  const rawAssets = raw.assets ?? []
  if (!Array.isArray(rawSlides) || !Array.isArray(rawAssets))
    throw new Error('invalid_presentation_design_contract')
  const slides = rawSlides.map((item, index) => {
    if (!item || typeof item !== 'object' || Array.isArray(item))
      throw new Error('invalid_presentation_design_contract')
    const slide = item as Record<string, unknown>
    const density = slide.density as PresentationDesignPagePlan['density']
    if (!['low', 'medium', 'high'].includes(density))
      throw new Error('invalid_presentation_design_contract')
    return {
      number: positiveInteger(slide.number, index + 1),
      title: optionalText(slide.title, 300),
      role: optionalText(slide.role, 500),
      claim: optionalText(slide.claim, 2_000),
      content: texts(slide.content ?? [], 20, 2_000, false),
      evidence: texts(slide.evidence ?? [], 20, 1_000, false),
      visualRoute: optionalText(slide.visualRoute, 1_000),
      layoutFamily: optionalText(slide.layoutFamily, 100),
      focalVisual: optionalText(slide.focalVisual, 1_000),
      density,
      assetIds: texts(slide.assetIds ?? [], 20, 100, false),
      acceptance: acceptanceRules(slide.acceptance),
    }
  })
  const assets = rawAssets.map((item) => {
    if (!item || typeof item !== 'object' || Array.isArray(item))
      throw new Error('invalid_presentation_design_contract')
    const asset = item as Record<string, unknown>
    const status = asset.status as PresentationAssetStatus
    if (!assetStatuses.includes(status)) throw new Error('invalid_presentation_design_contract')
    return {
      id: optionalText(asset.id, 100),
      slideNumbers: Array.isArray(asset.slideNumbers)
        ? asset.slideNumbers.filter(
            (number): number is number => Number.isSafeInteger(number) && Number(number) > 0,
          )
        : [],
      type: optionalText(asset.type, 100),
      role: optionalText(asset.role, 100),
      intent: optionalText(asset.intent, 1_000),
      source: optionalText(asset.source, 2_000),
      crop: optionalText(asset.crop, 200),
      placement: optionalText(asset.placement, 500),
      status,
      ...(optionalText(asset.localReference, 2_000)
        ? { localReference: optionalText(asset.localReference, 2_000) }
        : {}),
      ...(optionalText(asset.fallback, 2_000)
        ? { fallback: optionalText(asset.fallback, 2_000) }
        : {}),
    }
  })
  return {
    schemaVersion: 1,
    revision: positiveInteger(raw.revision, 1),
    status: contractStatus as PresentationDesignStatus,
    prototypePages: Array.isArray(raw.prototypePages)
      ? raw.prototypePages.filter(
          (number): number is number => Number.isSafeInteger(number) && Number(number) > 0,
        )
      : [],
    ...(discovery
      ? {
          discovery: {
            questionnaire: texts(discovery.questionnaire ?? [], 30, 1_000, false),
            openQuestions: texts(discovery.openQuestions ?? [], 30, 1_000, false),
            researchNotes: texts(discovery.researchNotes ?? [], 100, 2_000, false),
          },
        }
      : {}),
    brief: {
      topic: optionalText(brief.topic, 500),
      audience: optionalText(brief.audience, 500),
      occasion: optionalText(brief.occasion, 500),
      desiredOutcome: optionalText(brief.desiredOutcome, 1_000),
      language: optionalText(brief.language, 100),
      pageCount: positiveInteger(brief.pageCount, slides.length),
      aspectRatio: optionalText(brief.aspectRatio, 50),
      sourceConstraints: texts(brief.sourceConstraints ?? [], 30, 1_000, false),
    },
    narrative: {
      coreHook: optionalText(narrative.coreHook, 1_000),
      opening: optionalText(narrative.opening, 1_000),
      development: optionalText(narrative.development, 2_000),
      tension: optionalText(narrative.tension, 1_000),
      resolution: optionalText(narrative.resolution, 1_000),
      closingAction: optionalText(narrative.closingAction, 1_000),
    },
    visualSystem: {
      style: optionalText(visual.style, 6_000),
      colors: stringRecord(visual.colors),
      typography: stringRecord(visual.typography),
      safeMargin: optionalText(visual.safeMargin, 100),
      grid: optionalText(visual.grid, 100),
      imageTreatment: optionalText(visual.imageTreatment, 1_000),
      chartTreatment: optionalText(visual.chartTreatment, 1_000),
      antiPatterns: texts(visual.antiPatterns ?? [], 30, 500, false),
    },
    slides,
    assets,
    deckAcceptance: acceptanceRules(raw.deckAcceptance),
  }
}

export function validatePresentationDesignReadiness(contract: PresentationDesignContract): {
  ready: boolean
  issues: string[]
} {
  const issues: string[] = []
  const required: Array<[string, string]> = [
    ['brief.topic', contract.brief.topic],
    ['brief.audience', contract.brief.audience],
    ['brief.occasion', contract.brief.occasion],
    ['brief.desiredOutcome', contract.brief.desiredOutcome],
    ['brief.language', contract.brief.language],
    ['brief.aspectRatio', contract.brief.aspectRatio],
    ['narrative.coreHook', contract.narrative.coreHook],
    ['narrative.opening', contract.narrative.opening],
    ['narrative.development', contract.narrative.development],
    ['narrative.tension', contract.narrative.tension],
    ['narrative.resolution', contract.narrative.resolution],
    ['narrative.closingAction', contract.narrative.closingAction],
    ['visualSystem.style', contract.visualSystem.style],
    ['visualSystem.safeMargin', contract.visualSystem.safeMargin],
    ['visualSystem.grid', contract.visualSystem.grid],
    ['visualSystem.imageTreatment', contract.visualSystem.imageTreatment],
    ['visualSystem.chartTreatment', contract.visualSystem.chartTreatment],
  ]
  for (const [field, value] of required) if (!value) issues.push(`${field} is required`)
  if (contract.slides.length !== contract.brief.pageCount)
    issues.push('brief.pageCount must match slides.length')
  if (contract.slides.length === 0) issues.push('slides must not be empty')
  if (Object.keys(contract.visualSystem.colors).length === 0)
    issues.push('visualSystem.colors must not be empty')
  if (Object.keys(contract.visualSystem.typography).length === 0)
    issues.push('visualSystem.typography must not be empty')
  if (contract.visualSystem.antiPatterns.length === 0)
    issues.push('visualSystem.antiPatterns must not be empty')
  const expectedPrototypes = Math.min(3, contract.slides.length)
  if (
    contract.prototypePages.length !== expectedPrototypes ||
    new Set(contract.prototypePages).size !== contract.prototypePages.length ||
    contract.prototypePages.some((number) => number < 1 || number > contract.slides.length)
  )
    issues.push(`prototypePages must contain ${expectedPrototypes} unique slide numbers`)
  const assets = new Map(contract.assets.map((asset) => [asset.id, asset]))
  contract.slides.forEach((slide, index) => {
    if (
      !slide.title ||
      !slide.role ||
      !slide.claim ||
      slide.content.length === 0 ||
      !slide.visualRoute ||
      !slide.layoutFamily ||
      !slide.focalVisual
    )
      issues.push(`slides[${index}] is missing its director plan`)
    if (slide.acceptance.length === 0) issues.push(`slides[${index}].acceptance must not be empty`)
    if (
      slide.assetIds.length > 0 &&
      !slide.assetIds.every((id) => {
        const asset = assets.get(id)
        return asset && ['ready', 'fallback_ready'].includes(asset.status)
      })
    )
      issues.push(`slides[${index}].assetIds must reference a ready asset`)
  })
  contract.assets.forEach((asset, index) => {
    if (!['ready', 'fallback_ready'].includes(asset.status))
      issues.push(`assets[${index}] must be ready or fallback_ready`)
    if (!asset.id || !asset.type || !asset.role || !asset.intent || !asset.crop || !asset.placement)
      issues.push(`assets[${index}] is missing its asset plan`)
    if (asset.status === 'ready' && (!asset.source || !asset.localReference))
      issues.push(`assets[${index}].ready requires source and localReference`)
    if (
      asset.status === 'fallback_ready' &&
      (!asset.fallback || (!asset.source && !asset.localReference))
    )
      issues.push(
        `assets[${index}].fallback_ready requires fallback and a validated source or local reference`,
      )
  })
  const acceptanceIds = [
    ...contract.slides.flatMap((slide) => slide.acceptance.map((rule) => rule.id)),
    ...contract.deckAcceptance.map((rule) => rule.id),
  ]
  if (new Set(acceptanceIds).size !== acceptanceIds.length)
    issues.push('acceptance ids must be unique')
  if (contract.deckAcceptance.length === 0) issues.push('deckAcceptance must not be empty')
  return { ready: issues.length === 0, issues }
}

const bullets = (items: string[]): string => items.map((item) => `- ${item}`).join('\n')
const DESIGN_CONTRACT_MARKER = 'WISWORK_PRESENTATION_DESIGN_CONTRACT:'
const MAX_EMBEDDED_DESIGN_CONTRACT_CHARS = 2_000_000

/** Restore a structured snapshot while leaving old plain DESIGN.md files readable. */
export function extractPresentationDesignContract(
  designDocument: string,
): PresentationDesignContract | undefined {
  const marker = `<!-- ${DESIGN_CONTRACT_MARKER}`
  const start = designDocument.lastIndexOf(marker)
  if (start < 0) return undefined
  const encodedStart = start + marker.length
  const end = designDocument.indexOf(' -->', encodedStart)
  if (end < 0 || end - encodedStart > MAX_EMBEDDED_DESIGN_CONTRACT_CHARS) return undefined
  try {
    return parsePresentationDesignContract(
      JSON.parse(decodeURIComponent(designDocument.slice(encodedStart, end))) as unknown,
    )
  } catch {
    return undefined
  }
}

export function renderPresentationDesignContract(contract: PresentationDesignContract): string {
  const discovery = contract.discovery ?? {
    questionnaire: [],
    openQuestions: [],
    researchNotes: [],
  }
  const lines = [
    '# DESIGN.md',
    '',
    `Status: ${contract.status}`,
    `Revision: ${contract.revision}`,
    '',
    '## Brief',
    '',
    `- Topic: ${contract.brief.topic}`,
    `- Audience: ${contract.brief.audience}`,
    `- Occasion: ${contract.brief.occasion}`,
    `- Desired Outcome: ${contract.brief.desiredOutcome}`,
    `- Language: ${contract.brief.language}`,
    `- Page Count: ${contract.brief.pageCount}`,
    `- Aspect Ratio: ${contract.brief.aspectRatio}`,
    `- Source Constraints: ${contract.brief.sourceConstraints.join('; ') || 'none'}`,
    `- Prototype Pages: ${contract.prototypePages.join(', ')}`,
    '',
    '## Discovery Log',
    '',
    '- Questionnaire Choices:',
    ...(discovery.questionnaire.length
      ? discovery.questionnaire.map((item) => `  - ${item}`)
      : ['  - none']),
    '- Open Questions:',
    ...(discovery.openQuestions.length
      ? discovery.openQuestions.map((item) => `  - ${item}`)
      : ['  - none']),
    '- Research Notes:',
    ...(discovery.researchNotes.length
      ? discovery.researchNotes.map((item) => `  - ${item}`)
      : ['  - none']),
    '',
    '## Narrative',
    '',
    `- Core Hook: ${contract.narrative.coreHook}`,
    `- Opening: ${contract.narrative.opening}`,
    `- Development: ${contract.narrative.development}`,
    `- Tension: ${contract.narrative.tension}`,
    `- Resolution: ${contract.narrative.resolution}`,
    `- Closing Action: ${contract.narrative.closingAction}`,
    '',
    '## Visual System',
    '',
    `- Style: ${contract.visualSystem.style}`,
    ...Object.entries(contract.visualSystem.colors).map(
      ([name, value]) => `- Color ${name}: ${value}`,
    ),
    ...Object.entries(contract.visualSystem.typography).map(
      ([name, value]) => `- Typography ${name}: ${value}`,
    ),
    `- Safe Margin: ${contract.visualSystem.safeMargin}`,
    `- Grid: ${contract.visualSystem.grid}`,
    `- Image Treatment: ${contract.visualSystem.imageTreatment}`,
    `- Chart Treatment: ${contract.visualSystem.chartTreatment}`,
    '',
    '## Anti-patterns',
    '',
    bullets(contract.visualSystem.antiPatterns),
  ]
  for (const slide of contract.slides) {
    lines.push(
      '',
      `## Slide ${slide.number} — ${slide.title}`,
      '',
      `- Role: ${slide.role}`,
      `- Claim: ${slide.claim}`,
      '- Content:',
      ...slide.content.map((item) => `  - ${item}`),
      '- Evidence:',
      ...(slide.evidence.length ? slide.evidence.map((item) => `  - ${item}`) : ['  - none']),
      `- Visual Route: ${slide.visualRoute}`,
      `- Layout Family: ${slide.layoutFamily}`,
      `- Focal Visual: ${slide.focalVisual}`,
      `- Density: ${slide.density}`,
      `- Assets: ${slide.assetIds.join(', ') || 'none'}`,
      '- Acceptance:',
      ...slide.acceptance.map((rule) => `  - ${rule.id}: ${rule.criterion}`),
    )
  }
  lines.push('', '## Assets', '')
  for (const asset of contract.assets)
    lines.push(
      `### ${asset.id}`,
      '',
      `- Slides: ${asset.slideNumbers.join(', ') || 'none'}`,
      `- Type: ${asset.type}`,
      `- Role: ${asset.role}`,
      `- Intent: ${asset.intent}`,
      `- Source: ${asset.source}`,
      `- Crop: ${asset.crop}`,
      `- Placement: ${asset.placement}`,
      `- Status: ${asset.status}`,
      `- Local Reference: ${asset.localReference ?? 'none'}`,
      `- Fallback: ${asset.fallback ?? 'none'}`,
      '',
    )
  lines.push(
    '',
    '## Deck Acceptance',
    '',
    ...contract.deckAcceptance.map((r) => `- ${r.id}: ${r.criterion}`),
  )
  const snapshot = encodeURIComponent(JSON.stringify(contract))
  if (snapshot.length <= MAX_EMBEDDED_DESIGN_CONTRACT_CHARS)
    lines.push('', `<!-- ${DESIGN_CONTRACT_MARKER}${snapshot} -->`)
  return lines.join('\n').trim()
}

const deepFreeze = <T>(value: T): T => {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    Object.freeze(value)
    for (const child of Object.values(value)) deepFreeze(child)
  }
  return value
}

/** Lock a ready revision for production, then mark that same revision verified. */
export function transitionPresentationDesignContract(
  contract: PresentationDesignContract,
  nextStatus: 'producing' | 'verified',
): PresentationDesignContract {
  const valid =
    (contract.status === 'ready' && nextStatus === 'producing') ||
    (contract.status === 'producing' && nextStatus === 'verified')
  if (!valid) throw new Error('invalid_presentation_design_transition')
  if (nextStatus === 'producing' && !validatePresentationDesignReadiness(contract).ready)
    throw new Error('presentation_design_not_ready')
  return deepFreeze(parsePresentationDesignContract({ ...contract, status: nextStatus }))
}

/** Open a new editable draft revision and describe which prior checks became stale. */
export function revisePresentationDesignContract(
  contract: PresentationDesignContract,
  change: { reason: string; scope: PresentationDesignInvalidationScope },
): { contract: PresentationDesignContract; invalidation: PresentationDesignInvalidation } {
  if (!['producing', 'verified'].includes(contract.status))
    throw new Error('presentation_design_revision_not_locked')
  const reason = text(change.reason, 1_000)
  let scope: PresentationDesignInvalidationScope
  if (change.scope.type === 'slide') {
    const slideNumbers = [...new Set(change.scope.slideNumbers)].filter(
      (number) => Number.isSafeInteger(number) && number > 0 && number <= contract.slides.length,
    )
    if (slideNumbers.length === 0) throw new Error('invalid_presentation_design_invalidation')
    scope = { type: 'slide', slideNumbers }
  } else if (change.scope.type === 'global' || change.scope.type === 'narrative') {
    scope = { type: change.scope.type }
  } else {
    throw new Error('invalid_presentation_design_invalidation')
  }
  const revision = contract.revision + 1
  return {
    contract: parsePresentationDesignContract({ ...contract, revision, status: 'draft' }),
    invalidation: {
      previousRevision: contract.revision,
      revision,
      reason,
      scope,
    },
  }
}
