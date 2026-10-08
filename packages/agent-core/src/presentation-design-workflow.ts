export const PRESENTATION_DESIGN_WORKFLOW_PROMPT = `
## Shared presentation design workflow
For a whole-deck creation or redesign, work as a presentation director, not a text filler:
1. Brief: establish audience, occasion, desired decision, page count, source constraints, and one narrative conclusion. Ask only for missing choices that materially change the result.
2. DESIGN.md contract: immediately after the brief or questionnaire, draft a user-editable design contract covering the confirmed choices, open questions, research plan, concrete color tokens, typography hierarchy, safe margins/grid, image treatment, density limits, layout families, and composition rules. Create and show this draft immediately, before fact or image research. Keep the first draft compact: record the brief and discovery log, but keep the page plan and asset lists empty until research supplies them. Treat it as binding for every slide. As facts and assets are collected, update the same draft with evidence, provenance, and asset status. Do not mark the final contract ready until required fact research, image search, and asset validation are complete.
Write every user-facing DESIGN.md field in the user's language. Search queries and source titles may remain in their source language, but explain their intent, placement, and status in the user's language.
3. Deck plan: give every slide one conclusion-led headline, a narrative role, one focal visual, supporting evidence, a layout family, asset needs/provenance, density, and slide-specific acceptance criteria. Do not use the same layout family on adjacent slides unless continuity requires it.
4. Asset plan: search real people, products, places, brands, and current facts; validate selected image sources before finalizing DESIGN.md as ready. Generate only abstract/custom illustration when generation exists. Use native editable charts for data. Never invent precise data or image URLs.
Each plan_deck submission replaces the complete contract; it is not an asset patch. Keep research candidates in the draft, but include only selected production assets in the final ready inventory and preserve research history in discovery.researchNotes. Remove unused candidates, never required visuals merely to pass the gate. For each selected image, source records provenance (use source_url/sourceUrl when returned); localReference is the exact direct image_url/imageUrl returned by image_search, or a host-approved local asset reference, not the source webpage. Search results alone do not prove downloaded image usability. Asset status validated is not ready: after the required validation and placement decisions, explicitly submit ready with both source and localReference. Do not relabel unresolved or unvalidated assets to bypass the gate. A native fallback must be explicit, usable, and recorded with fallback_ready, its fallback description, and provenance or an approved reference. If validation is unavailable, keep the contract draft or use that explicit native fallback and explain the limitation. No hidden asset-finalization tool runs after plan_deck.
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

const normalizePresentationDesignDocument = (value: string): string =>
  !value.includes('\n') &&
  /^\s*#\s*DESIGN\.md\\n/i.test(value) &&
  (value.match(/\\n/g) ?? []).length > 1
    ? value.replace(/\\r\\n|\\n/g, '\n')
    : value

/** Extract a design snapshot from either desktop prose or Office JSON tool output. */
export function extractPresentationDesignDocument(output: string): string | undefined {
  try {
    const value = JSON.parse(output) as { designMd?: unknown }
    if (typeof value.designMd === 'string') {
      const designMd = normalizePresentationDesignDocument(value.designMd)
      if (hasPresentationDesignBody(designMd)) return designMd.trim()
    }
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
const stringArray = (maxItems: number | undefined, maxLength: number) => ({
  type: 'array',
  ...(maxItems === undefined ? {} : { maxItems }),
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
          assetIds: stringArray(undefined, 100),
          acceptance: { type: 'array', minItems: 1, maxItems: 20, items: acceptanceSchema() },
        },
      },
    },
    assets: {
      type: 'array',
      description:
        'Full replacement inventory, not a patch. Drafts may contain research candidates. A ready contract contains only selected production assets, all ready or fallback_ready; preserve unused research in discovery.researchNotes.',
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
          source: {
            ...boundedString(2_000),
            description:
              'Provenance: the image search source_url/sourceUrl or other verified origin. This is not necessarily a usable image reference.',
          },
          crop: boundedString(200),
          placement: boundedString(500),
          status: {
            type: 'string',
            enum: ['needed', 'searching', 'downloaded', 'validated', 'ready', 'fallback_ready'],
            description:
              'validated is still pre-production. Explicitly use ready only after validation with source, localReference, crop, and placement complete; fallback_ready requires an explicit usable fallback. Never promote an unresolved asset just to pass readiness.',
          },
          localReference: {
            ...boundedString(2_000),
            description:
              'Required for ready: exact direct image_url/imageUrl returned by image_search, or a host-approved local asset reference. Do not substitute a source webpage or invent a URL.',
          },
          fallback: {
            ...boundedString(2_000),
            description:
              'Required for fallback_ready: describe the actual usable native chart/vector/text alternative, with source or localReference. Update the slide visual plan to match.',
          },
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
  maxItems: number | undefined,
  maxLength: number,
  required: boolean,
): string[] => {
  if (
    !Array.isArray(value) ||
    (required && value.length === 0) ||
    (maxItems !== undefined && value.length > maxItems)
  )
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
      image_queries: texts(page.image_queries ?? [], undefined, 200, false),
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
      assetIds: texts(slide.assetIds ?? [], undefined, 100, false),
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

/** Explain a rejected submission without changing the contract or claiming asset validation. */
export function formatPresentationDesignReadinessFailure(
  contract: PresentationDesignContract,
  issues: readonly string[],
): string {
  const details = contract.assets.flatMap((asset, index) => {
    const missingForReady = (['id', 'type', 'role', 'intent', 'crop', 'placement'] as const).filter(
      (field) => !asset[field],
    ) as string[]
    if (asset.status === 'fallback_ready') {
      if (!asset.fallback) missingForReady.push('fallback')
      if (!asset.source && !asset.localReference) missingForReady.push('source or localReference')
    } else {
      if (!asset.source) missingForReady.push('source')
      if (!asset.localReference) missingForReady.push('localReference')
    }
    if (['ready', 'fallback_ready'].includes(asset.status) && missingForReady.length === 0)
      return []
    return [
      {
        path: `assets[${index}]`,
        id: asset.id,
        status: asset.status,
        source: Boolean(asset.source),
        localReference: Boolean(asset.localReference),
        missingForReady,
      },
    ]
  })
  const recovery = [
    ...(details.length
      ? [
          'For ready assets, source is provenance; localReference must be the exact returned image_url/imageUrl or a host-approved asset reference, not a webpage. validated does not mean ready. Complete validation and placement before explicitly submitting ready; never promote an unresolved asset just to pass the gate. If unavailable, retain draft or declare a usable native fallback with fallback_ready, fallback, and source or localReference. Remove only unused candidates from the final production inventory; preserve research notes and required visuals.',
        ]
      : []),
    'Resubmit the full corrected contract with plan_deck; no asset-finalization tool runs automatically. The active contract is unchanged and this submission has not authorized production.',
  ].join('\n')
  const heading = details.length
    ? 'Asset readiness details (reference presence is not proof of validation):'
    : ''
  const omittedSummary = (issueCount: number, assetCount: number) =>
    `Omitted ${issueCount} readiness issues; ${assetCount} asset details to keep this response bounded.`
  const encoder = new TextEncoder()
  // Reserve repair instructions and the largest possible omission counts before adding diagnostics.
  const reserved = encoder.encode(
    `${heading}\n${omittedSummary(issues.length, details.length)}\n${recovery}\n`,
  ).byteLength
  const takeWithin = (lines: readonly string[], budget: number) => {
    const kept: string[] = []
    let bytes = 0
    for (const line of lines) {
      const size = encoder.encode(line).byteLength + 1
      if (bytes + size > budget) break
      kept.push(line)
      bytes += size
    }
    return { text: kept.join('\n'), count: kept.length, bytes }
  }
  const issueBlock = takeWithin(issues, Math.min(4_000, 16_000 - reserved))
  const assetBlock = takeWithin(
    details.map((detail) => JSON.stringify(detail)),
    16_000 - reserved - issueBlock.bytes,
  )
  return [
    issueBlock.text,
    heading,
    assetBlock.text,
    omittedSummary(issues.length - issueBlock.count, details.length - assetBlock.count),
    recovery,
  ]
    .filter(Boolean)
    .join('\n')
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
  const chinese = /(?:中文|Chinese|zh(?:-|_|$))/i.test(contract.brief.language)
  const l = (english: string, localized: string): string => (chinese ? localized : english)
  const none = l('none', '无')
  const discovery = contract.discovery ?? {
    questionnaire: [],
    openQuestions: [],
    researchNotes: [],
  }
  const lines = [
    '# DESIGN.md',
    '',
    `${l('Status', '状态')}: ${contract.status}`,
    `${l('Revision', '修订版本')}: ${contract.revision}`,
    '',
    `## ${l('Brief', '任务定义')}`,
    '',
    `- ${l('Topic', '主题')}: ${contract.brief.topic}`,
    `- ${l('Audience', '受众')}: ${contract.brief.audience}`,
    `- ${l('Occasion', '场景')}: ${contract.brief.occasion}`,
    `- ${l('Desired Outcome', '预期结果')}: ${contract.brief.desiredOutcome}`,
    `- ${l('Language', '语言')}: ${contract.brief.language}`,
    `- ${l('Page Count', '页数')}: ${contract.brief.pageCount}`,
    `- ${l('Aspect Ratio', '页面比例')}: ${contract.brief.aspectRatio}`,
    `- ${l('Source Constraints', '来源约束')}: ${contract.brief.sourceConstraints.join('; ') || none}`,
    `- ${l('Prototype Pages', '原型页')}: ${contract.prototypePages.join(', ')}`,
    '',
    `## ${l('Discovery Log', '调研记录')}`,
    '',
    `- ${l('Questionnaire Choices', '问卷选择')}:`,
    ...(discovery.questionnaire.length
      ? discovery.questionnaire.map((item) => `  - ${item}`)
      : [`  - ${none}`]),
    `- ${l('Open Questions', '待确认问题')}:`,
    ...(discovery.openQuestions.length
      ? discovery.openQuestions.map((item) => `  - ${item}`)
      : [`  - ${none}`]),
    `- ${l('Research Notes', '调研笔记')}:`,
    ...(discovery.researchNotes.length
      ? discovery.researchNotes.map((item) => `  - ${item}`)
      : [`  - ${none}`]),
    '',
    `## ${l('Narrative', '故事线')}`,
    '',
    `- ${l('Core Hook', '核心主张')}: ${contract.narrative.coreHook}`,
    `- ${l('Opening', '开场')}: ${contract.narrative.opening}`,
    `- ${l('Development', '展开')}: ${contract.narrative.development}`,
    `- ${l('Tension', '矛盾')}: ${contract.narrative.tension}`,
    `- ${l('Resolution', '解决')}: ${contract.narrative.resolution}`,
    `- ${l('Closing Action', '结尾行动')}: ${contract.narrative.closingAction}`,
    '',
    `## ${l('Visual System', '视觉系统')}`,
    '',
    `- ${l('Style', '风格')}: ${contract.visualSystem.style}`,
    ...Object.entries(contract.visualSystem.colors).map(
      ([name, value]) => `- ${l('Color', '颜色')} ${name}: ${value}`,
    ),
    ...Object.entries(contract.visualSystem.typography).map(
      ([name, value]) => `- ${l('Typography', '字体')} ${name}: ${value}`,
    ),
    `- ${l('Safe Margin', '安全边距')}: ${contract.visualSystem.safeMargin}`,
    `- ${l('Grid', '网格')}: ${contract.visualSystem.grid}`,
    `- ${l('Image Treatment', '图片处理')}: ${contract.visualSystem.imageTreatment}`,
    `- ${l('Chart Treatment', '图表处理')}: ${contract.visualSystem.chartTreatment}`,
    '',
    `## ${l('Anti-patterns', '禁止项')}`,
    '',
    bullets(contract.visualSystem.antiPatterns),
  ]
  for (const slide of contract.slides) {
    lines.push(
      '',
      `## ${l('Slide', '第')} ${slide.number}${chinese ? ' 页' : ''} — ${slide.title}`,
      '',
      `- ${l('Role', '页面作用')}: ${slide.role}`,
      `- ${l('Claim', '核心结论')}: ${slide.claim}`,
      `- ${l('Content', '内容')}:`,
      ...slide.content.map((item) => `  - ${item}`),
      `- ${l('Evidence', '证据')}:`,
      ...(slide.evidence.length ? slide.evidence.map((item) => `  - ${item}`) : [`  - ${none}`]),
      `- ${l('Visual Route', '视觉路线')}: ${slide.visualRoute}`,
      `- ${l('Layout Family', '版式类型')}: ${slide.layoutFamily}`,
      `- ${l('Focal Visual', '焦点视觉')}: ${slide.focalVisual}`,
      `- ${l('Density', '信息密度')}: ${slide.density}`,
      `- ${l('Assets', '素材')}: ${slide.assetIds.join(', ') || none}`,
      `- ${l('Acceptance', '验收标准')}:`,
      ...slide.acceptance.map((rule) => `  - ${rule.id}: ${rule.criterion}`),
    )
  }
  lines.push('', `## ${l('Assets', '素材清单')}`, '')
  for (const asset of contract.assets)
    lines.push(
      `### ${asset.id}`,
      '',
      `- ${l('Slides', '使用页面')}: ${asset.slideNumbers.join(', ') || none}`,
      `- ${l('Type', '类型')}: ${asset.type}`,
      `- ${l('Role', '作用')}: ${asset.role}`,
      `- ${l('Intent', '用途')}: ${asset.intent}`,
      `- ${l('Source', '来源')}: ${asset.source}`,
      `- ${l('Crop', '裁切')}: ${asset.crop}`,
      `- ${l('Placement', '位置')}: ${asset.placement}`,
      `- ${l('Status', '状态')}: ${asset.status}`,
      `- ${l('Local Reference', '本地引用')}: ${asset.localReference ?? none}`,
      `- ${l('Fallback', '备用方案')}: ${asset.fallback ?? none}`,
      '',
    )
  lines.push(
    '',
    `## ${l('Deck Acceptance', '整套验收标准')}`,
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
