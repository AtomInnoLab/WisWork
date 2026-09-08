export const PRESENTATION_DESIGN_WORKFLOW_PROMPT = `
## Shared presentation design workflow
For a whole-deck creation or redesign, work as a presentation director, not a text filler:
1. Brief: establish audience, occasion, desired decision, page count, source constraints, and one narrative conclusion. Ask only for missing choices that materially change the result.
2. DESIGN.md contract: before editing, state a user-editable design contract covering concrete color tokens, typography hierarchy, safe margins/grid, image treatment, density limits, layout families, and composition rules. Treat it as binding for every slide.
3. Deck plan: give every slide one conclusion-led headline, a narrative role, one focal visual, supporting evidence, a layout family, asset needs/provenance, density, and slide-specific acceptance criteria. Do not use the same layout family on adjacent slides unless continuity requires it.
4. Asset plan: search real people, products, places, brands, and current facts; generate only abstract/custom illustration when generation exists. Use native editable charts for data. Never invent precise data or image URLs.
5. Prototype gate: first create or identify the cover, one representative content page, and one complex visual page. Screenshot and review those pages before continuing the remaining production batches. When fewer than three slides are requested, review every slide.
6. Production loop: complete the remaining work in batches of 2–3 slides. After each batch, inspect screenshots plus geometry, repair concrete defects, and re-screenshot changed slides before continuing.
7. Quality bar: each slide needs one clear conclusion and one focal visual; readable hierarchy, deliberate whitespace, aligned geometry, sufficient contrast, relevant imagery, and no accidental overflow, overlap, distortion, placeholder content, or repetitive card grids. Review design-system consistency and rhythm across adjacent slides, then run final whole-deck verification.
Keep all edits native, editable, reversible, and within the host's permission model. If a capability is unavailable, use the declared fallback and report the specific unresolved limitation.
`.trim()

export function buildPresentationDesignDocument(style: string): string {
  return `# DESIGN.md\n\n${style.trim()}`
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
  if (
    Object.keys(plan).some(
      (key) => !['core_hook', 'style', 'pages', 'prototype_pages'].includes(key),
    )
  )
    throw new Error('invalid_presentation_plan')
  if (!Array.isArray(plan.pages) || plan.pages.length === 0 || plan.pages.length > 12)
    throw new Error('invalid_presentation_plan')
  const pages = plan.pages.map((raw) => {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw))
      throw new Error('invalid_presentation_plan')
    const page = raw as Record<string, unknown>
    if (
      Object.keys(page).some(
        (key) =>
          ![
            'title',
            'type',
            'brief',
            'layout',
            'purpose',
            'visual',
            'evidence',
            'acceptance',
            'density',
            'image_queries',
          ].includes(key),
      ) ||
      !['low', 'medium', 'high'].includes(String(page.density))
    )
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
