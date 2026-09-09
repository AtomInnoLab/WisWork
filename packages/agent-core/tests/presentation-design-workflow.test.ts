import { describe, expect, it } from 'vitest'
import {
  buildPresentationDesignDocument,
  extractPresentationDesignContract,
  extractPresentationDesignDocument,
  parsePresentationDesignContract,
  parsePresentationDesignPlan,
  PRESENTATION_DESIGN_WORKFLOW_PROMPT,
  PRESENTATION_DESIGN_CONTRACT_SCHEMA,
  revisePresentationDesignContract,
  renderPresentationDesignContract,
  transitionPresentationDesignContract,
  validatePresentationDesignReadiness,
} from '../src/presentation-design-workflow'

describe('presentation design workflow', () => {
  it('defines the shared design, prototype, batch, and review contract', () => {
    expect(PRESENTATION_DESIGN_WORKFLOW_PROMPT).toContain('DESIGN.md')
    expect(PRESENTATION_DESIGN_WORKFLOW_PROMPT).toContain('Create and show this draft immediately')
    expect(PRESENTATION_DESIGN_WORKFLOW_PROMPT).toContain('update the same draft')
    expect(PRESENTATION_DESIGN_WORKFLOW_PROMPT).toContain(
      'keep the page plan and asset lists empty',
    )
    expect(PRESENTATION_DESIGN_WORKFLOW_PROMPT).toContain('representative content page')
    expect(PRESENTATION_DESIGN_WORKFLOW_PROMPT).toContain('2–3 slides')
    expect(PRESENTATION_DESIGN_WORKFLOW_PROMPT).toContain('screenshot')
    expect(PRESENTATION_DESIGN_WORKFLOW_PROMPT).toContain('one focal visual')
  })

  it('exports one bounded, forward-compatible tool schema for modern contracts', () => {
    expect(PRESENTATION_DESIGN_CONTRACT_SCHEMA).toMatchObject({
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
        schemaVersion: { const: 1 },
        status: { enum: ['draft', 'ready', 'producing', 'verified'] },
        slides: { maxItems: 12 },
        assets: { maxItems: 120 },
      },
    })
    const properties = PRESENTATION_DESIGN_CONTRACT_SCHEMA.properties as Record<
      string,
      Record<string, unknown>
    >
    expect(properties.brief?.additionalProperties).toBe(true)
    expect(properties.slides?.items).toMatchObject({ additionalProperties: true })
    expect(properties.assets?.items).toMatchObject({ additionalProperties: true })
  })

  it('renders a user-editable design artifact', () => {
    expect(buildPresentationDesignDocument('  Background: #0A0A0A  ')).toBe(
      '# DESIGN.md\n\nBackground: #0A0A0A',
    )
    expect(() => buildPresentationDesignDocument('   ')).toThrow('empty_presentation_design')
  })

  it('extracts timeline snapshots from desktop prose and Office JSON', () => {
    expect(
      extractPresentationDesignDocument(
        'Plan confirmed:\n\n# DESIGN.md\n\nAccent: green\n\n# Deck Plan\nPage 1',
      ),
    ).toBe('# DESIGN.md\n\nAccent: green')
    expect(
      extractPresentationDesignDocument(
        JSON.stringify({ status: 'planned', designMd: '# DESIGN.md\n\nAccent: blue' }),
      ),
    ).toBe('# DESIGN.md\n\nAccent: blue')
    expect(extractPresentationDesignDocument('# DESIGN.md\n\n   ')).toBeUndefined()
  })

  it('normalizes one strict plan shared by every host', () => {
    const plan = parsePresentationDesignPlan({
      core_hook: ' One story ',
      style: ' Background: #000000 ',
      prototype_pages: [0],
      pages: [
        {
          title: ' Cover ',
          brief: ' Opening ',
          layout: 'cover',
          purpose: 'Open',
          visual: 'One hero',
          acceptance: ['Clear hierarchy'],
          density: 'low',
        },
      ],
    })
    expect(plan.core_hook).toBe('One story')
    expect(plan.pages[0]?.image_queries).toEqual([])
    expect(() =>
      parsePresentationDesignPlan({ ...plan, prototype_pages: [], pages: plan.pages }),
    ).toThrow('invalid_presentation_plan')
  })

  it('normalizes the legacy plan_deck payload into a versioned design contract', () => {
    const contract = parsePresentationDesignContract({
      core_hook: 'One story',
      style: 'Dark editorial',
      prototype_pages: [0],
      pages: [
        {
          title: 'Cover',
          brief: 'Opening',
          layout: 'cover',
          purpose: 'Open',
          visual: 'One hero',
          acceptance: ['Clear hierarchy'],
          density: 'low',
          future_page_field: true,
        },
      ],
      future_plan_field: true,
    })
    expect(contract).toMatchObject({
      schemaVersion: 1,
      revision: 1,
      status: 'draft',
      prototypePages: [1],
      narrative: { coreHook: 'One story' },
      visualSystem: { style: 'Dark editorial' },
    })
    expect(contract.slides[0]).toMatchObject({ number: 1, title: 'Cover' })
  })

  it('accepts a complete contract, ignores future fields, and validates production readiness', () => {
    const contract = parsePresentationDesignContract({
      schemaVersion: 1,
      revision: 2,
      status: 'ready',
      prototypePages: [1],
      discovery: {
        questionnaire: ['Audience: business users'],
        openQuestions: ['Confirm source licensing'],
        researchNotes: ['Official product screenshot selected'],
      },
      brief: {
        topic: 'LLM',
        audience: 'Business users',
        occasion: 'Internal briefing',
        desiredOutcome: 'Choose a pilot',
        language: 'zh-CN',
        pageCount: 1,
        aspectRatio: '16:9',
        sourceConstraints: ['Use primary sources'],
      },
      narrative: {
        coreHook: 'LLM is a new interface',
        opening: 'Start with ChatGPT',
        development: 'Definition to applications',
        tension: 'Hallucinations',
        resolution: 'Human oversight',
        closingAction: 'Pilot a verifiable workflow',
      },
      visualSystem: {
        style: 'Dark editorial',
        colors: { background: '#0A0A0A', text: '#FFFFFF' },
        typography: { title: '32 pt', body: '18 pt' },
        safeMargin: '64 px',
        grid: '12 columns',
        imageTreatment: 'High contrast',
        chartTreatment: 'Direct labels',
        antiPatterns: ['No repetitive card grids'],
      },
      slides: [
        {
          number: 1,
          title: 'LLM is a new interface',
          role: 'Opening',
          claim: 'LLM changes how people use software',
          content: ['Definition'],
          evidence: ['Source note'],
          visualRoute: 'Hero statement',
          layoutFamily: 'cover',
          focalVisual: 'Abstract network',
          density: 'low',
          assetIds: ['hero'],
          acceptance: [{ id: 'A1.1', criterion: 'Headline is immediately legible' }],
          futureSlideField: 'ignored',
        },
      ],
      assets: [
        {
          id: 'hero',
          slideNumbers: [1],
          type: 'generated-illustration',
          role: 'substantive',
          intent: 'Support the opening claim',
          source: 'generated',
          crop: '16:9',
          placement: 'full bleed',
          status: 'ready',
          localReference: 'asset://hero',
        },
      ],
      deckAcceptance: [{ id: 'D1', criterion: 'One claim per slide' }],
      futureContractField: { ignored: true },
    })

    expect(validatePresentationDesignReadiness(contract)).toEqual({ ready: true, issues: [] })
    expect(contract.brief.sourceConstraints).toEqual(['Use primary sources'])
    expect(contract.prototypePages).toEqual([1])
    expect(renderPresentationDesignContract(contract)).toContain(
      '## 第 1 页 — LLM is a new interface',
    )
    expect(renderPresentationDesignContract(contract)).toContain('## 调研记录')
    expect(renderPresentationDesignContract(contract)).toContain(
      'Official product screenshot selected',
    )
    expect(renderPresentationDesignContract(contract)).toContain('A1.1')
    expect(renderPresentationDesignContract(contract)).toContain('- 内容:')
    expect(renderPresentationDesignContract(contract)).toContain('  - Definition')
    expect(renderPresentationDesignContract(contract)).toContain('- 证据:')
    expect(renderPresentationDesignContract(contract)).toContain('- 来源: generated')
    expect(renderPresentationDesignContract(contract)).toContain('- 本地引用: asset://hero')
    expect(extractPresentationDesignContract(renderPresentationDesignContract(contract))).toEqual(
      contract,
    )
    const chinese = {
      ...contract,
      brief: { ...contract.brief, language: '中文' },
    }
    const chineseDesign = renderPresentationDesignContract(chinese)
    expect(chineseDesign).toContain('## 任务定义')
    expect(chineseDesign).toContain('## 调研记录')
    expect(chineseDesign).toContain('## 视觉系统')
    expect(chineseDesign).toContain('## 素材清单')
    expect(chineseDesign).toContain('## 整套验收标准')
    expect(extractPresentationDesignContract(chineseDesign)).toEqual(chinese)
    expect(extractPresentationDesignContract('# DESIGN.md\n\nLegacy prose')).toBeUndefined()
  })

  it('defaults missing optional schema-1 transport fields instead of rejecting older snapshots', () => {
    expect(parsePresentationDesignContract({ schemaVersion: 1 })).toMatchObject({
      schemaVersion: 1,
      revision: 1,
      status: 'draft',
      prototypePages: [],
      slides: [],
      assets: [],
      deckAcceptance: [],
    })
  })

  it('blocks incomplete assets, unsafe fallbacks, and duplicate acceptance ids', () => {
    const contract = parsePresentationDesignContract({
      schemaVersion: 1,
      status: 'ready',
      prototypePages: [1],
      brief: {
        topic: 'LLM',
        audience: 'Business users',
        occasion: 'Briefing',
        desiredOutcome: 'Choose a pilot',
        language: 'zh-CN',
        pageCount: 1,
        aspectRatio: '16:9',
        sourceConstraints: [],
      },
      narrative: {
        coreHook: 'A new interface',
        opening: 'Familiar example',
        development: 'Explain value',
        tension: 'Risk',
        resolution: 'Controls',
        closingAction: 'Pilot',
      },
      visualSystem: {
        style: 'Dark editorial',
        colors: { background: '#000' },
        typography: { title: '32pt' },
        safeMargin: '64px',
        grid: '12 columns',
        imageTreatment: 'High contrast',
        chartTreatment: 'Direct labels',
        antiPatterns: ['No repetitive cards'],
      },
      slides: [
        {
          number: 1,
          title: 'Cover',
          role: 'Opening',
          claim: 'A new interface',
          content: ['Definition'],
          evidence: [],
          visualRoute: 'Hero',
          layoutFamily: 'cover',
          focalVisual: 'Network',
          density: 'low',
          assetIds: ['hero'],
          acceptance: [{ id: 'D1', criterion: 'Readable' }],
        },
      ],
      assets: [
        {
          id: 'hero',
          slideNumbers: [1],
          type: 'image',
          role: 'substantive',
          intent: 'Hero',
          source: '',
          crop: '16:9',
          placement: 'full bleed',
          status: 'fallback_ready',
        },
        {
          id: 'unused',
          slideNumbers: [],
          type: 'image',
          role: 'substantive',
          intent: 'Inventory item',
          source: 'official',
          crop: '1:1',
          placement: 'right',
          status: 'needed',
        },
      ],
      deckAcceptance: [{ id: 'D1', criterion: 'One claim per slide' }],
    })
    expect(validatePresentationDesignReadiness(contract).issues).toEqual(
      expect.arrayContaining([
        'assets[0].fallback_ready requires fallback and a validated source or local reference',
        'assets[1] must be ready or fallback_ready',
        'acceptance ids must be unique',
      ]),
    )
  })

  it('reports actionable readiness issues before production', () => {
    const contract = parsePresentationDesignContract({
      core_hook: 'One story',
      style: 'Dark editorial',
      prototype_pages: [0],
      pages: [
        {
          title: 'Cover',
          brief: 'Opening',
          layout: 'cover',
          purpose: 'Open',
          visual: 'One hero',
          acceptance: ['Clear hierarchy'],
          density: 'low',
          image_queries: ['abstract network'],
        },
      ],
    })
    expect(validatePresentationDesignReadiness(contract)).toEqual({
      ready: false,
      issues: expect.arrayContaining([
        'brief.topic is required',
        'slides[0].assetIds must reference a ready asset',
      ]),
    })
  })

  it('immutably locks a ready revision for production and verifies it', () => {
    const draft = parsePresentationDesignContract({
      core_hook: 'One story',
      style: 'Dark editorial',
      prototype_pages: [0],
      pages: [
        {
          title: 'Cover',
          brief: 'Opening',
          layout: 'cover',
          purpose: 'Open',
          visual: 'One hero',
          acceptance: ['Clear hierarchy'],
          density: 'low',
        },
      ],
    })
    const ready = {
      ...draft,
      status: 'ready' as const,
      brief: {
        ...draft.brief,
        topic: 'LLM',
        audience: 'Business users',
        occasion: 'Briefing',
        desiredOutcome: 'Choose a pilot',
        language: 'zh-CN',
        aspectRatio: '16:9',
      },
      narrative: {
        ...draft.narrative,
        opening: 'Familiar example',
        development: 'Explain value',
        tension: 'Risk',
        resolution: 'Controls',
        closingAction: 'Pilot',
      },
      visualSystem: {
        ...draft.visualSystem,
        colors: { background: '#000' },
        typography: { title: '32pt' },
        safeMargin: '64px',
        grid: '12 columns',
        imageTreatment: 'High contrast',
        chartTreatment: 'Direct labels',
        antiPatterns: ['No repetitive cards'],
      },
      deckAcceptance: [{ id: 'D1', criterion: 'One claim per slide' }],
    }
    const producing = transitionPresentationDesignContract(ready, 'producing')
    expect(producing).not.toBe(ready)
    expect(producing.status).toBe('producing')
    expect(Object.isFrozen(producing)).toBe(true)
    expect(Object.isFrozen(producing.slides)).toBe(true)
    expect(ready.status).toBe('ready')
    expect(transitionPresentationDesignContract(producing, 'verified').status).toBe('verified')
    expect(() => transitionPresentationDesignContract(ready, 'verified')).toThrow(
      'invalid_presentation_design_transition',
    )
    expect(() =>
      transitionPresentationDesignContract({ ...draft, status: 'ready' }, 'producing'),
    ).toThrow('presentation_design_not_ready')
  })

  it('revises a locked contract into a new draft with explicit invalidation scope', () => {
    const producing = {
      ...parsePresentationDesignContract({
        core_hook: 'One story',
        style: 'Dark editorial',
        prototype_pages: [0],
        pages: [
          {
            title: 'Cover',
            brief: 'Opening',
            layout: 'cover',
            purpose: 'Open',
            visual: 'One hero',
            acceptance: ['Clear hierarchy'],
            density: 'low',
          },
        ],
      }),
      revision: 3,
      status: 'producing' as const,
    }
    const revised = revisePresentationDesignContract(producing, {
      reason: 'Hero crop does not fit',
      scope: { type: 'slide', slideNumbers: [1] },
    })
    expect(revised.contract).toMatchObject({ revision: 4, status: 'draft' })
    expect(revised.contract).not.toBe(producing)
    expect(revised.invalidation).toEqual({
      previousRevision: 3,
      revision: 4,
      reason: 'Hero crop does not fit',
      scope: { type: 'slide', slideNumbers: [1] },
    })
    expect(() =>
      revisePresentationDesignContract(
        { ...producing, status: 'draft' },
        {
          reason: 'No-op',
          scope: { type: 'global' },
        },
      ),
    ).toThrow('presentation_design_revision_not_locked')
  })
})
