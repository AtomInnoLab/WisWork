import { describe, expect, it } from 'vitest'
import {
  buildPresentationDesignDocument,
  extractPresentationDesignDocument,
  parsePresentationDesignPlan,
  PRESENTATION_DESIGN_WORKFLOW_PROMPT,
} from '../src/presentation-design-workflow'

describe('presentation design workflow', () => {
  it('defines the shared design, prototype, batch, and review contract', () => {
    expect(PRESENTATION_DESIGN_WORKFLOW_PROMPT).toContain('DESIGN.md')
    expect(PRESENTATION_DESIGN_WORKFLOW_PROMPT).toContain('representative content page')
    expect(PRESENTATION_DESIGN_WORKFLOW_PROMPT).toContain('2–3 slides')
    expect(PRESENTATION_DESIGN_WORKFLOW_PROMPT).toContain('screenshot')
    expect(PRESENTATION_DESIGN_WORKFLOW_PROMPT).toContain('one focal visual')
  })

  it('renders a user-editable design artifact', () => {
    expect(buildPresentationDesignDocument('  Background: #0A0A0A  ')).toBe(
      '# DESIGN.md\n\nBackground: #0A0A0A',
    )
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
})
