import { describe, expect, it } from 'vitest'
import { renderPresentationDesignContract } from '@wiswork/agent-core'
import { presentationDesignLifecycle } from '../src/renderer/ai/presentation-design-ui'

const contract = {
  schemaVersion: 1 as const,
  revision: 4,
  status: 'producing' as const,
  prototypePages: [1],
  brief: {
    topic: 'T',
    audience: 'A',
    occasion: 'O',
    desiredOutcome: 'D',
    language: 'en',
    pageCount: 1,
    aspectRatio: '16:9',
    sourceConstraints: [],
  },
  narrative: {
    coreHook: 'H',
    opening: 'O',
    development: 'D',
    tension: 'T',
    resolution: 'R',
    closingAction: 'C',
  },
  visualSystem: {
    style: 'S',
    colors: {},
    typography: {},
    safeMargin: '1',
    grid: '1',
    imageTreatment: 'I',
    chartTreatment: 'C',
    antiPatterns: [],
  },
  slides: [
    {
      number: 1,
      title: 'T',
      role: 'R',
      claim: 'C',
      content: [],
      evidence: [],
      visualRoute: 'V',
      layoutFamily: 'cover',
      focalVisual: 'F',
      density: 'low' as const,
      assetIds: [],
      acceptance: [{ id: 'A1', criterion: 'C' }],
    },
  ],
  assets: [],
  deckAcceptance: [{ id: 'D1', criterion: 'C' }],
}

describe('presentation design timeline lifecycle', () => {
  it('labels producing and verified snapshots and keeps them read-only', () => {
    expect(presentationDesignLifecycle(renderPresentationDesignContract(contract))).toEqual({
      label: 'DESIGN.md · locked · r4',
      editable: false,
    })
    expect(
      presentationDesignLifecycle(
        renderPresentationDesignContract({ ...contract, status: 'verified' }),
      ),
    ).toEqual({ label: 'DESIGN.md · verified · r4', editable: false })
  })
})
