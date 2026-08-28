import { describe, expect, it } from 'vitest'
import { buildQualityIdentityMap } from '../src/main/operations/quality-identity'

describe('quality identity projection', () => {
  it('recurses groups, keeps durable creation IDs distinct, and skips legacy runtime-only IDs', () => {
    const child = { id: 'runtime-child', creationId: 'creation-child', type: 'shape' }
    const slide = {
      durableId: 'ppt/slides/slide1.xml',
      elements: [
        { id: 'runtime-a', creationId: 'creation-a', type: 'shape' },
        { id: 'legacy-only', type: 'shape' },
        { id: 'group', creationId: 'creation-group', type: 'group', children: [child] },
      ],
    }
    const result = buildQualityIdentityMap(slide as never)
    expect(result).toEqual({
      slideId: 'ppt/slides/slide1.xml',
      elementIds: {
        'runtime-a': 'creation-a',
        group: 'creation-group',
        'runtime-child': 'creation-child',
      },
      truncated: false,
    })
    expect(result.elementIds).not.toHaveProperty('legacy-only')
  })
})
