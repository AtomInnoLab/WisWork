import { describe, expect, it } from 'vitest'
import {
  affectedStyleSlideIds,
  parsePowerPointStyleDependencies,
} from '../src/skills/powerpoint/presentation-style-dependencies.js'

const slides = [
  { slideId: 'b', masterId: 'm1', layoutId: 'l2' },
  { slideId: 'a', masterId: 'm1', layoutId: 'l1' },
  { slideId: 'c', masterId: 'm2', layoutId: 'l1' },
]
describe('native presentation style dependencies', () => {
  it('copies and sorts complete snapshots, retaining empty presentations', () => {
    const parsed = parsePowerPointStyleDependencies({ slides })
    expect(parsed.slides.map((slide) => slide.slideId)).toEqual(['a', 'b', 'c'])
    expect(parsed.slides[0]).not.toBe(slides[1])
    expect(parsePowerPointStyleDependencies({ slides: [] })).toEqual({ slides: [] })
  })
  it('retains the complete dependency graph beyond old 100/512 page windows', () => {
    const large = Array.from({ length: 600 }, (_, i) => ({
      slideId: `s${i}`,
      masterId: 'm1',
      layoutId: 'l1',
    }))
    const parsed = parsePowerPointStyleDependencies({ slides: large })
    expect(parsed.slides).toHaveLength(600)
    expect(
      affectedStyleSlideIds(parsed, [
        { op: 'set_master_theme_color', master_id: 'm1', theme_color: 'Accent1', color: '#FFFFFF' },
      ]),
    ).toHaveLength(600)
  })
  it('resolves master/theme references and exact master/layout pairs', () => {
    const snapshot = { slides }
    const theme = {
      op: 'set_master_theme_color',
      master_id: 'm1',
      theme_color: 'Accent1',
      color: '#000000',
    } as const
    expect(affectedStyleSlideIds(snapshot, [theme])).toEqual(['a', 'b'])
    expect(
      affectedStyleSlideIds(snapshot, [
        {
          op: 'set_master_background',
          master_id: 'm1',
          fill: { type: 'solid', color: '#000000', transparency: 0 },
        },
      ]),
    ).toEqual(['a', 'b'])
    const layout = {
      op: 'set_layout_background_following',
      master_id: 'm1',
      layout_id: 'l1',
      follow_master: true,
      show_master_graphics: true,
    } as const
    expect(affectedStyleSlideIds(snapshot, [layout])).toEqual(['a'])
    expect(affectedStyleSlideIds(snapshot, [layout, theme])).toEqual(['a', 'b'])
    expect(affectedStyleSlideIds(snapshot, [{ ...theme, master_id: 'absent' }])).toEqual([])
  })
  it.each([
    undefined,
    {},
    { slides: new Array(1) },
    { slides: [], extra: true },
    { slides: [slides[0], slides[0]] },
    ...['slideId', 'masterId', 'layoutId'].flatMap((key) =>
      [undefined, '', ' ', 'x'.repeat(257), 'a\u0000', 'a\u007f'].map((value) => ({
        slides: [{ ...slides[0], [key]: value }],
      })),
    ),
    { slides: [{ ...slides[0], extra: true }] },
  ])('rejects invalid or incomplete snapshot %#', (value) => {
    expect(() => parsePowerPointStyleDependencies(value)).toThrow('office_read_failed')
  })
})
