import { expect, it } from 'vitest'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { officeOperationsForSlideIR } from '../src/skills/powerpoint/presentation-office-ir'

it('maps shared SlideIR text and shape into native Office point geometry and style', () => {
  const deck = benchmarkDeck()
  expect(officeOperationsForSlideIR(deck.slides[3]!, deck.style, 2)).toMatchObject([
    {
      op: 'add_text_box',
      name: 'title',
      left: 72,
      top: 72,
      width: 720,
      height: 72,
      fontFace: 'Microsoft YaHei',
      fontSize: 32,
      color: '172033',
    },
    {
      op: 'add_geometric_shape',
      name: 'step',
      shape: 'roundRect',
      fill: '2255AA',
      lineColor: '2255AA',
      left: 72,
      top: 180,
      width: 216,
      height: 144,
    },
  ])
})

it('rejects an unsupported page before emitting a partial native write plan', () => {
  const deck = benchmarkDeck()
  expect(() => officeOperationsForSlideIR(deck.slides[2]!, deck.style, 0)).toThrow(
    'office_api_unsupported',
  )
  const duplicate = {
    ...deck.slides[3]!,
    elements: [deck.slides[3]!.elements[0]!, deck.slides[3]!.elements[0]!],
  }
  expect(() => officeOperationsForSlideIR(duplicate, deck.style, 0)).toThrow('invalid_tool_input')
})

it('maps a shared SlideIR table to a native Office table operation', () => {
  const deck = benchmarkDeck()
  expect(officeOperationsForSlideIR(deck.slides[5]!, deck.style, 0)[1]).toMatchObject({
    op: 'add_native_table', name: 'table', left: 72, top: 180,
    fontFace: 'Microsoft YaHei', color: '172033', rows: expect.arrayContaining([expect.arrayContaining(['120'])]),
  })
})
