import { expect, it } from 'vitest'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { openPptx } from '@wiswork/pptx-engine'
import { officeOperationsForSlideIR } from '../src/skills/powerpoint/presentation-office-ir'

it('maps shared SlideIR text and shape into native Office point geometry and style', () => {
  const deck = benchmarkDeck()
  expect(officeOperationsForSlideIR(deck.slides[3]!, deck.style, 2, deck.claims)).toMatchObject([
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
      margin: 0,
      verticalAlignment: 'top',
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
    {
      op: 'add_text_box',
      name: 'source-attribution',
      text: '[source-1] 研究报告（合成基准） · 第 1 页',
    },
  ])
})

it('preserves explicit SlideIR text alignment in the Office operation', () => {
  const deck = benchmarkDeck()
  deck.slides[0]!.elements[0] = { ...deck.slides[0]!.elements[0]!, kind: 'text', text: 'Centered', align: 'center' }
  expect(officeOperationsForSlideIR(deck.slides[0]!, deck.style, 0, deck.claims)[0]).toMatchObject({
    op: 'add_text_box', align: 'center', text: 'Centered',
  })
})

it('rejects an unsupported page before emitting a partial native write plan', () => {
  const deck = benchmarkDeck()
  expect(() => officeOperationsForSlideIR(deck.slides[2]!, deck.style, 0, deck.claims)).toThrow(
    'office_api_unsupported',
  )
  const duplicate = {
    ...deck.slides[3]!,
    elements: [deck.slides[3]!.elements[0]!, deck.slides[3]!.elements[0]!],
  }
  expect(() => officeOperationsForSlideIR(duplicate, deck.style, 0, deck.claims)).toThrow('invalid_tool_input')
})

it('maps a shared SlideIR table to a native Office table operation', () => {
  const deck = benchmarkDeck()
  expect(officeOperationsForSlideIR(deck.slides[5]!, deck.style, 0, deck.claims)[1]).toMatchObject({
    op: 'add_native_table', name: 'table', left: 72, top: 180,
    fontFace: 'Microsoft YaHei', fontSize: 16, color: '172033', rows: expect.arrayContaining([expect.arrayContaining(['120'])]),
  })
})

it('keeps supported Office operation structure aligned with the PptxGenJS benchmark output', async () => {
  const deck = benchmarkDeck()
  const { bytes } = await compilePresentationDeck(deck)
  const actual = (await openPptx(bytes)).deck
  for (const pageIndex of [0, 1, 3, 4, 5, 7]) {
    const operations = officeOperationsForSlideIR(deck.slides[pageIndex]!, deck.style, pageIndex, deck.claims)
    const elements = actual.slides[pageIndex]!.elements
    expect(elements.map((element) => element.name).sort()).toEqual(operations.map((op) => op.name).sort())
    for (const operation of operations) {
      const element = elements.find((item) => item.name === operation.name)!
      expect(element).toBeDefined()
      expect(Math.abs((element.transform.offset.x * 72) / 914400 - operation.left)).toBeLessThan(1.5)
      expect(Math.abs((element.transform.offset.y * 72) / 914400 - operation.top)).toBeLessThan(1.5)
      expect(Math.abs((element.transform.offset.cx * 72) / 914400 - operation.width)).toBeLessThan(1.5)
      expect(Math.abs((element.transform.offset.cy * 72) / 914400 - operation.height)).toBeLessThan(1.5)
      if (operation.op === 'add_text_box') {
        expect(element.type).toBe('shape')
        if (element.type === 'shape') {
          expect(element.text?.paragraphs.flatMap((paragraph) => paragraph.runs.map((run) => run.text)).join('')).toBe(operation.text)
          expect(element.text?.paragraphs[0]?.runs[0]?.fontSize).toBe(operation.fontSize)
        }
      } else if (operation.op === 'add_geometric_shape') {
        expect(element.type).toBe('shape')
        if (element.type === 'shape')
          expect(element.presetGeometry).toBe({ rect: 'rect', ellipse: 'ellipse', roundRect: 'roundRect' }[operation.shape])
      } else if (operation.op === 'add_native_table') {
        expect(element.type).toBe('table')
        if (element.type === 'table') {
          expect(element.rows.map((row) => row.map((cell) => cell.text?.paragraphs.flatMap((paragraph) => paragraph.runs.map((run) => run.text)).join('') ?? ''))).toEqual(operation.rows)
          expect(element.rows[0]?.[0]?.text?.paragraphs[0]?.runs[0]?.fontSize).toBe(operation.fontSize)
        }
      }
    }
  }
})
