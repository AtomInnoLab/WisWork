import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { officeOperationsForSlideIR } from '../src/skills/powerpoint/presentation-office-ir'
import { observePowerPointNativeAdd } from '../src/skills/powerpoint/presentation-native-add-observation'

async function fixture() {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const additions = [
    {
      kind: 'text' as const,
      id: 'added-text',
      x: 1,
      y: 3,
      w: 4,
      h: 0.5,
      text: 'New exact text',
      fontSize: 20,
      bold: true,
    },
    {
      kind: 'shape' as const,
      id: 'added-shape',
      x: 6,
      y: 3,
      w: 2,
      h: 1,
      shape: 'rect' as const,
      fill: '123456',
      lineColor: '654321',
    },
    {
      kind: 'table' as const,
      id: 'added-table',
      x: 1,
      y: 4,
      w: 4,
      h: 1,
      rows: [
        ['A', 'B'],
        ['C', 'D'],
      ],
      fontSize: 16,
    },
  ]
  deck.slides[0]!.elements.push(...additions)
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes),
    path = 'ppt/slides/slide1.xml'
  // Model an Office native TextBox export using the compiled node; preserve all actual compiler IDs.
  const xml = (await zip.file(path)!.async('string')).replace(
    /(<p:cNvPr[^>]*name="added-text"[^]*?<p:cNvSpPr)\/>/,
    '$1 txBox="1"/>',
  )
  const nodes = [...xml.matchAll(/<p:(sp|graphicFrame)\b[^]*?<\/p:\1>/g)].map((match) => match[0])
  const added = additions.map((element) =>
    nodes.find((node) => node.includes(`name="${element.id}"`))!,
  )
  expect(added.every(Boolean)).toBe(true)
  const baseline = added.reduce((value, node) => value.replace(node, ''), xml)
  const operations = officeOperationsForSlideIR(
    { ...deck.slides[0]!, elements: additions, claimIds: [] },
    deck.style,
    0,
  )
  const pack = async (content: string) => {
    zip.file(path, content)
    return zip.generateAsync({ type: 'base64' })
  }
  const before = await pack(baseline)
  const prefix = (count: number) =>
    baseline.replace('</p:spTree>', added.slice(0, count).join('') + '</p:spTree>')
  return { zip, xml, baseline, added, before, operations, pack, prefix }
}

it('observes real native additions as none, a preserved prefix and complete, returning only real package IDs', async () => {
  const f = await fixture()
  for (const count of [0, 1, 2, 3]) {
    const result = await observePowerPointNativeAdd(
      f.before,
      await f.pack(f.prefix(count)),
      f.operations,
    )
    expect(result.status).toBe(count === 0 ? 'none' : count === 3 ? 'complete' : 'prefix_partial')
    expect(result.completedCount).toBe(count)
    expect(result.observed).toEqual(
      f.added.slice(0, count).map((node, operationIndex) => ({
        operationIndex,
        packageShapeId: /<p:cNvPr\b[^>]*id="([^"]+)"/.exec(node)![1],
      })),
    )
  }
})

it('rejects wrong declared fields, gaps, duplicates and extra nodes without claiming a prefix', async () => {
  const f = await fixture()
  for (const xml of [
    f.prefix(3).replace('New exact text', 'Wrong text'),
    f.prefix(3).replace('val="123456"', 'val="999999"'),
    f.baseline.replace('</p:spTree>', f.added[1] + '</p:spTree>'),
    f.prefix(1).replace('</p:spTree>', f.added[0] + '</p:spTree>'),
    f.prefix(3).replace('name="added-shape"', 'name="unexpected"'),
  ]) {
    expect(
      await observePowerPointNativeAdd(f.before, await f.pack(xml), f.operations),
    ).toMatchObject({ status: 'conflict', completedCount: 0, observed: [] })
  }
})

it('requires all old object XML and dependent package parts to remain preserved', async () => {
  const f = await fixture()
  const changedOld = f.prefix(3).replace(/(<a:rPr[^>]*sz=")[0-9]+/, '$19900')
  expect(
    (await observePowerPointNativeAdd(f.before, await f.pack(changedOld), f.operations)).status,
  ).toBe('conflict')
  f.zip.file(
    'ppt/theme/theme1.xml',
    (await f.zip.file('ppt/theme/theme1.xml')!.async('string')).replace('000000', '111111'),
  )
  expect(
    (await observePowerPointNativeAdd(f.before, await f.pack(f.prefix(3)), f.operations)).status,
  ).toBe('conflict')
})

it('fails closed for malformed input and throws cancelled without replay or mutations', async () => {
  const f = await fixture()
  expect((await observePowerPointNativeAdd(f.before, 'not-pptx', f.operations)).status).toBe(
    'conflict',
  )
  expect(
    (await observePowerPointNativeAdd(f.before, f.before, [...f.operations, f.operations[0]!]))
      .status,
  ).toBe('conflict')
  const abort = new AbortController()
  abort.abort()
  await expect(
    observePowerPointNativeAdd(f.before, f.before, f.operations, abort.signal),
  ).rejects.toThrow('cancelled')
})

it('checks every declared text, shape and table field rather than name and text alone', async () => {
  const f = await fixture()
  const mutations = [
    f.added[0]!.replace(/(sz=")[0-9]+/, '$19900'),
    f.added[0]!.replace('b="1"', 'b="0"'),
    f.added[0]!.replace(/(typeface=")[^"]+/, '$1WrongFont'),
    f.added[0]!.replace('anchor="t"', 'anchor="b"'),
    f.added[0]!.replace('lIns="0"', 'lIns="12700"'),
    f.added[0]!.replace('</p:spPr>', '<a:effectLst/></p:spPr>'),
    f.added[1]!.replace('val="654321"', 'val="000000"'),
    f.added[1]!.replace('prst="rect"', 'prst="ellipse"'),
    f.added[2]!.replace(/(<a:gridCol w=")[0-9]+/, '$11'),
    f.added[2]!.replace(/(<a:tr h=")[0-9]+/, '$11'),
    f.added[2]!.replace(
      /(<a:tcPr[^>]*\bmarL=")([0-9]+)/,
      (_m, prefix, size) => prefix + (Number(size) + 12700),
    ),
    f.added[2]!.replace(/(<a:lnL w=")[0-9]+/, '$199999'),
    f.added[2]!.replace(/(sz=")[0-9]+/, '$19900'),
  ]
  for (const [index, node] of mutations.entries()) {
    const original = index < 6 ? f.added[0]! : index < 8 ? f.added[1]! : f.added[2]!
    expect(node, String(index)).not.toBe(original)
    expect(
      (
        await observePowerPointNativeAdd(
          f.before,
          await f.pack(f.prefix(3).replace(original, node)),
          f.operations,
        )
      ).status,
    ).toBe('conflict')
  }
})

it('preserves opaque old object XML and binary parts, accepts XML indentation and ignores ZIP metadata only', async () => {
  const f = await fixture()
  const old = f.baseline.replace(
    '</p:spPr>',
    '<a:effectLst><a:outerShdw blurRad="100"/></a:effectLst></p:spPr>',
  )
  f.zip.file('ppt/embeddings/preserved.bin', new Uint8Array([0, 1, 2, 3]))
  const before = await f.pack(old)
  const complete = old.replace('</p:spTree>', f.added.join('') + '</p:spTree>')
  expect(
    (await observePowerPointNativeAdd(before, await f.pack(complete), f.operations)).status,
  ).toBe('complete')
  expect(
    (
      await observePowerPointNativeAdd(
        before,
        await f.pack(complete.replace(/></g, '>\n<')),
        f.operations,
      )
    ).status,
  ).toBe('complete')
  f.zip.file('ppt/embeddings/preserved.bin', new Uint8Array([0, 1, 2, 4]))
  expect(
    (await observePowerPointNativeAdd(before, await f.pack(complete), f.operations)).status,
  ).toBe('conflict')
})

it('does not certify missing dependency parts or declared operations aimed at different pages', async () => {
  const f = await fixture()
  f.zip.remove('ppt/theme/theme1.xml')
  const broken = await f.pack(f.prefix(3))
  expect((await observePowerPointNativeAdd(broken, broken, f.operations)).status).toBe('conflict')
  const mixed = f.operations.map((operation, index) => ({ ...operation, slide_index: index }))
  expect((await observePowerPointNativeAdd(f.before, f.before, mixed)).status).toBe('conflict')
})

it('rejects cNvPr ID collisions with the slide root group and with old objects', async () => {
  const f = await fixture()
  const rootCollision = f.prefix(1).replace(/(<p:cNvPr id=")[0-9]+(" name="added-text")/, '$11$2')
  expect(
    (await observePowerPointNativeAdd(f.before, await f.pack(rootCollision), f.operations)).status,
  ).toBe('conflict')
  const oldCollision = f.baseline.replace(/(<p:cNvPr id=")[0-9]+(" name="title")/, '$11$2')
  const before = await f.pack(oldCollision)
  expect((await observePowerPointNativeAdd(before, before, f.operations)).status).toBe('conflict')
})

it('checks cancellation during asynchronous package observation without returning an apparent result', async () => {
  const f = await fixture(),
    controller = new AbortController()
  const result = observePowerPointNativeAdd(f.before, f.before, f.operations, controller.signal)
  controller.abort()
  await expect(result).rejects.toThrow('cancelled')
})

it('rejects undeclared nested text formatting and text-box appearance', async () => {
  const f = await fixture()
  for (const node of [
    f.added[0]!.replace(/<a:pPr([^>]*)>/, '<a:pPr$1><a:buChar char="•"/>'),
    f.added[0]!.replace('<a:bodyPr ', '<a:bodyPr vert="vert" '),
    f.added[0]!.replace('<a:rPr ', '<a:rPr i="1" '),
    f.added[0]!.replace('<a:rPr ', '<a:rPr u="sng" '),
    f.added[0]!.replace('<a:lstStyle/>', '<a:lstStyle><a:lvl1pPr/></a:lstStyle>'),
    f.added[0]!.replace('<a:endParaRPr ', '<a:endParaRPr i="1" '),
    f.added[0]!.replace('<a:noFill/>', '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'),
    f.added[0]!.replace(
      '<a:ln></a:ln>',
      '<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
    ),
  ]) {
    expect(node).not.toBe(f.added[0])
    expect(
      (
        await observePowerPointNativeAdd(
          f.before,
          await f.pack(f.prefix(3).replace(f.added[0]!, node)),
          f.operations,
        )
      ).status,
    ).toBe('conflict')
  }
})

it('requires exact native text-box versus geometric-shape roles and rejects undeclared actions and color transforms', async () => {
  const f = await fixture()
  for (const node of [
    f.added[0]!.replace('txBox="1"', ''),
    f.added[0]!.replace('txBox="1"', 'txBox="false"'),
    f.added[0]!.replace('</p:cNvPr>', '<a:hlinkClick r:id="rId1"/></p:cNvPr>'),
    f.added[0]!.replace(
      /<a:srgbClr val="([^"]+)"\/>/,
      '<a:srgbClr val="$1"><a:alpha val="0"/></a:srgbClr>',
    ),
  ]) {
    expect(node).not.toBe(f.added[0])
    expect(
      (
        await observePowerPointNativeAdd(
          f.before,
          await f.pack(f.prefix(3).replace(f.added[0]!, node)),
          f.operations,
        )
      ).status,
    ).toBe('conflict')
  }
  const wrongShape = f.added[1]!.replace('<p:cNvSpPr/>', '<p:cNvSpPr txBox="1"/>')
  expect(wrongShape).not.toBe(f.added[1])
  expect(
    (
      await observePowerPointNativeAdd(
        f.before,
        await f.pack(f.prefix(3).replace(f.added[1]!, wrongShape)),
        f.operations,
      )
    ).status,
  ).toBe('conflict')
})

it('requires the exact requested italic state for a durably added text box', async () => {
  const f = await fixture()
  const operations = structuredClone(f.operations)
  const text = operations[0]!
  if (text.op !== 'add_text_box') throw Error('unexpected fixture')
  text.italic = true
  const node = f.added[0]!.replace(
    /<a:rPr\b([^>]*)>/g,
    (_all, attrs) => `<a:rPr${attrs.replace(/\si="[^"]*"/g, '')} i="1">`,
  )
  const current = f.baseline.replace('</p:spTree>', node + '</p:spTree>')
  expect(
    await observePowerPointNativeAdd(f.before, await f.pack(current), operations),
  ).toMatchObject({ status: 'prefix_partial', completedCount: 1 })
  expect(
    await observePowerPointNativeAdd(f.before, await f.pack(f.prefix(1)), operations),
  ).toMatchObject({ status: 'conflict' })
})

it.each(['\r', '\v', '\r\n', 'wrong text'])(
  'checks equivalent paragraph separators in native-add package proof: %j',
  async (separator) => {
    const f = await fixture()
    const operations = structuredClone(f.operations)
    const text = operations[0]!
    if (text.op !== 'add_text_box') throw Error('unexpected fixture')
    text.text =
      separator === 'wrong text'
        ? 'First\nDifferent\nThird'
        : ['First', 'Second', 'Third'].join(separator)
    const paragraph = f.added[0]!.match(/<a:p>[^]*?<\/a:p>/)![0]
    const paragraphs = ['First', 'Second', 'Third']
      .map((line) => paragraph.replace(/<a:t>[^]*?<\/a:t>/, `<a:t>${line}</a:t>`))
      .join('')
    const node = f.added[0]!.replace(paragraph, paragraphs)
    const current = f.baseline.replace('</p:spTree>', node + '</p:spTree>')
    const proof = await observePowerPointNativeAdd(f.before, await f.pack(current), operations)
    expect(proof).toMatchObject(
      separator === 'wrong text'
        ? { status: 'conflict', completedCount: 0 }
        : { status: 'prefix_partial', completedCount: 1 },
    )
  },
)
