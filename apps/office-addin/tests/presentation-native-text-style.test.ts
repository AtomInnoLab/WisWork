import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { verifyNativeTextStylePackage } from '../src/skills/powerpoint/presentation-native-text-style'

async function packages(change: (xml: string) => string) {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const zip = await JSZip.loadAsync((await compilePresentationDeck(deck)).bytes)
  const path = 'ppt/slides/slide1.xml'
  const xml = await zip.file(path)!.async('string')
  const shape = [...xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)]
    .map((match) => match[0])
    .find((node) => /<a:rPr\b/.test(node))!
  const id = /<p:cNvPr\b[^>]*\bid="(\d+)"/.exec(shape)![1]!
  const before = await zip.generateAsync({ type: 'base64' })
  zip.file(path, xml.replace(shape, change(shape)))
  return { before, after: await zip.generateAsync({ type: 'base64' }), id }
}

it('allows only the requested bold attribute in the saved page package', async () => {
  const p = await packages((xml) =>
    xml.replace(
      /<a:rPr\b([^>]*)>/g,
      (_all, attrs) => `<a:rPr${attrs.replace(/\sb="[^"]*"/g, '')} b="1">`,
    ),
  )
  await expect(
    verifyNativeTextStylePackage(p.before, p.after, p.id, {
      op: 'set_shape_text_style',
      slide_index: 0,
      shape_id: 'sdk',
      bold: true,
    }),
  ).resolves.toBeUndefined()
})

it.each(['text', 'geometry', 'underline'] as const)(
  'rejects unrequested %s changes even when requested style reads back',
  async (kind) => {
    const p = await packages((xml) => {
      if (kind === 'text') return xml.replace(/<a:t>[^<]*<\/a:t>/, '<a:t>tampered</a:t>')
      if (kind === 'geometry') return xml.replace(/<a:off x="\d+"/, '<a:off x="999"')
      return xml.replace(/<a:rPr\b/, '<a:rPr u="sng"')
    })
    await expect(
      verifyNativeTextStylePackage(p.before, p.after, p.id, {
        op: 'set_shape_text_style',
        slide_index: 0,
        shape_id: 'sdk',
        bold: true,
      }),
    ).rejects.toThrow('office_verify_failed')
  },
)

it.each(['ppt/theme/theme1.xml', 'ppt/slides/_rels/slide1.xml.rels'])(
  'rejects shared resource or relationship changes at %s',
  async (path) => {
    const p = await packages((xml) => xml)
    const zip = await JSZip.loadAsync(p.after, { base64: true })
    const before = await zip.file(path)!.async('string')
    zip.file(path, before.replace(/<([^!?][^ >]*)/, '<$1 changed="true"'))
    const after = await zip.generateAsync({ type: 'base64' })
    await expect(
      verifyNativeTextStylePackage(p.before, after, p.id, {
        op: 'set_shape_text_style',
        slide_index: 0,
        shape_id: 'sdk',
        bold: true,
      }),
    ).rejects.toThrow('office_verify_failed')
  },
)

it('rejects a slide background side effect outside the targeted text shape', async () => {
  const p = await packages((xml) => xml)
  const zip = await JSZip.loadAsync(p.after, { base64: true })
  const path = 'ppt/slides/slide1.xml'
  const xml = await zip.file(path)!.async('string')
  zip.file(path, xml.replace('<p:cSld', '<p:cSld changed="true"'))
  await expect(
    verifyNativeTextStylePackage(p.before, await zip.generateAsync({ type: 'base64' }), p.id, {
      op: 'set_shape_text_style',
      slide_index: 0,
      shape_id: 'sdk',
      bold: true,
    }),
  ).rejects.toThrow('office_verify_failed')
})
