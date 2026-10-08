import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { replacePowerPointShapeGeometryPackage } from '../src/skills/powerpoint/presentation-geometry-revision-package.js'

async function fixture() {
  const deck = benchmarkDeck()
  deck.slides = [{ ...deck.slides[0]!, claimIds: [], elements: [deck.slides[0]!.elements[0]!] }]
  const source = Buffer.from((await compilePresentationDeck(deck)).bytes).toString('base64')
  const zip = await JSZip.loadAsync(source, { base64: true })
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const shape = [...xml.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)].find(([part]) =>
    part.includes('name="title"'),
  )![0]
  const shapeId = /<p:cNvPr\b[^>]*\bid="(\d+)"/.exec(shape)![1]!
  const off = /<a:off x="(\d+)" y="(\d+)"\/>/.exec(shape)!
  const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(shape)!
  const before = {
    left: Number(off[1]) / 12700,
    top: Number(off[2]) / 12700,
    width: Number(ext[1]) / 12700,
    height: Number(ext[2]) / 12700,
  }
  return { source, zip, xml, shape, shapeId, before }
}

it('moves one ordinary native text shape without changing its text or other package parts', async () => {
  const f = await fixture()
  const after = { ...f.before, top: f.before.top + 5.76 }
  const result = await replacePowerPointShapeGeometryPackage(f.source, f.shapeId, f.before, after)
  const revised = await JSZip.loadAsync(result.base64, { base64: true })
  const xml = await revised.file('ppt/slides/slide1.xml')!.async('string')
  expect(xml).toContain(
    `<a:off x="${Math.round(after.left * 12700)}" y="${Math.round(after.top * 12700)}"/>`,
  )
  expect(
    xml.replace(
      `<a:off x="${Math.round(after.left * 12700)}" y="${Math.round(after.top * 12700)}"/>`,
      `<a:off x="${Math.round(f.before.left * 12700)}" y="${Math.round(f.before.top * 12700)}"/>`,
    ),
  ).toBe(f.xml)
  expect(result.beforeDigest).not.toBe(result.afterDigest)
})

it('rejects a stale geometry baseline and unsupported rotation', async () => {
  const f = await fixture()
  await expect(
    replacePowerPointShapeGeometryPackage(
      f.source,
      f.shapeId,
      { ...f.before, top: 0 },
      { ...f.before, top: 5 },
    ),
  ).rejects.toThrow('presentation_baseline_changed')
  const rotated = f.shape.replace(/<a:xfrm\b[^>]*>/, '<a:xfrm rot="60000">')
  expect(rotated).not.toBe(f.shape)
  f.zip.file('ppt/slides/slide1.xml', f.xml.replace(f.shape, rotated))
  await expect(
    replacePowerPointShapeGeometryPackage(
      await f.zip.generateAsync({ type: 'base64' }),
      f.shapeId,
      f.before,
      { ...f.before, top: f.before.top + 5.76 },
    ),
  ).rejects.toThrow('presentation_existing_target_unsupported')
})
