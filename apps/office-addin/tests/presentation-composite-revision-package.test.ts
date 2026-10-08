import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { openPptx } from '@wiswork/pptx-engine'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { preparePowerPointCompositePagePackage } from '../src/skills/powerpoint/presentation-composite-revision-package.js'
import { inspectPowerPointPicturePackage } from '../src/skills/powerpoint/powerpoint-package.js'

async function fixture() {
  const deck = benchmarkDeck()
  const slide = deck.slides[2]!
  slide.elements.push({ kind: 'text', id: 'caption', x: 1, y: 6, w: 5, h: 0.5, text: '图片说明' })
  deck.slides = [slide]
  const source = Buffer.from((await compilePresentationDeck(deck)).bytes).toString('base64')
  const zip = await JSZip.loadAsync(source, { base64: true })
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const item = (tag: 'sp' | 'pic', name: string) =>
    [...xml.matchAll(new RegExp(`<p:${tag}\\b[^>]*>[\\s\\S]*?<\\/p:${tag}>`, 'g'))].find(([part]) =>
      part.includes(`name="${name}"`),
    )![0]
  const shapeId = (value: string) => /<p:cNvPr\b[^>]*\bid="(\d+)"/.exec(value)![1]!
  const caption = item('sp', 'caption')
  const off = /<a:off x="(\d+)" y="(\d+)"\/>/.exec(caption)!
  const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(caption)!
  const before = {
    left: Number(off[1]) / 12700,
    top: Number(off[2]) / 12700,
    width: Number(ext[1]) / 12700,
    height: Number(ext[2]) / 12700,
  }
  const image = new PNG({ width: 2, height: 1 })
  image.data[0] = 255
  return {
    source,
    titleId: shapeId(item('sp', 'title')),
    captionId: shapeId(caption),
    pictureId: shapeId(item('pic', 'image')),
    before,
    image: PNG.sync.write(image).toString('base64'),
  }
}

it('prepares one editable page revision containing text, geometry, and picture changes', async () => {
  const f = await fixture()
  const beforePicture = await inspectPowerPointPicturePackage(f.source, f.pictureId)
  const revision = await preparePowerPointCompositePagePackage(f.source, {
    text: { shapeId: f.titleId, start: 0, before: '研究图文', after: '研究图表页' },
    geometry: {
      shapeId: f.captionId,
      before: f.before,
      after: { ...f.before, top: f.before.top + 5.76 },
    },
    picture: { shapeId: f.pictureId, image: { mime: 'image/png', base64: f.image } },
  })
  expect(revision.beforeDigest).not.toBe(revision.afterDigest)
  expect(revision.changedRuns).toBe(1)
  expect(revision.mediaDigest).not.toBe(beforePicture.mediaDigest)
  const reopened = await openPptx(Buffer.from(revision.base64, 'base64'))
  expect(reopened.deck.slides).toHaveLength(1)
  const elements = reopened.deck.slides[0]!.elements
  expect(elements.map((element) => element.name)).toEqual(
    (await openPptx(Buffer.from(f.source, 'base64'))).deck.slides[0]!.elements.map(
      (element) => element.name,
    ),
  )
  expect(elements.find((element) => element.name === 'image')?.type).toBe('picture')
  const xml = await (
    await JSZip.loadAsync(revision.base64, { base64: true })
  )
    .file('ppt/slides/slide1.xml')!
    .async('string')
  expect(xml).toContain('<a:t>研究图表页</a:t>')
  expect(xml).toContain(`y="${Math.round((f.before.top + 5.76) * 12700)}"`)
})

it('preserves styled runs during an explicit length-changing composite text edit', async () => {
  const f = await fixture()
  const zip = await JSZip.loadAsync(f.source, { base64: true })
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  zip.file(
    'ppt/slides/slide1.xml',
    xml.replace('<a:t>研究图文</a:t>', '<a:t>研究</a:t></a:r><a:r><a:rPr b="1"/><a:t>图文</a:t>'),
  )
  const source = await zip.generateAsync({ type: 'base64' })
  const revision = await preparePowerPointCompositePagePackage(source, {
    text: {
      shapeId: f.titleId,
      start: 1,
      before: '究图',
      after: '技术路线',
      runReplacements: ['技术', '路线'],
    },
    geometry: {
      shapeId: f.captionId,
      before: f.before,
      after: { ...f.before, top: f.before.top + 5.76 },
    },
    picture: { shapeId: f.pictureId, image: { mime: 'image/png', base64: f.image } },
  })
  expect(revision.changedRuns).toBe(2)
  const revisedXml = await (
    await JSZip.loadAsync(revision.base64, { base64: true })
  )
    .file('ppt/slides/slide1.xml')!
    .async('string')
  expect(revisedXml).toContain('<a:t>研技术</a:t></a:r><a:r><a:rPr b="1"/><a:t>路线文</a:t>')
  expect(revisedXml).toContain(`y="${Math.round((f.before.top + 5.76) * 12700)}"`)
})

it('rejects overlapping target identities before changing a package', async () => {
  const f = await fixture()
  await expect(
    preparePowerPointCompositePagePackage(f.source, {
      text: { shapeId: f.titleId, start: 0, before: '研究图文', after: '研究图表页' },
      geometry: {
        shapeId: f.titleId,
        before: f.before,
        after: { ...f.before, top: f.before.top + 5.76 },
      },
      picture: { shapeId: f.pictureId, image: { mime: 'image/png', base64: f.image } },
    }),
  ).rejects.toThrow('invalid_tool_input')
})
