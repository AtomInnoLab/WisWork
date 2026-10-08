import { readFileSync } from 'node:fs'
import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { openPptx } from '@wiswork/pptx-engine'
import { preparePowerPointCompositePagePackage } from '../src/skills/powerpoint/presentation-composite-revision-package.js'
import { inspectPowerPointPicturePackage } from '../src/skills/powerpoint/powerpoint-package.js'

const materials = new URL(
  '../../../docs/product/ppt-benchmark-materials/PPT-P0-19/',
  import.meta.url,
)

it('revises the exact three native objects in the frozen P0-19 page while preserving other objects', async () => {
  const archive = await JSZip.loadAsync(
    readFileSync(new URL('wiswork-image-dense-research-draft.pptx', materials)),
  )
  for (let page = 1; page <= 8; page++) {
    if (page !== 4) {
      archive.remove(`ppt/slides/slide${page}.xml`)
      archive.remove(`ppt/slides/_rels/slide${page}.xml.rels`)
    }
  }
  const presentation = await archive.file('ppt/presentation.xml')!.async('string')
  archive.file(
    'ppt/presentation.xml',
    presentation.replace(/<p:sldId\b[^>]*\/>/g, (entry) =>
      entry.includes('r:id="rId5"') ? entry : '',
    ),
  )
  const source = await archive.generateAsync({ type: 'base64' })
  const slideBefore = await archive.file('ppt/slides/slide4.xml')!.async('string')
  const caption = [...slideBefore.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)].find(([xml]) =>
    xml.includes('id="6"'),
  )![0]
  const off = /<a:off x="(\d+)" y="(\d+)"\/>/.exec(caption)!
  const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(caption)!
  const before = {
    left: Number(off[1]) / 12700,
    top: Number(off[2]) / 12700,
    width: Number(ext[1]) / 12700,
    height: Number(ext[2]) / 12700,
  }
  const replacement = readFileSync(new URL('images/schematic-12.png', materials)).toString('base64')
  const pictureBefore = await inspectPowerPointPicturePackage(source, '7')
  const revised = await preparePowerPointCompositePagePackage(source, {
    text: { shapeId: '3', start: 0, before: '参与者与证据路径', after: '参与者与证据链' },
    geometry: { shapeId: '6', before, after: { ...before, top: before.top + 5.76 } },
    picture: { shapeId: '7', image: { mime: 'image/png', base64: replacement } },
  })
  const slideAfter = await (
    await JSZip.loadAsync(revised.base64, { base64: true })
  )
    .file('ppt/slides/slide4.xml')!
    .async('string')
  const objects = (xml: string) =>
    [...xml.matchAll(/<p:(sp|pic|graphicFrame|cxnSp)\b[^>]*>[\s\S]*?<\/p:\1>/g)].map(
      ([item]) => item,
    )
  const beforeObjects = objects(slideBefore)
  const afterObjects = objects(slideAfter)
  expect(afterObjects).toHaveLength(beforeObjects.length)
  const id = (xml: string) => /<p:cNvPr\b[^>]*\bid="(\d+)"/.exec(xml)?.[1]
  for (let i = 0; i < beforeObjects.length; i++) {
    expect(id(afterObjects[i]!)).toBe(id(beforeObjects[i]!))
    if (!['3', '6', '7'].includes(id(beforeObjects[i]!)!))
      expect(afterObjects[i]).toBe(beforeObjects[i])
  }
  expect(slideAfter).toContain('<a:t>参与者与证据链</a:t>')
  expect(slideAfter).toContain(`y="${Math.round((before.top + 5.76) * 12700)}"`)
  expect((await inspectPowerPointPicturePackage(revised.base64, '7')).mediaDigest).not.toBe(
    pictureBefore.mediaDigest,
  )
  expect((await openPptx(Buffer.from(revised.base64, 'base64'))).deck.slides).toHaveLength(1)
})
