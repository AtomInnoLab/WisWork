import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import PptxGenJS from 'pptxgenjs'
import { inspectPowerPointPicturePackage } from '../src/skills/powerpoint/powerpoint-package'
const png =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII='
async function fixture() {
  const pptx = new PptxGenJS()
  const slide = pptx.addSlide()
  slide.addImage({ data: `data:image/png;base64,${png}`, x: 1, y: 1, w: 2, h: 2 })
  slide.addText('preserve', { x: 0, y: 0, w: 1, h: 1 })
  const zip = await JSZip.loadAsync(await pptx.write({ outputType: 'uint8array' }))
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const id = /<p:pic>[\s\S]*?<p:cNvPr id="(\d+)"/.exec(xml)![1]!
  return { zip, xml, id, base64: await zip.generateAsync({ type: 'base64' }) }
}
describe('ordinary embedded picture package proof', () => {
  it('inspects a real PptxGenJS PNG picture with SHA256 media and semantic fingerprint', async () => {
    const { base64, id } = await fixture()
    const proof = await inspectPowerPointPicturePackage(base64, id)
    expect(proof.mediaDigest).toMatch(/^[a-f0-9]{64}$/)
    expect(proof.pictureFingerprint).toMatch(/^[a-f0-9]{64}$/)
    expect(proof.shapeIds).toHaveLength(2)
    await expect(inspectPowerPointPicturePackage(base64, '999')).rejects.toThrow(
      'office_api_unsupported',
    )
    await expect(inspectPowerPointPicturePackage(base64, proof.shapeIds[1]!)).rejects.toThrow(
      'office_api_unsupported',
    )
  })
  it.each([
    'crop',
    'effect',
    'linked',
    'animation',
    'duplicate-relation',
    'black-white-mode',
    'rotation-disabled',
  ] as const)('rejects unsupported %s without changing bytes', async (variant) => {
    const { zip, xml, id } = await fixture()
    if (variant === 'crop')
      zip.file(
        'ppt/slides/slide1.xml',
        xml.replace('<a:stretch>', '<a:srcRect l="1000"/><a:stretch>'),
      )
    if (variant === 'effect')
      zip.file(
        'ppt/slides/slide1.xml',
        xml.replace('</p:spPr>', '<a:effectLst><a:grayscl/></a:effectLst></p:spPr>'),
      )
    if (variant === 'linked') zip.file('ppt/slides/slide1.xml', xml.replace('r:embed=', 'r:link='))
    if (variant === 'animation')
      zip.file('ppt/slides/slide1.xml', xml.replace('</p:sld>', '<p:timing/></p:sld>'))
    if (variant === 'black-white-mode')
      zip.file('ppt/slides/slide1.xml', xml.replace('<p:spPr>', '<p:spPr bwMode="gray">'))
    if (variant === 'rotation-disabled')
      zip.file(
        'ppt/slides/slide1.xml',
        xml.replace('<p:blipFill>', '<p:blipFill rotWithShape="0">'),
      )
    if (variant === 'duplicate-relation') {
      const rel = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
      const item = /<Relationship\s[^>]*\/>/.exec(rel)![0]
      zip.file(
        'ppt/slides/_rels/slide1.xml.rels',
        rel.replace('</Relationships>', `${item}</Relationships>`),
      )
    }
    await expect(
      inspectPowerPointPicturePackage(await zip.generateAsync({ type: 'base64' }), id),
    ).rejects.toThrow('office_api_unsupported')
  })
})

it('captures exact original media only after the ordinary-picture proof succeeds', async () => {
  const f = await fixture()
  let original: string | undefined
  const proof = await inspectPowerPointPicturePackage(f.base64, f.id, undefined, (value) => {
    original = value
  })
  expect(original).toBe(png)
  expect(Object.keys(proof).sort()).toEqual(['mediaDigest', 'pictureFingerprint', 'shapeIds'])
  original = undefined
  await expect(
    inspectPowerPointPicturePackage(f.base64, '999', undefined, (value) => {
      original = value
    }),
  ).rejects.toThrow()
  expect(original).toBeUndefined()
})
