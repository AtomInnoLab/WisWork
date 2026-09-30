import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { equivalentNativePresentationPage } from '../src/skills/powerpoint/presentation-page-equivalence.js'
import { presentationPackageDigest } from '../src/skills/powerpoint/powerpoint-package.js'

async function page(options: {
  shapeId?: number
  imageId?: string
  mediaName?: string
  media?: number[]
  mediaSize?: number
  x?: number
  chart?: boolean
  title?: string
  theme?: string
}) {
  const imageId = options.imageId ?? 'rId5'
  const mediaName = options.mediaName ?? 'image1.png'
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<Types/>')
  zip.file(
    'ppt/presentation.xml',
    '<p:presentation><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
  )
  zip.file('ppt/theme/theme1.xml', `<theme>${options.theme ?? 'same'}</theme>`)
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="${options.shapeId ?? 2}" name="title"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${options.x ?? 100}" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>${options.title ?? 'Native title'}</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="3" name="picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="${imageId}"/></p:blipFill></p:pic>${options.chart ? '<p:graphicFrame/>' : ''}</p:spTree></p:cSld></p:sld>`,
  )
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="${imageId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/></Relationships>`,
  )
  zip.file(
    `ppt/media/${mediaName}`,
    options.mediaSize
      ? new Uint8Array(options.mediaSize)
      : new Uint8Array(options.media ?? [1, 2, 3]),
  )
  return zip.generateAsync({ type: 'base64' })
}

it('accepts only harmless shape IDs, relationship IDs and media filenames changing', async () => {
  expect(
    await equivalentNativePresentationPage(
      await page({}),
      await page({ shapeId: 20, imageId: 'rId9', mediaName: 'image9.png' }),
    ),
  ).toBe(true)
  expect(
    await equivalentNativePresentationPage(await page({}), await page({ media: [1, 2, 4] })),
  ).toBe(false)
  expect(
    await equivalentNativePresentationPage(await page({}), await page({ mediaName: 'image9.jpg' })),
  ).toBe(false)
  expect(
    await equivalentNativePresentationPage(await page({}), await page({ theme: 'changed' })),
  ).toBe(false)
  expect(await equivalentNativePresentationPage(await page({}), await page({ x: 101 }))).toBe(false)
  expect(
    await equivalentNativePresentationPage(await page({ title: ' ' }), await page({ title: '\t' })),
  ).toBe(false)
  expect(
    await equivalentNativePresentationPage(
      await page({ chart: true }),
      await page({ chart: true, shapeId: 20 }),
    ),
  ).toBe(false)
})

it('compares a prepared picture larger than the edit-package per-entry limit', async () => {
  const source = await page({ mediaSize: 3 * 1024 * 1024 })
  await expect(presentationPackageDigest(source)).rejects.toThrow('invalid_tool_input')
  expect(
    await equivalentNativePresentationPage(
      source,
      await page({ mediaSize: 3 * 1024 * 1024, shapeId: 20 }),
    ),
  ).toBe(true)
})
