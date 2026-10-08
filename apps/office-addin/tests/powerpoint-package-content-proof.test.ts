import JSZip from 'jszip'
import { expect, it } from 'vitest'
import {
  editPowerPointPackage,
  verifyImportedPowerPointPackage,
  verifyImportedPowerPointPackageContent,
} from '../src/skills/powerpoint/powerpoint-package.js'
async function fixture(media = Uint8Array.from([0xb9, 0x31, 0x56, 0xc4, 0x6b, 0xd9, 0x4d, 0xe1])) {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<Types/>')
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>')
  zip.file('ppt/media/image1.bin', media)
  return zip.generateAsync({ type: 'base64' })
}
async function edited() {
  return editPowerPointPackage(await fixture(), 'slide', [
    {
      path: 'ppt/slides/slide1.xml',
      xml: '<p:sld xmlns:p="urn:p"><p:cSld name="edited"/></p:sld>',
    },
  ])
}
it('compares complete imported entry bytes against the prepared package', async () => {
  const e = await edited()
  expect(await verifyImportedPowerPointPackageContent(e.base64, e)).toBe(true)
})
it('rejects different protected media bytes with the same length and FNV summary', async () => {
  const e = await edited(),
    zip = await JSZip.loadAsync(e.base64, { base64: true })
  zip.file(
    'ppt/media/image1.bin',
    Uint8Array.from([0x5b, 0x9e, 0x9c, 0xd7, 0xfd, 0x87, 0x20, 0x42]),
  )
  const actual = await zip.generateAsync({ type: 'base64' })
  expect(await verifyImportedPowerPointPackage(actual, e)).toBe(true)
  expect(await verifyImportedPowerPointPackageContent(actual, e)).toBe(false)
})
it('retains accepted background normalization while proving every protected entry', async () => {
  const zip = await JSZip.loadAsync(await fixture(), { base64: true })
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    '<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree/></p:cSld></p:sldMaster>',
  )
  const black =
    '<p:sldMaster xmlns:a="urn:a" xmlns:p="urn:p"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></p:bgPr></p:bg><p:spTree/></p:cSld></p:sldMaster>'
  const e = await editPowerPointPackage(await zip.generateAsync({ type: 'base64' }), 'master', [
    { path: 'ppt/slideMasters/slideMaster1.xml', xml: black },
  ])
  const normalized = await JSZip.loadAsync(e.base64, { base64: true })
  normalized.file(
    'ppt/slideMasters/slideMaster1.xml',
    '<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a">\n<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></p:bgPr></p:bg><p:spTree/></p:cSld></p:sldMaster>',
  )
  const actual = await normalized.generateAsync({ type: 'base64' })
  expect(await verifyImportedPowerPointPackage(actual, e)).toBe(true)
  expect(await verifyImportedPowerPointPackageContent(actual, e)).toBe(true)
  normalized.file(
    'ppt/media/image1.bin',
    Uint8Array.from([0x5b, 0x9e, 0x9c, 0xd7, 0xfd, 0x87, 0x20, 0x42]),
  )
  expect(
    await verifyImportedPowerPointPackageContent(
      await normalized.generateAsync({ type: 'base64' }),
      e,
    ),
  ).toBe(false)
})
it('copies expected package and program metadata before asynchronous inspection', async () => {
  const e = await edited(),
    actual = e.base64
  const pending = verifyImportedPowerPointPackageContent(actual, e)
  e.base64 = 'external alias'
  e.afterXml['ppt/slides/slide1.xml'] = '<external/>'
  expect(await pending).toBe(true)
})
