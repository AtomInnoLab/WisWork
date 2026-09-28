import JSZip from 'jszip'
import PptxGenJS from 'pptxgenjs'
import { expect, it } from 'vitest'
import { inspectPowerPointTextRunPackage } from '../src/skills/powerpoint/presentation-text-run-package.js'

const slide = (runs: string) =>
  `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="7" name="Mixed"/></p:nvSpPr><p:txBody><a:p>${runs}</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`
const run = (value: string, properties = '') =>
  `<a:r><a:rPr ${properties}/><a:t>${value}</a:t></a:r>`
async function packageWith(xml: string) {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', xml)
  return zip.generateAsync({ type: 'base64' })
}

it('binds one text run while allowing only its text payload to change', async () => {
  const original = await inspectPowerPointTextRunPackage(
    await packageWith(slide(run('Before', 'b="1"') + run(' tail', 'i="1"'))),
    '7',
    'Before tail',
    0,
    6,
  )
  const changed = await inspectPowerPointTextRunPackage(
    await packageWith(slide(run('Result', 'b="1"') + run(' tail', 'i="1"'))),
    '7',
    'Result tail',
    0,
    6,
  )
  expect(changed.structureDigest).toBe(original.structureDigest)
  const formatted = await inspectPowerPointTextRunPackage(
    await packageWith(slide(run('Result', 'b="0"') + run(' tail', 'i="1"'))),
    '7',
    'Result tail',
    0,
    6,
  )
  expect(formatted.structureDigest).not.toBe(original.structureDigest)
  const neighbor = await inspectPowerPointTextRunPackage(
    await packageWith(slide(run('Result', 'b="1"') + run(' other', 'i="1"'))),
    '7',
    'Result other',
    0,
    6,
  )
  expect(neighbor.structureDigest).not.toBe(original.structureDigest)
})

it('rejects spans across runs, links, fields, and mismatched host text', async () => {
  const mixed = await packageWith(slide(run('Before', 'b="1"') + run(' tail', 'i="1"')))
  await expect(inspectPowerPointTextRunPackage(mixed, '7', 'Before tail', 4, 4)).rejects.toThrow(
    'presentation_existing_target_unsupported',
  )
  await expect(inspectPowerPointTextRunPackage(mixed, '7', 'Stale text!', 0, 6)).rejects.toThrow(
    'presentation_existing_target_unsupported',
  )
  await expect(inspectPowerPointTextRunPackage(mixed, '8', 'Before tail', 0, 6)).rejects.toThrow(
    'presentation_existing_target_unsupported',
  )
  const linked = await packageWith(
    slide(`<a:r><a:rPr><a:hlinkClick r:id="rId1"/></a:rPr><a:t>Before</a:t></a:r>`),
  )
  await expect(inspectPowerPointTextRunPackage(linked, '7', 'Before', 0, 6)).rejects.toThrow(
    'presentation_existing_target_unsupported',
  )
  const field = await packageWith(slide('<a:fld><a:t>Before</a:t></a:fld>'))
  await expect(inspectPowerPointTextRunPackage(field, '7', 'Before', 0, 6)).rejects.toThrow(
    'presentation_existing_target_unsupported',
  )
})

it('finds a bounded run in a complete PptxGenJS page package', async () => {
  const deck = new PptxGenJS()
  deck.addSlide().addText('before', { x: 1, y: 1, w: 3, h: 1, bold: true })
  const bytes = await deck.write({ outputType: 'nodebuffer' })
  const zip = await JSZip.loadAsync(bytes as Uint8Array)
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const shapeId = /<p:sp>[\s\S]*?<p:cNvPr id="([^"]+)"/.exec(xml)?.[1]
  expect(shapeId).toBeDefined()
  const packageBase64 = await zip.generateAsync({ type: 'base64' })
  expect(
    (await inspectPowerPointTextRunPackage(packageBase64, shapeId!, 'before', 0, 3))
      .structureDigest,
  ).toMatch(/^[a-f0-9]{64}$/)
})
