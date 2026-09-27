import { expect, it } from 'vitest'
import { PNG } from 'pngjs'
import JSZip from 'jszip'
import { execFileSync, spawnSync } from 'node:child_process'
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { openPptx } from '@wiswork/pptx-engine'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { inspectPowerPointPicturePackage } from '../src/skills/powerpoint/powerpoint-package'
import { replacePowerPointPictureMediaPackage } from '../src/skills/powerpoint/presentation-picture-package'

const sofficeAvailable = spawnSync('soffice', ['--version'], { timeout: 5_000 }).status === 0
if (process.env.WISWORK_REQUIRE_LIBREOFFICE === '1' && !sofficeAvailable)
  throw new Error('LibreOffice is required for the picture round-trip gate')

it('replaces only the selected native picture media in an exported one-page package', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const source = Buffer.from((await compilePresentationDeck(deck)).bytes).toString('base64')
  const zip = await JSZip.loadAsync(source, { base64: true })
  const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const picture = [...slideXml.matchAll(/<p:pic\b[^]*?<\/p:pic>/g)].find(([xml]) => xml.includes('name="image"'))![0]
  const shapeId = picture.match(/<p:cNvPr\b[^>]*\bid="(\d+)"/)![1]!
  const original = await inspectPowerPointPicturePackage(source, shapeId)
  const replacement = new PNG({ width: 2, height: 1 })
  replacement.data[0] = 255
  const changed = await replacePowerPointPictureMediaPackage(source, shapeId, {
    mime: 'image/png', base64: PNG.sync.write(replacement).toString('base64'),
  })
  expect(changed.beforeDigest).not.toBe(changed.afterDigest)
  expect(changed.mediaDigest).not.toBe(original.mediaDigest)
  const updated = await inspectPowerPointPicturePackage(changed.base64, shapeId)
  expect(updated.shapeIds).toEqual(original.shapeIds)
  expect(updated.mediaDigest).toBe(changed.mediaDigest)
  const reopened = await openPptx(Buffer.from(changed.base64, 'base64'))
  expect(reopened.deck.slides).toHaveLength(1)
  expect(reopened.deck.slides[0]!.elements.map((item) => item.name)).toEqual(
    (await openPptx(Buffer.from(source, 'base64'))).deck.slides[0]!.elements.map((item) => item.name),
  )
  expect(reopened.deck.slides[0]!.elements.find((item) => item.name === 'image')?.type).toBe('picture')
})

it('rejects a missing picture ID before changing the package', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const source = Buffer.from((await compilePresentationDeck(deck)).bytes).toString('base64')
  const replacement = new PNG({ width: 2, height: 1 })
  replacement.data[0] = 255
  await expect(replacePowerPointPictureMediaPackage(source, '999999', {
    mime: 'image/png', base64: PNG.sync.write(replacement).toString('base64'),
  })).rejects.toThrow('office_api_unsupported')
})

it('does not change a second picture that shares the original image relationship', async () => {
  const deck = benchmarkDeck()
  const slide = deck.slides[2]!
  const image = slide.elements[1]!
  if (image.kind !== 'image') throw new Error('benchmark image missing')
  slide.elements.push({ ...image, id: 'second-image', x: 5 })
  deck.slides = [slide]
  const zip = await JSZip.loadAsync((await compilePresentationDeck(deck)).bytes)
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const pictures = [...xml.matchAll(/<p:pic\b[^]*?<\/p:pic>/g)]
  expect(pictures).toHaveLength(2)
  const firstId = pictures[0]![0].match(/<p:cNvPr\b[^>]*\bid="(\d+)"/)![1]!
  const secondId = pictures[1]![0].match(/<p:cNvPr\b[^>]*\bid="(\d+)"/)![1]!
  const firstEmbed = pictures[0]![0].match(/r:embed="([^"]+)"/)![1]!
  zip.file('ppt/slides/slide1.xml', xml.replace(pictures[1]![0], pictures[1]![0].replace(/r:embed="[^"]+"/, `r:embed="${firstEmbed}"`)))
  const source = await zip.generateAsync({ type: 'base64' })
  const before = await inspectPowerPointPicturePackage(source, secondId)
  const replacement = new PNG({ width: 2, height: 1 })
  replacement.data[0] = 255
  const changed = await replacePowerPointPictureMediaPackage(source, firstId, {
    mime: 'image/png', base64: PNG.sync.write(replacement).toString('base64'),
  })
  expect((await inspectPowerPointPicturePackage(changed.base64, secondId)).mediaDigest).toBe(before.mediaDigest)
  expect((await inspectPowerPointPicturePackage(changed.base64, firstId)).mediaDigest).toBe(changed.mediaDigest)
})

it.skipIf(!sofficeAvailable)('keeps the revised picture and native text through a LibreOffice save and reopen', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const source = Buffer.from((await compilePresentationDeck(deck)).bytes).toString('base64')
  const sourceZip = await JSZip.loadAsync(source, { base64: true })
  const sourceXml = await sourceZip.file('ppt/slides/slide1.xml')!.async('string')
  const shapeId = sourceXml.match(/<p:pic\b[^]*?<p:cNvPr\b[^>]*\bid="(\d+)"/)![1]!
  const replacement = new PNG({ width: 2, height: 1 })
  replacement.data[0] = 255
  const changed = await replacePowerPointPictureMediaPackage(source, shapeId, {
    mime: 'image/png', base64: PNG.sync.write(replacement).toString('base64'),
  })
  const directory = mkdtempSync(join(tmpdir(), 'wiswork-picture-roundtrip-'))
  const inputDirectory = join(directory, 'input')
  const outputDirectory = join(directory, 'output')
  mkdirSync(inputDirectory)
  mkdirSync(outputDirectory)
  try {
    const input = join(inputDirectory, 'revised.pptx')
    writeFileSync(input, Buffer.from(changed.base64, 'base64'))
    execFileSync('soffice', [
      `-env:UserInstallation=file://${join(directory, 'profile')}`,
      '--headless', '--convert-to', 'pptx:Impress MS PowerPoint 2007 XML',
      '--outdir', outputDirectory, input,
    ], { timeout: 60_000, stdio: 'pipe' })
    const reopened = await openPptx(readFileSync(join(outputDirectory, 'revised.pptx')))
    expect(reopened.deck.slides).toHaveLength(1)
    const elements = reopened.deck.slides[0]!.elements
    expect(elements.some((element) => element.type === 'picture')).toBe(true)
    expect(elements.some((element) => element.type !== 'picture')).toBe(true)
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
}, 75_000)
