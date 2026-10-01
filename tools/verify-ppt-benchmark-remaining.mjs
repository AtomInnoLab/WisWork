import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'

const root = join(fileURLToPath(new URL('../docs/product/ppt-benchmark-materials/', import.meta.url)))
const path = (id, name) => join(root, `PPT-P0-${id}`, name)
const pdf = (id, name, pages, anchor) => {
  const location = path(id, name)
  const info = execFileSync('pdfinfo', [location], { encoding: 'utf8' })
  assert.match(info, new RegExp(`^Pages:\\s+${pages}$`, 'm'), `${id}/${name}: page count`)
  const first = execFileSync('pdftotext', ['-f', '1', '-l', '1', location, '-'], {
    encoding: 'utf8',
  })
  assert(first.replaceAll(/\s+/g, ' ').includes(anchor), `${id}/${name}: source anchor`)
}
const slide = async (zip, number) => {
  const file = zip.file(`ppt/slides/slide${number}.xml`)
  assert(file, `slide ${number} missing`)
  return file.async('string')
}
const slides = (zip) => Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))

pdf('01', 'deardorff-2020-article.pdf', 11, 'Assessing the impact of introductory programming workshops')
pdf('01', 'deardorff-2020-checklist.pdf', 1, 'Reproducibility Score Card')
pdf('02', 'helps-2014-white-noise.pdf', 10, 'Different Effects of Adding White Noise')
pdf('02', 'han-2013-speech-noise.pdf', 7, 'Low-Arousal Speech Noise Improves Performance')
pdf('02', 'mohanathasan-2025-conversation-noise.pdf', 22, 'Exploring short-term memory and listening effort')
pdf('03', 'hess-peterson-2015-article.pdf', 16, 'Bicycles May Use Full Lane')
pdf('03', 'hess-peterson-2015-dictionary.pdf', 2, 'Data Dictionary')
const csv = readFileSync(path('03', 'treatment-outcomes.csv'), 'utf8').trim().split(/\r?\n/)
assert.equal(csv[0], 'treatment,n,permitted_2_agree,safe_2_agree,permitted_4_agree,safe_4_agree')
assert(csv.length > 4, 'P0-03: treatment outcomes missing')
assert(csv.slice(1).every((row) => row.split(',').length === 6), 'P0-03: malformed data row')

pdf('12', 'deardorff-2020-article.pdf', 11, 'Assessing the impact of introductory programming workshops')
pdf('12', 'deardorff-2020-checklist.pdf', 1, 'Reproducibility Score Card')
const kit = JSON.parse(readFileSync(path('12', 'brand-kit.json'), 'utf8'))
assert.equal(kit.revision, 1)
assert(kit.allowedColors.includes('102A43') && kit.allowedColors.includes('007F86'))
assert(kit.layoutComponents.some((item) => item.id === 'benchmark-content'))
const template = await JSZip.loadAsync(readFileSync(path('12', 'wiswork-benchmark-brand-template.pptx')))
assert.equal(slides(template).length, 3)
assert((await slide(template, 1)).includes('研究汇报标题'))
assert((await slide(template, 3)).includes('原生可编辑图表'))

pdf('19', 'deardorff-2020-article.pdf', 11, 'Assessing the impact of introductory programming workshops')
const rights = JSON.parse(readFileSync(path('19', 'asset-rights.json'), 'utf8'))
assert.equal(rights.version, 1)
assert.equal(rights.assets.length, 12)
for (const [index, asset] of rights.assets.entries()) {
  const name = `images/schematic-${String(index + 1).padStart(2, '0')}.png`
  assert.equal(asset.file, name)
  assert.match(asset.usePermission, /Authorized for this benchmark/)
  assert.equal(createHash('sha256').update(readFileSync(path('19', name))).digest('hex'), asset.sha256)
}
const draft = await JSZip.loadAsync(readFileSync(path('19', 'wiswork-image-dense-research-draft.pptx')))
assert.equal(slides(draft).length, 8)
assert.equal(Object.keys(draft.files).filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)).length, 2)
const page4 = await slide(draft, 4)
assert(page4.includes('参与者与证据路径'))
for (const id of ['3', '6', '7']) assert(page4.includes(`id="${id}"`), `P0-19: object ${id}`)
assert.equal((page4.match(/<p:pic\b/g) ?? []).length, 2)

console.log('PPT-P0-01/02/03/12/19: source pages, anchors, data, template and native draft structure verified')
