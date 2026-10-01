import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { pdfToPages } from '../../../../packages/file-parse/src/pdf.ts'

const root = dirname(fileURLToPath(import.meta.url))
const lines = (await readFile(join(root, 'SHA256SUMS'), 'utf8')).trim().split('\n')
assert.equal(lines.length, 15)
const hashes = new Map()
for (const line of lines) {
  const match = /^([a-f0-9]{64}) {2}([A-Za-z0-9./-]+)$/.exec(line)
  assert.ok(match, 'invalid SHA256SUMS row')
  const [, expected, name] = match
  assert.ok(!name.includes('..') && !hashes.has(name), 'invalid or repeated file name')
  const actual = createHash('sha256')
    .update(await readFile(join(root, name)))
    .digest('hex')
  assert.equal(actual, expected, `changed ${name}`)
  hashes.set(name, actual)
}
const rights = JSON.parse(await readFile(join(root, 'asset-rights.json'), 'utf8'))
assert.equal(rights.version, 1)
assert.equal(rights.assets.length, 12)
const imageDigests = new Set()
for (const [index, asset] of rights.assets.entries()) {
  const name = `images/schematic-${String(index + 1).padStart(2, '0')}.png`
  assert.equal(asset.file, name)
  assert.equal(asset.sha256, hashes.get(name))
  assert.ok(asset.usePermission && asset.attribution)
  imageDigests.add(asset.sha256)
  const image = PNG.sync.read(await readFile(join(root, name)))
  assert.deepEqual([image.width, image.height], [960, 540])
}
assert.equal(imageDigests.size, 12)
const source = await readFile(join(root, 'deardorff-2020-article.pdf'))
const original = await readFile(join(root, '../PPT-P0-01/deardorff-2020-article.pdf'))
assert.ok(source.equals(original), 'source PDF differs from original candidate')
const pdfPages = await pdfToPages(source)
assert.equal(pdfPages.length, 11)
assert.match(pdfPages[4], /average score for the pre-workshop checklist was 1\.6 out of 6/)
assert.match(pdfPages[4], /average score increased from 1\.6 to 2\.2/)
assert.match(pdfPages[4], /p = 0\.318/)
assert.match(pdfPages[4], /Use open source software 7 10/)
const zip = await JSZip.loadAsync(
  await readFile(join(root, 'wiswork-image-dense-research-draft.pptx')),
)
const slides = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
const charts = Object.keys(zip.files).filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))
assert.equal(slides.length, 8)
assert.equal(charts.length, 2)
const pictureCounts = []
const expectedImages = [[1, 2], [3, 4], [5, 6], [7, 8], [9], [10], [11, 12], [1, 3]]
const placements = []
for (let page = 1; page <= 8; page++) {
  const xml = await zip.file(`ppt/slides/slide${page}.xml`)?.async('string')
  assert.ok(xml?.includes('<a:t>'), `blank slide ${page}`)
  const pictures = [...xml.matchAll(/<p:pic>([\s\S]*?)<\/p:pic>/g)].map((match) => match[1])
  pictureCounts.push(pictures.length)
  const rels = await zip.file(`ppt/slides/_rels/slide${page}.xml.rels`)?.async('string')
  assert.ok(rels)
  for (const [slot, number] of expectedImages[page - 1].entries()) {
    const picture = pictures[slot]
    assert.ok(picture, `missing page ${page} image slot ${slot}`)
    assert.ok(picture.includes(`descr="自制示意图 ${number}，非研究测量数据"`))
    const id = /<a:blip r:embed="([^"]+)"/.exec(picture)?.[1]
    assert.ok(id)
    const relation = [...rels.matchAll(/<Relationship\b[^>]*\/>/g)]
      .map((match) => match[0])
      .find((value) => value.includes(`Id="${id}"`))
    assert.ok(relation?.includes('/relationships/image"'))
    const target = /Target="([^"]+)"/.exec(relation)?.[1]
    assert.match(target, /^\.\.\/media\/image-\d+-\d+\.png$/)
    const embedded = await zip.file(`ppt/${target.slice(3)}`)?.async('nodebuffer')
    assert.ok(embedded)
    assert.equal(
      createHash('sha256').update(embedded).digest('hex'),
      rights.assets[number - 1].sha256,
    )
    const geometry = /<a:off x="(\d+)" y="(\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(picture)
    assert.deepEqual(geometry?.slice(1), [
      String(slot ? 6181344 : 566928),
      '1417320',
      '5285232',
      '2971800',
    ])
    placements.push({ page, slot, number })
  }
  assert.equal(pictures.length, expectedImages[page - 1].length)
}
assert.deepEqual(pictureCounts, [2, 2, 2, 2, 1, 1, 2, 2])
assert.deepEqual(
  placements.filter(({ number }) => number === 1).map(({ page }) => page),
  [1, 8],
)
assert.deepEqual(
  placements.filter(({ number }) => number === 3).map(({ page }) => page),
  [2, 8],
)
assert.deepEqual(placements.find(({ page, slot }) => page === 4 && slot === 0)?.number, 7)
assert.deepEqual(
  placements.filter(({ number }) => number === 12).map(({ page }) => page),
  [7],
)
for (const [chartIndex, values] of [
  [1, ['1.6', '2.2']],
  [2, ['7', '10']],
]) {
  const chart = await zip.file(`ppt/charts/chart${chartIndex}.xml`)?.async('string')
  assert.ok(chart)
  assert.deepEqual(
    [...chart.matchAll(/<c:v>([^<]+)<\/c:v>/g)].map((match) => match[1]),
    ['论文报告值', '前测', '三个月', ...values],
  )
  const chartRels = await zip.file(`ppt/charts/_rels/chart${chartIndex}.xml.rels`)?.async('string')
  assert.ok(
    chartRels?.includes(
      `Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet${chartIndex}.xlsx"`,
    ),
  )
  assert.match(chart, /<c:externalData r:id="rId1"/)
  const workbook = await zip
    .file(`ppt/embeddings/Microsoft_Excel_Worksheet${chartIndex}.xlsx`)
    ?.async('nodebuffer')
  assert.ok(workbook)
  const sheet = await (
    await JSZip.loadAsync(workbook)
  )
    .file('xl/worksheets/sheet1.xml')
    ?.async('string')
  assert.ok(sheet)
  for (const [row, value] of values.entries())
    assert.ok(sheet.includes(`<c r="B${row + 2}"><v>${value}</v></c>`))
}
process.stdout.write(
  'PPT-P0-13 candidate verified: 8 pages, 12 licensed images, exact target and repeated placements, 2 native chart workbooks.\n',
)
