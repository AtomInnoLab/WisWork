import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { PNG } from 'pngjs'

const root = dirname(fileURLToPath(import.meta.url))
const read = (name) => readFile(join(root, name))
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')
const assert = (condition, message) => {
  if (!condition) throw Error(message)
}
const rights = JSON.parse((await read('asset-rights.json')).toString('utf8'))
const scenario = JSON.parse((await read('fault-scenario.json')).toString('utf8'))
assert(rights.version === 1 && rights.reviewStatus === 'pending_human', 'rights_status')
assert(rights.candidates.length === 2 && scenario.assetId === rights.assetId, 'candidate_scope')
assert(scenario.caseId === 'PPT-P0-14' && scenario.sourceDocuments.length === 2, 'scenario_scope')
const hashes = new Set()
for (const candidate of rights.candidates) {
  assert(candidate.url.startsWith('https://svs.gsfc.nasa.gov/vis/'), 'source_url')
  const bytes = await read(candidate.localBackup)
  assert(sha(bytes) === candidate.sha256 && !hashes.has(candidate.sha256), 'image_digest')
  hashes.add(candidate.sha256)
  const png = PNG.sync.read(bytes)
  assert(png.width === candidate.width && png.height === candidate.height, 'image_dimensions')
}
const article = (await read('nasa-2024-article.html')).toString('utf8')
const svs = (await read('nasa-svs-5450.html')).toString('utf8')
const guidelines = (await read('nasa-media-guidelines.html')).toString('utf8')
assert(
  article.includes('2024 Was the Warmest Year on Record') && article.includes('1.28'),
  'article',
)
assert(
  svs.includes('Global Temperature Anomalies from 1880 to 2024') &&
    svs.includes('2024GISTEMPMap_2K.png') &&
    svs.includes('1.28'),
  'svs',
)
assert(guidelines.includes('NASA Images and Media Usage Guidelines'), 'guidelines')
const zip = await JSZip.loadAsync(await read('p0-14-reference.pptx'))
const slides = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
assert(slides.length === 8, 'slide_count')
const table = await zip.file('ppt/slides/slide6.xml')?.async('string')
assert(table?.includes('<a:tbl>') && table.includes('一次受控超时'), 'native_failure_table')
const slide3 = await zip.file('ppt/slides/slide3.xml')?.async('string')
assert(slide3?.includes('+1.28'), 'source_number')
const media = Object.keys(zip.files).filter(
  (name) => name.startsWith('ppt/media/') && !zip.files[name].dir,
)
assert(media.length === 2, 'image_insertion')
for (const name of media)
  assert(
    sha(await zip.file(name).async('nodebuffer')) === rights.candidates[1].sha256,
    'inserted_image',
  )
console.log(
  'PPT-P0-14 materials verified: 2 NASA images, 2 source pages, rights guidance, 8 slides, native table',
)
