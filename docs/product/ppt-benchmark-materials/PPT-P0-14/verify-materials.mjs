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
const plan = JSON.parse((await read('reference-plan.json')).toString('utf8'))
const deck = JSON.parse((await read('reference-deck.json')).toString('utf8'))
assert(plan.projectId === deck.id && plan.domain === 'science', 'candidate_identity')
assert(plan.slides.length === 8 && deck.slides.length === 8, 'candidate_page_count')
assert(
  plan.slides.every(
    (page, index) => page.id === deck.slides[index].id && page.title === deck.slides[index].title,
  ),
  'candidate_page_binding',
)
const snapshots = new Map([
  [
    'https://science.nasa.gov/earth/earth-observatory/2024-was-the-warmest-year-on-record-153806/',
    await read('nasa-2024-article.html'),
  ],
  ['https://svs.gsfc.nasa.gov/5450', await read('nasa-svs-5450.html')],
  [
    'https://www.nasa.gov/nasa-brand-center/images-and-media/',
    await read('nasa-media-guidelines.html'),
  ],
])
assert(
  plan.sources.length >= 5 &&
    plan.sources.every((source) => {
      const snapshot = snapshots.get(source.uri)
      return (
        snapshot &&
        sha(snapshot) === source.snapshotAttachmentId &&
        snapshot.toString('utf8').includes(source.excerpt)
      )
    }),
  'candidate_source_snapshots',
)
const claims = new Set(plan.claims.map((claim) => claim.id))
const sourceIds = new Set(plan.sources.map((source) => source.id))
const sourceById = new Map(plan.sources.map((source) => [source.id, source]))
assert(
  claims.size === plan.claims.length &&
    deck.claims.length === plan.claims.length &&
    plan.claims.every(
      (claim, index) =>
        claim.reviewStatus === 'needs_review' &&
        claim.sourceIds.length > 0 &&
        claim.sourceIds.every((id) => sourceIds.has(id)) &&
        deck.claims[index]?.id === claim.id &&
        deck.claims[index]?.text === claim.statement &&
        deck.claims[index]?.source ===
          claim.sourceIds.map((id) => sourceById.get(id).uri).join(' ; '),
    ) &&
    plan.slides.every((page) => page.claimIds.every((id) => claims.has(id))),
  'candidate_claim_binding',
)
assert(
  deck.assets.length === 1 &&
    deck.assets[0].id === rights.assetId &&
    deck.assets[0].attachmentId === rights.candidates[1].sha256 &&
    ['p01', 'p04'].every((pageId) =>
      deck.slides
        .find((page) => page.id === pageId)
        ?.elements.some(
          (element) =>
            element.kind === 'image' && element.assetId === rights.assetId && element.altText,
        ),
    ) &&
    deck.slides
      .find((page) => page.id === 'p06')
      ?.elements.some((element) => element.kind === 'table' && element.rows.length === 5),
  'candidate_native_asset_and_table',
)
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
  'PPT-P0-14 materials verified: 2 NASA images, 3 source snapshots, 8-page candidate, native table',
)
