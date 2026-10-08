import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'

const root = dirname(fileURLToPath(import.meta.url))
const read = (name) => readFile(join(root, name))
const json = async (name) => JSON.parse(await readFile(join(root, name), 'utf8'))
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')
const assert = (condition, code) => {
  if (!condition) throw new Error(`P0-12 candidate invalid: ${code}`)
}
const [plan, deck, kit, article, checklist, template] = await Promise.all([
  json('reference-plan.json'),
  json('reference-deck.json'),
  json('brand-kit.json'),
  read('deardorff-2020-article.pdf'),
  read('deardorff-2020-checklist.pdf'),
  read('wiswork-benchmark-brand-template.pptx'),
])
assert(
  plan.projectId === deck.id && plan.slides.length === 8 && deck.slides.length === 8,
  'identity',
)
assert(JSON.stringify(plan.brandKit) === JSON.stringify(kit), 'brand_kit')
assert(
  deck.style.fontFace === 'WisWork Benchmark Display 2026' &&
    deck.style.fontFallbacks?.join(',') === 'Noto Sans CJK SC' &&
    JSON.stringify(plan.style) === JSON.stringify(deck.style),
  'font_fallback',
)
assert(
  plan.sources.some((source) => source.snapshotAttachmentId === sha(article)) &&
    plan.sources.some((source) => source.snapshotAttachmentId === sha(checklist)),
  'source_digests',
)
const sources = new Map(plan.sources.map((source) => [source.id, source]))
assert(
  plan.claims.length === deck.claims.length &&
    plan.claims.every((claim, index) => {
      const evidence = claim.sourceIds.map((id) => sources.get(id))
      return (
        claim.reviewStatus === 'needs_review' &&
        evidence.every(Boolean) &&
        deck.claims[index]?.id === claim.id &&
        deck.claims[index]?.text === claim.statement &&
        deck.claims[index]?.source === evidence.map((item) => item.uri).join(' ; ') &&
        deck.claims[index]?.locator === evidence.map((item) => item.locator).join(' ; ')
      )
    }),
  'claim_sources',
)
const components = new Map(kit.layoutComponents.map((item) => [item.id, item]))
assert(
  plan.slides.every((task, index) => {
    const page = deck.slides[index]
    if (task.id !== page?.id || task.title !== page.title) return false
    if (!task.layoutComponentId) return true
    const component = components.get(task.layoutComponentId)
    return (
      component?.layout === task.layout &&
      component.slots.every((slot) => {
        const element = page.elements.find((item) => item.id === slot.id)
        return (
          element?.kind === slot.kind &&
          ['x', 'y', 'w', 'h'].every((key) => element[key] === slot[key])
        )
      })
    )
  }),
  'template_slots',
)
const chart = deck.slides[5].elements.find((element) => element.kind === 'chart')
assert(
  chart?.categories?.join(',') === '前测 n=14,三个月后 n=12' &&
    chart.series[0]?.values?.join(',') === '7,10' &&
    plan.slides[5].chartData?.[0]?.elementId === chart.id,
  'native_chart',
)
const zip = await JSZip.loadAsync(template)
const slideFiles = Object.keys(zip.files).filter((path) =>
  /^ppt\/slides\/slide\d+\.xml$/.test(path),
)
assert(slideFiles.length === 3, 'template_pages')
assert((await zip.file('ppt/slides/slide1.xml').async('string')).includes('研究汇报标题'), 'cover')
assert(
  (await zip.file('ppt/slides/slide2.xml').async('string')).includes('右侧证据或图片区域'),
  'content',
)
console.log(
  'PPT-P0-12 candidate verified: 2 frozen PDFs, Brand Kit slots, font fallback, 8 pages, native chart',
)
