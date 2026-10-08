const { readFile, writeFile } = require('node:fs/promises')
const { createHash } = require('node:crypto')
const { join } = require('node:path')
const prettier = require('prettier')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')
const { checkPresentationChartData } = require('@wiswork/pptx-engine/presentation-chart-data')
const {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
} = require('@wiswork/pptx-engine/presentation-plan')

const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')

async function main() {
  const root = __dirname
  const parent = join(root, '../PPT-P0-01')
  const scenario = JSON.parse(await readFile(join(root, 'scenario.json'), 'utf8'))
  const rightsBytes = await readFile(join(root, 'asset-rights.json'))
  const rights = JSON.parse(rightsBytes.toString('utf8'))
  const article = await readFile(join(root, scenario.documentInputs[0]))
  const checklist = await readFile(join(root, scenario.documentInputs[1]))
  if (scenario.caseId !== 'PPT-P0-16' || scenario.pages.length !== 8)
    throw new Error('P0-16 scenario changed')
  const plan = JSON.parse(await readFile(join(parent, 'reference-plan.json'), 'utf8'))
  const deck = JSON.parse(await readFile(join(parent, 'reference-deck.json'), 'utf8'))
  if (plan.sources.some((source) => source.snapshotAttachmentId !== sha(article)))
    throw new Error('P0-01 source article differs')
  plan.projectId = deck.id = 'p0-16-resume-reference'
  plan.title = deck.title = '计算可复现性研究：断线续跑八页候选稿'
  plan.brief.objective = '基于两份冻结 PDF 与两张授权示意图制作八页研究汇报，并验证同任务断线续跑'
  plan.brief.requiredContent.push('两张自制示意图及权利声明')
  plan.brief.constraints.push('不得将示意图当成研究测量结果')
  const { pdfToPages } = await import('../../../../packages/file-parse/src/pdf.ts')
  const checklistPages = await pdfToPages(checklist)
  const checklistExcerpt = checklistPages
    .flat()
    .find((page) => page.trim())
    ?.trim()
    .slice(0, 400)
  if (!checklistExcerpt) throw new Error('checklist PDF has no extractable text')
  plan.sources.push({
    id: 'supplement-checklist',
    title: 'Deardorff 2020 补充清单 PDF',
    uri: 'doi:10.1371/journal.pone.0230697.s001',
    snapshotAttachmentId: sha(checklist),
    locator: 'PDF 第 1 页',
    excerpt: checklistExcerpt,
  })
  const rightsText = rightsBytes.toString('utf8')
  const rightsExcerpt = rightsText.slice(
    rightsText.indexOf('"assets"'),
    rightsText.indexOf('"assets"') + 370,
  )
  plan.sources.push({
    id: 'illustration-rights',
    title: 'P0-16 自制示意图授权声明',
    uri: 'local:asset-rights.json',
    snapshotAttachmentId: sha(rightsBytes),
    locator: 'assets 1–2',
    excerpt: rightsExcerpt,
  })
  plan.claims.push({
    id: 'checklist-design',
    statement: '六项清单来自论文补充材料，研究表述仍待科研审阅',
    type: 'fact',
    sourceIds: ['supplement-checklist'],
    confidence: 'low',
    reviewStatus: 'needs_review',
  })
  plan.claims.push({
    id: 'illustrations',
    statement: '两张自制示意图只用于说明，不代表研究测量数据',
    type: 'judgment',
    sourceIds: ['illustration-rights'],
    confidence: 'low',
    reviewStatus: 'needs_review',
  })
  deck.claims.push(
    {
      id: 'checklist-design',
      text: plan.claims.at(-2).statement,
      source: plan.sources.at(-2).uri,
      locator: 'PDF 第 1 页',
    },
    {
      id: 'illustrations',
      text: plan.claims.at(-1).statement,
      source: 'local:asset-rights.json',
      locator: 'assets 1–2',
    },
  )
  for (const [index, file] of scenario.imageInputs.entries()) {
    const bytes = await readFile(join(root, file))
    if (rights.assets[index]?.sha256 !== sha(bytes) || !rights.assets[index]?.usePermission)
      throw new Error(`image authorization changed: ${file}`)
    const id = `schematic-${index + 1}`
    deck.assets.push({
      id,
      mime: 'image/png',
      width: 960,
      height: 540,
      base64: bytes.toString('base64'),
      source: rights.assets[index].attribution,
      license: 'owned',
    })
    const page = index === 0 ? 1 : 3
    const slide = deck.slides[page]
    slide.claimIds.push('illustrations')
    plan.slides[page].claimIds.push('illustrations')
    plan.slides[page].requiredAssets.push(id)
    const body = slide.elements.find((element) => element.id === 'body')
    body.h = 2.2
    slide.elements.push({
      kind: 'image',
      id: `${id}-image`,
      assetId: id,
      x: 0.9,
      y: 4.05,
      w: 5.5,
      h: 1.7,
      fit: 'contain',
      altText: `WisWork 自制示意图 ${index + 1}；非研究测量数据`,
    })
  }
  deck.slides[3].claimIds.push('checklist-design')
  plan.slides[3].claimIds.push('checklist-design')
  for (const slide of deck.slides) {
    const footer = slide.elements.find((element) => element.id === 'footer')
    if (footer) footer.y = 6.62
  }
  const parsed = parsePresentationPlan(plan)
  assertDeckMatchesPresentationPlan(deck, parsed)
  const chart = deck.slides[5].elements.find((element) => element.kind === 'chart')
  if (
    checkPresentationChartData(parsed, deck.slides[5].id, [chart]).charts.some(
      (entry) => entry.findings.length,
    )
  )
    throw new Error('chart does not match source-bound plan')
  const compiled = await compilePresentationDeck(deck)
  const format = (value) =>
    prettier.format(JSON.stringify(value), { parser: 'json', printWidth: 100 })
  await writeFile(join(root, 'reference-plan.json'), await format(plan))
  await writeFile(join(root, 'reference-deck.json'), await format(deck))
  await writeFile(join(root, 'wiswork-generated-candidate.pptx'), compiled.bytes)
  const frozen = [
    ...scenario.documentInputs,
    ...scenario.imageInputs,
    'asset-rights.json',
    'scenario.json',
    'reference-plan.json',
    'reference-deck.json',
    'wiswork-generated-candidate.pptx',
  ]
  await writeFile(
    join(root, 'SHA256SUMS'),
    (
      await Promise.all(
        frozen.map(async (file) => `${sha(await readFile(join(root, file)))}  ${file}`),
      )
    ).join('\n') + '\n',
  )
  console.log('P0-16 source-bound eight-page reference generated')
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
