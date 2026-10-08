const { readFile, writeFile, mkdir } = require('node:fs/promises')
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
  const original = join(root, '../PPT-P0-15')
  const plan = JSON.parse(await readFile(join(original, 'reference-plan.json'), 'utf8'))
  const deck = JSON.parse(await readFile(join(original, 'reference-deck.json'), 'utf8'))
  const pagePlan = JSON.parse(await readFile(join(root, 'page-plan.json'), 'utf8'))
  const rightsBytes = await readFile(join(root, 'asset-rights.json'))
  const rightsText = rightsBytes.toString('utf8')
  const rights = JSON.parse(rightsText)
  const sourceBytes = await readFile(join(root, 'deardorff-2020-article.pdf'))
  if (
    pagePlan.pages.length !== 8 ||
    pagePlan.pages[3].id !== 'p04' ||
    !pagePlan.pages[3].revisionTarget
  )
    throw new Error('P0-18 revision scope changed')
  if (
    plan.sources.some(
      (source) =>
        !source.uri.startsWith('local:') && source.snapshotAttachmentId !== sha(sourceBytes),
    )
  )
    throw new Error('P0-18 paper differs from source-bound parent')
  plan.projectId = deck.id = 'p0-18-parent-candidate'
  plan.title = deck.title = '计算可复现性研究：失败页单独重做候选'
  plan.brief.objective = '先生产八页父稿，再只重做第 4 页的原生三步证据路径'
  plan.brief.constraints.push('第 1、2、3、5、6、7、8 页不得重编译或重插入')
  const rightsSource = plan.sources.find((source) => source.id === 'illustration-rights')
  rightsSource.snapshotAttachmentId = sha(rightsBytes)
  rightsSource.uri = 'local:asset-rights.json'
  const index07 = rightsText.indexOf('"assetId": "P0-13-IMG-07"')
  const index08 = rightsText.indexOf('"assetId": "P0-13-IMG-08"')
  if (index07 < 0 || index08 < 0) throw new Error('P0-18 image rights missing')
  rightsSource.locator = 'P0-13-IMG-07'
  rightsSource.excerpt = rightsText.slice(index07, index07 + 340)
  plan.sources.push({
    id: 'illustration-rights-08',
    title: '自制示意图 08 权利声明',
    uri: 'local:asset-rights.json',
    snapshotAttachmentId: sha(rightsBytes),
    locator: 'P0-13-IMG-08',
    excerpt: rightsText.slice(index08, index08 + 340),
  })
  const claim = plan.claims.find((item) => item.id === 'illustration')
  claim.statement = '第 4 页两张自制示意图不是论文测量数据，权利来自本包声明。'
  claim.sourceIds = ['illustration-rights', 'illustration-rights-08']
  const deckClaim = deck.claims.find((item) => item.id === 'illustration')
  deckClaim.text = claim.statement
  deckClaim.source = 'local:asset-rights.json ; local:asset-rights.json'
  deckClaim.locator = 'P0-13-IMG-07 ; P0-13-IMG-08'
  const first = rights.assets.find((asset) => asset.assetId === 'P0-13-IMG-07')
  const second = rights.assets.find((asset) => asset.assetId === 'P0-13-IMG-08')
  for (const item of [first, second]) {
    const bytes = await readFile(join(root, item.file))
    if (sha(bytes) !== item.sha256 || !item.usePermission)
      throw new Error(`P0-18 image unauthorized: ${item.file}`)
  }
  if (sha(Buffer.from(deck.assets[0].base64, 'base64')) !== first.sha256)
    throw new Error('P0-18 first image differs from parent')
  deck.assets.push({
    id: 'second-illustration',
    mime: 'image/png',
    width: 960,
    height: 540,
    base64: (await readFile(join(root, second.file))).toString('base64'),
    source: second.attribution,
    license: 'owned',
  })
  plan.slides[3].requiredAssets = ['participant-illustration', 'second-illustration']
  plan.slides[3].purpose = pagePlan.pages[3].role
  plan.slides[3].acceptanceCriteria = [
    '原生三步流程',
    '两张自制示意图仅作辅助',
    '其他七页不得重编译',
  ]
  const initial = deck.slides[3]
  initial.notes = `预设布局失败：${pagePlan.pages[3].initialQaFinding}`
  initial.elements = initial.elements.filter((element) =>
    ['title', 'body', 'source'].includes(element.id),
  )
  initial.elements.find((element) => element.id === 'body').text =
    '两张自制示意图仅用于说明研究背景，尚未构成明确的可编辑三步证据路径。'
  initial.elements.find((element) => element.id === 'body').h = 0.9
  initial.elements.push(
    {
      kind: 'image',
      id: 'left-image',
      assetId: 'participant-illustration',
      x: 0.85,
      y: 3.0,
      w: 5.5,
      h: 2.3,
      fit: 'contain',
      altText: '自制示意图 07；非研究测量数据',
    },
    {
      kind: 'image',
      id: 'right-image',
      assetId: 'second-illustration',
      x: 6.9,
      y: 3.0,
      w: 5.5,
      h: 2.3,
      fit: 'contain',
      altText: '自制示意图 08；非研究测量数据',
    },
  )
  const parsed = parsePresentationPlan(plan)
  assertDeckMatchesPresentationPlan(deck, parsed)
  for (const index of [4, 5]) {
    const chart = deck.slides[index].elements.find((element) => element.kind === 'chart')
    if (
      !chart ||
      checkPresentationChartData(parsed, deck.slides[index].id, [chart]).charts.some(
        (item) => item.findings.length,
      )
    )
      throw new Error(`P0-18 chart mismatch: ${index + 1}`)
  }
  const revision = structuredClone(deck)
  const revised = revision.slides[3]
  revised.notes = '仅第 4 页修订：原生形状和文字表达招募、工作坊、三个月后访谈；图片为自制示意。'
  revised.elements = revised.elements.filter((element) =>
    ['title', 'body', 'source'].includes(element.id),
  )
  revised.elements.find((element) => element.id === 'body').text =
    '研究证据路径：招募与前测 → 编程工作坊 → 三个月后访谈；示意图不是测量数据。'
  revised.elements.push(
    ...['招募与前测', '编程工作坊', '三个月后访谈'].flatMap((label, index) => {
      const x = 0.9 + index * 4.1
      return [
        {
          kind: 'shape',
          id: `step-${index + 1}-shape`,
          shape: 'roundRect',
          x,
          y: 2.75,
          w: 3.4,
          h: 0.9,
        },
        {
          kind: 'text',
          id: `step-${index + 1}-label`,
          x: x + 0.16,
          y: 3.02,
          w: 3.05,
          h: 0.35,
          text: label,
          fontSize: 18,
          color: 'FFFFFF',
        },
      ]
    }),
    {
      kind: 'image',
      id: 'left-image',
      assetId: 'participant-illustration',
      x: 0.9,
      y: 4.35,
      w: 5.2,
      h: 1.35,
      fit: 'contain',
      altText: '自制示意图 07；非研究测量数据',
    },
    {
      kind: 'image',
      id: 'right-image',
      assetId: 'second-illustration',
      x: 7.0,
      y: 4.35,
      w: 5.2,
      h: 1.35,
      fit: 'contain',
      altText: '自制示意图 08；非研究测量数据',
    },
    {
      kind: 'text',
      id: 'image-caption',
      x: 0.9,
      y: 5.82,
      w: 11.4,
      h: 0.35,
      text: '两图均为 WisWork 自制示意，不是研究测量结果',
      fontSize: 11,
    },
  )
  assertDeckMatchesPresentationPlan(revision, parsed)
  const parentPptx = await compilePresentationDeck(deck)
  const revisedPagePptx = await compilePresentationDeck({ ...revision, slides: [revised] })
  const parentPages = await Promise.all(
    deck.slides.map((slide) => compilePresentationDeck({ ...deck, slides: [slide] })),
  )
  // A derived task keeps the seven parent page packages verbatim. It compiles only p04.
  const revisedPages = parentPages.map((page, index) => (index === 3 ? revisedPagePptx : page))
  for (const index of [0, 1, 2, 4, 5, 6, 7])
    if (sha(parentPages[index].bytes) !== sha(revisedPages[index].bytes))
      throw new Error(`P0-18 preserved page ${index + 1} recompiled differently`)
  const artifactManifest = {
    version: 1,
    projectId: plan.projectId,
    parentPageDigests: parentPages.map((page) => sha(page.bytes)),
    revisedPageDigests: revisedPages.map((page) => sha(page.bytes)),
    preservedPageIds: pagePlan.preserveUnchangedPageIds,
    revisedPageId: 'p04',
  }
  const format = (value) =>
    prettier.format(JSON.stringify(value), { parser: 'json', printWidth: 100 })
  await writeFile(join(root, 'reference-plan.json'), await format(plan))
  await writeFile(join(root, 'reference-deck.json'), await format(deck))
  await writeFile(
    join(root, 'revised-page-deck.json'),
    await format({ ...revision, slides: [revised] }),
  )
  await writeFile(join(root, 'page-artifact-manifest.json'), await format(artifactManifest))
  await writeFile(join(root, 'wiswork-parent-candidate.pptx'), parentPptx.bytes)
  await writeFile(join(root, 'revised-page-candidate.pptx'), revisedPagePptx.bytes)
  await mkdir(join(root, 'parent-pages'), { recursive: true })
  for (const [index, page] of parentPages.entries())
    await writeFile(
      join(root, `parent-pages/p${String(index + 1).padStart(2, '0')}.pptx`),
      page.bytes,
    )
  const names = [
    'deardorff-2020-article.pdf',
    'wiswork-image-dense-research-draft.pptx',
    'page-plan.json',
    'asset-rights.json',
    ...rights.assets.map((asset) => asset.file),
    'reference-plan.json',
    'reference-deck.json',
    'revised-page-deck.json',
    'page-artifact-manifest.json',
    'wiswork-parent-candidate.pptx',
    'revised-page-candidate.pptx',
    ...parentPages.map((_, index) => `parent-pages/p${String(index + 1).padStart(2, '0')}.pptx`),
  ]
  await writeFile(
    join(root, 'SHA256SUMS'),
    (
      await Promise.all(
        names.map(async (name) => `${sha(await readFile(join(root, name)))}  ${name}`),
      )
    ).join('\n') + '\n',
  )
  console.log('P0-18 eight-page parent and isolated revised page generated')
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
