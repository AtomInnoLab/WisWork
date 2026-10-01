const { readFile, writeFile } = require('node:fs/promises')
const { createHash } = require('node:crypto')
const { join } = require('node:path')
const prettier = require('prettier')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')
const { checkPresentationChartData } = require('@wiswork/pptx-engine/presentation-chart-data')
const {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
  presentationPlanClaims,
} = require('@wiswork/pptx-engine/presentation-plan')

const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')

async function main() {
  const root = __dirname
  const scenario = JSON.parse(await readFile(join(root, 'scenario.json'), 'utf8'))
  const illustration = (await readFile(join(root, 'images/schematic-07.png'))).toString('base64')
  const titles = scenario.boundGeneration.expectedTitles
  const bodies = [
    'Deardorff 等，PLOS ONE 2020。此稿是现稿修改路径对照，研究主张仍待科研审阅。',
    '研究关注编程工作坊与生物医学工作流程的计算可重复性；观察变化不等于因果证明。',
    '工作坊前访谈 14 人，三个月后有 12 人完成后访谈；两时点分母不同。',
    '先招募并访谈，再开展工作坊，三个月后进行后访谈。自制示意图不代表研究测量数据。',
    '六项清单平均分从 1.6/6 到 2.2/6；论文报告 p=0.318，未达统计显著。',
    'Table 1：开源软件使用计数前测 7/14，后测 10/12，不能将计数差视作同一人的逐个变化。',
    '小样本、招募或应答偏差、单人编码及统计功效不足限制了推断。',
    '结论和局限并列；科研审阅与真实 PowerPoint 验收尚未完成。',
  ]
  const paperBytes = await readFile(join(root, scenario.source))
  const rightsBytes = await readFile(join(root, 'asset-rights.json'))
  const imageBytes = await readFile(join(root, 'images/schematic-07.png'))
  const rights = JSON.parse(rightsBytes.toString('utf8'))
  const imageRight = rights.assets.find((asset) => asset.file === 'images/schematic-07.png')
  if (!imageRight || imageRight.sha256 !== sha(imageBytes) || scenario.caseId !== 'PPT-P0-15')
    throw new Error('P0-15 frozen rights or scope changed')
  const { pdfToPages } = await import('../../../../packages/file-parse/src/pdf.ts')
  const paperPages = await pdfToPages(paperBytes)
  if (!paperPages[0].includes('Published: July 8, 2020'))
    throw new Error('paper publication date changed')
  const page5 = paperPages[4]
  const page9 = paperPages[8]
  const excerpt = (page, start, end) => {
    const from = page.indexOf(start)
    const to = page.indexOf(end, from)
    if (from < 0 || to < 0) throw new Error(`paper source missing: ${start}`)
    return page.slice(from, to + end.length)
  }
  const rightsText = rightsBytes.toString('utf8')
  const rightsOffset = rightsText.indexOf('"assetId": "P0-13-IMG-07"')
  if (rightsOffset < 0) throw new Error('schematic rights missing')
  const paperUri = 'https://doi.org/10.1371/journal.pone.0230697'
  const sources = [
    {
      id: 'paper-method',
      title: '原论文摘要的访谈方法与起始样本',
      uri: paperUri,
      snapshotAttachmentId: sha(paperBytes),
      locator: 'PDF 1 methods',
      excerpt: excerpt(
        paperPages[0],
        'This mixed methods study consisted',
        'introductory programming workshop.',
      ),
      asOf: '2020-07-08',
    },
    {
      id: 'paper-sample',
      title: '原论文样本与三个月后访谈',
      uri: paperUri,
      snapshotAttachmentId: sha(paperBytes),
      locator: 'PDF 5 sample',
      excerpt: excerpt(page5, '13 of 14', 'participated in the post-workshop interviews.'),
      asOf: '2020-07-08',
    },
    {
      id: 'paper-score',
      title: '原论文清单均值与显著性',
      uri: paperUri,
      snapshotAttachmentId: sha(paperBytes),
      locator: 'PDF 5 checklist',
      excerpt: excerpt(
        page5,
        'The average score for the pre-workshop checklist was 1.6',
        'p = 0.318',
      ),
      asOf: '2020-07-08',
    },
    {
      id: 'paper-table',
      title: '原论文 Table 1 开源软件使用人数',
      uri: paperUri,
      snapshotAttachmentId: sha(paperBytes),
      locator: 'PDF 5 Table 1',
      excerpt: page5.split('\n').find((line) => line.includes('Use open source software 7 10')),
      asOf: '2020-07-08',
    },
    {
      id: 'paper-limitations',
      title: '原论文局限段落',
      uri: paperUri,
      snapshotAttachmentId: sha(paperBytes),
      locator: 'PDF 9 limitations',
      excerpt: excerpt(
        page9,
        'The analysis and coding for this project',
        'quantitative analysis lacked appropriate power.',
      ),
      asOf: '2020-07-08',
    },
    {
      id: 'illustration-rights',
      title: '自制示意图 07 权利声明',
      uri: 'local:asset-rights.json',
      snapshotAttachmentId: sha(rightsBytes),
      locator: 'P0-13-IMG-07',
      excerpt: rightsText.slice(rightsOffset, rightsOffset + 320),
      asOf: '2026-09-28',
    },
  ]
  if (sources.some((source) => !source.excerpt)) throw new Error('empty P0-15 source')
  const professionalContext = (limitations) => ({
    domain: 'science',
    materialKind: 'paper',
    publicationId: '10.1371/journal.pone.0230697',
    version: '2020 PLOS ONE article',
    sample: '14 pre-workshop, 12 post-workshop interviews',
    method: 'qualitative interviews and six-item checklist',
    statisticalBasis: 'reported p=0.318; paired change not statistically significant',
    limitations,
  })
  const claim = (id, statement, sourceIds, limitations, type = 'fact') => ({
    id,
    statement,
    sourceIds,
    type,
    confidence: 'low',
    reviewStatus: 'needs_review',
    professionalContext: professionalContext(limitations),
  })
  const claims = [
    claim(
      'sample',
      '工作坊前访谈 14 人，三个月后完成访谈 12 人；两时点分母不同。',
      ['paper-method', 'paper-sample'],
      '小样本与失访，不可作因果解释',
    ),
    claim(
      'score',
      '六项清单均值从 1.6/6 到 2.2/6，p=0.318，未达统计显著。',
      ['paper-score'],
      '不可宣称工作坊导致改善',
    ),
    claim(
      'open-source',
      '开源软件使用人数前测 7/14、后测 10/12；分母不同。',
      ['paper-table', 'paper-sample'],
      '不能把计数差写成相同个人的逐一变化',
    ),
    claim(
      'limitations',
      '小样本、招募与应答偏差、单人编码和统计功效不足限制推断。',
      ['paper-limitations'],
      '论文局限需科研审阅',
    ),
    claim(
      'illustration',
      '第 4 页自制示意图不是论文测量数据，权利来自本包声明。',
      ['illustration-rights'],
      '图示内容和版权声明待人工复核',
      'judgment',
    ),
  ]
  const style = {
    fontFace: 'Noto Sans CJK SC',
    background: 'FFFFFF',
    textColor: '122B43',
    accentColor: '008086',
  }
  const projectId = 'p0-15-generated-candidate'
  const claimIds = [
    [],
    [],
    ['sample'],
    ['illustration'],
    ['score'],
    ['open-source'],
    ['limitations'],
    ['limitations'],
  ]
  const sections = [
    'research_question',
    'research_question',
    'methods_and_sample',
    'methods_and_sample',
    'results_and_data',
    'results_and_data',
    'scope_and_limitations',
    'research_references',
  ]
  const chartPoints = (sourceId, values, claimId) => {
    const source = sources.find((item) => item.id === sourceId)
    return values.map((value) => {
      const text = String(value)
      const offset = source.excerpt.indexOf(text)
      if (offset < 0) throw new Error(`numeric paper evidence missing: ${sourceId} ${value}`)
      return {
        value,
        claimId,
        basis: { kind: 'source', sourceId, excerptOffset: offset, excerptText: text },
      }
    })
  }
  const plan = {
    version: 1,
    projectId,
    title: '计算可复现性工作坊：同内容生成绑定稿候选',
    domain: 'science',
    brief: {
      objective: '从冻结论文和授权示意图生成八页原生稿，供与任意现稿修改路径对照',
      audience: '科研专业审阅者',
      language: 'zh-CN',
      minutes: 8,
      requiredContent: [
        '14/12 不同分母',
        '1.6/6 到 2.2/6 且不显著',
        'Table 1 原生图表',
        '自制示意图权利',
      ],
      constraints: [
        '不得推断因果效果',
        '示意图不是研究测量数据',
        '生成后的真实 PowerPoint 导入回执待验',
      ],
    },
    sources,
    claims,
    style,
    slides: titles.map((title, index) => ({
      id: `p${String(index + 1).padStart(2, '0')}`,
      title,
      purpose: '保留论文事实与自制示意图的来源及科研解释边界',
      claimIds: claimIds[index],
      domainSection: sections[index],
      layout:
        index === 0
          ? 'cover'
          : index === 7
            ? 'summary'
            : index === 4 || index === 5
              ? 'chart'
              : 'content',
      requiredAssets: index === 3 ? ['participant-illustration'] : [],
      acceptanceCriteria: ['八页原生可编辑', '第 4 页两对象可单独修改', '图表数据与论文一致'],
      ...(index === 4
        ? {
            chartData: [
              {
                elementId: 'checklist-chart',
                categories: ['前测', '三个月后'],
                series: [
                  { name: '六项清单均值', points: chartPoints('paper-score', [1.6, 2.2], 'score') },
                ],
              },
            ],
          }
        : {}),
      ...(index === 5
        ? {
            chartData: [
              {
                elementId: 'software-chart',
                categories: ['前测 n=14', '后测 n=12'],
                series: [
                  {
                    name: '开源软件使用人数',
                    points: chartPoints('paper-table', [7, 10], 'open-source'),
                  },
                ],
              },
            ],
          }
        : {}),
    })),
  }
  const parsedPlan = parsePresentationPlan(plan)
  const deck = {
    version: 1,
    id: projectId,
    title: plan.title,
    style,
    assets: [
      {
        id: 'participant-illustration',
        mime: 'image/png',
        width: 960,
        height: 540,
        base64: illustration,
        source: 'WisWork 自制示意图 07',
        license: 'owned',
      },
    ],
    claims: presentationPlanClaims(parsedPlan),
    slides: titles.map((title, index) => ({
      id: `p${String(index + 1).padStart(2, '0')}`,
      title,
      claimIds: plan.slides[index].claimIds,
      notes: '来源：Deardorff et al., PLOS ONE 2020；本稿为待审候选。',
      elements: [
        { kind: 'text', id: 'title', x: 0.62, y: 0.45, w: 12, h: 0.75, text: title, fontSize: 26 },
        {
          kind: 'text',
          id: 'body',
          x: 0.85,
          y: 1.8,
          w: 11.4,
          h: 2.0,
          text: bodies[index],
          fontSize: 19,
        },
        {
          kind: 'text',
          id: 'source',
          role: 'decoration',
          x: 0.85,
          y: 6.62,
          w: 11.4,
          h: 0.25,
          text: '来源：Deardorff et al., PLOS ONE 2020；待科研审阅',
          fontSize: 9,
        },
        ...(index === 3
          ? [
              {
                kind: 'image',
                id: 'participant-image',
                assetId: 'participant-illustration',
                x: 0.85,
                y: 4.05,
                w: 5.5,
                h: 1.7,
                fit: 'contain',
                altText: '自制示意图 07：招募、工作坊与三个月后访谈；非研究测量数据',
              },
              {
                kind: 'text',
                id: 'left-caption',
                x: 0.85,
                y: 5.9,
                w: 5.5,
                h: 0.35,
                text: '招募、工作坊、三个月后访谈',
                fontSize: 12,
              },
            ]
          : []),
        ...(index === 4
          ? [
              {
                kind: 'chart',
                id: 'checklist-chart',
                x: 1.0,
                y: 4.0,
                w: 10,
                h: 2.5,
                chartType: 'bar',
                categories: ['前测', '三个月后'],
                series: [{ name: '六项清单均值', values: [1.6, 2.2] }],
              },
            ]
          : []),
        ...(index === 5
          ? [
              {
                kind: 'chart',
                id: 'software-chart',
                x: 1.0,
                y: 4.0,
                w: 10,
                h: 2.5,
                chartType: 'bar',
                categories: ['前测 n=14', '后测 n=12'],
                series: [{ name: '开源软件使用人数', values: [7, 10] }],
              },
            ]
          : []),
      ],
    })),
  }
  assertDeckMatchesPresentationPlan(deck, parsedPlan)
  for (const index of [4, 5]) {
    const chart = deck.slides[index].elements.find((element) => element.kind === 'chart')
    const report = checkPresentationChartData(parsedPlan, deck.slides[index].id, [chart])
    if (report.charts.some((item) => item.findings.length))
      throw new Error(`P0-15 chart data mismatch: ${index}`)
  }
  const compiled = await compilePresentationDeck(deck)
  for (const [name, value] of [
    ['reference-plan.json', plan],
    ['reference-deck.json', deck],
  ]) {
    const path = join(root, name)
    await writeFile(
      path,
      await prettier.format(JSON.stringify(value), {
        ...(await prettier.resolveConfig(path)),
        parser: 'json',
      }),
    )
  }
  await writeFile(join(root, 'wiswork-generated-candidate.pptx'), compiled.bytes)
  console.log(
    `PPT-P0-15 generated candidate: ${compiled.bytes.length} bytes, ${deck.slides.length} pages`,
  )
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
