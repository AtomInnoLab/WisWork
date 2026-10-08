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
const cases = {
  science: {
    audience: '科研专业审阅者',
    boundary: '小样本观察，不作因果推断；统计表述待科研审阅',
    bodies: [
      'Deardorff 2020，PLOS ONE。科研汇报候选稿；本项目资料只取本论文。',
      '问题、方法、清单结果、Table 1 与局限依次呈现；图表保留不同分母。',
      '论文使用工作坊前后访谈观察计算工作流可复现性。',
      '前测 14 人，三个月后访谈 12 人；两时点分母不同。',
      '六项清单均分 1.6/6 到 2.2/6，p=0.318；不能称为统计显著改善。',
      'Table 1 开源软件使用人数前测 7/14、后测 10/12。',
      '小样本及分析者单人编码限制推断范围。',
      '研究主张待科研审阅；原生对象和真实 PowerPoint 保存重开待验。',
    ],
    terms: [
      'This mixed methods study consisted',
      'only 12 researchers',
      '1.6 to 2.2',
      'Use open source software 7 10',
      'small sample size',
    ],
    chart: { categories: ['前测 n=14', '后测 n=12'], values: [7, 10], name: '使用开源软件的人数' },
  },
  legal: {
    audience: '法律专业审阅者',
    boundary: '只说明 EDPB 指南文本及版本，不提供个案法律意见',
    bodies: [
      'EDPB Guidelines 07/2020 Version 2.1；本稿是指南说明候选稿。',
      '依次说明版本、控制者与处理者概念、角色判断及待审边界。',
      '封面载明 Version 2.1，并区分 2021 年通过与后续版本修订。',
      '指南将控制者关联于处理目的和方式的决定；具体关系需结合事实。',
      '先识别处理目的，再核对处理方式与实际决定权。',
      '处理者义务涉及 Article 28 安排；本页只摘要指南，不代替法规原文。',
      '共同控制与处理者关系依具体活动判断，不可按名称自动认定。',
      '本稿待法律审阅；不作个案法律意见或现行效力结论。',
    ],
    terms: [
      'Version 2.1',
      'A controller determines the purposes and means',
      'purposes and means',
      'Article 28',
      'joint controllers',
    ],
  },
  finance: {
    audience: '财务专业审阅者',
    boundary: '只报告 FY2024 10-K 历史披露值，不虚构预测或投资意见',
    bodies: [
      'Apple FY2024 Form 10-K；本稿只使用该份冻结审计年报。',
      '依次呈现报告期、审计身份、收入、利润与现金流、年度对照及口径。',
      '财政年度截至 2024-09-28；与上一财年的日期不同。',
      '合并净销售额 FY2024 为 391,035 百万美元，FY2023 为 383,285 百万美元。',
      '净利润与现金流须分别按审计年报原表及单位核对，不将利润等同现金。',
      '原生图表并列展示两年净销售额原值；单位为百万美元。',
      '比较只覆盖历史披露；重述、分类口径和风险仍待财务审阅。',
      '本稿待财务审阅；不含预测或投资建议。',
    ],
    terms: [
      'For the fiscal year ended September 28, 2024',
      'We have audited',
      'Total net sales 391,035 383,285',
      'Net income 93,736',
      'Risk Factors',
    ],
    chart: {
      categories: ['FY2023', 'FY2024'],
      values: [383285, 391035],
      name: '净销售额（百万美元）',
    },
  },
}

async function main() {
  const root = __dirname
  const scenario = JSON.parse(await readFile(join(root, 'scenario.json'), 'utf8'))
  const { pdfToPages } = await import('../../../../packages/file-parse/src/pdf.ts')
  const files = []
  for (const document of scenario.documents) {
    const spec = cases[document.key]
    if (!spec || document.slideTitles.length !== 8 || spec.bodies.length !== 8)
      throw new Error('P0-17 case mismatch')
    const bytes = await readFile(join(root, document.source))
    const pages = await pdfToPages(bytes)
    if (pages.length !== document.pages) throw new Error(`${document.key} page count changed`)
    const sources = spec.terms.map((term, index) => {
      const pageIndex = pages.findIndex((page) => page.toLowerCase().includes(term.toLowerCase()))
      if (pageIndex < 0) throw new Error(`${document.key} missing source: ${term}`)
      const page = pages[pageIndex]
      const offset = page.toLowerCase().indexOf(term.toLowerCase())
      return {
        id: `${document.key}-source-${index + 1}`,
        title: `${document.title}，原件第 ${pageIndex + 1} 页`,
        uri: `local:${document.source}`,
        snapshotAttachmentId: sha(bytes),
        locator: `PDF 第 ${pageIndex + 1} 页`,
        excerpt: page.slice(
          Math.max(0, offset - 30),
          Math.min(page.length, offset + Math.max(150, term.length + 80)),
        ),
      }
    })
    const claims = sources.map((source, index) => ({
      id: `${document.key}-claim-${index + 1}`,
      statement: spec.bodies[index + 2],
      type: 'fact',
      sourceIds: [source.id],
      confidence: 'low',
      reviewStatus: 'needs_review',
    }))
    const style = {
      fontFace: 'Noto Sans CJK SC',
      background: 'FFFFFF',
      textColor: '173248',
      accentColor: '087D83',
    }
    const projectId = `p0-17-${document.key}`
    const plan = {
      version: 1,
      projectId,
      title: document.title,
      brief: {
        objective: `为 ${document.marker} 独立生成八页，不引用其他两个项目来源`,
        audience: spec.audience,
        language: 'zh-CN',
        minutes: 8,
        requiredContent: document.slideTitles,
        constraints: [spec.boundary, '真实 PowerPoint 三窗口与另存验收待完成'],
      },
      sources,
      claims,
      style,
      slides: document.slideTitles.map((title, index) => ({
        id: `p${String(index + 1).padStart(2, '0')}`,
        title,
        purpose: `仅呈现 ${document.marker} 项目来源及其专业边界`,
        claimIds: index >= 2 && index <= 6 ? [claims[index - 2].id] : [],
        layout:
          index === 0
            ? 'cover'
            : index === 7
              ? 'summary'
              : spec.chart && index === 5
                ? 'chart'
                : 'content',
        requiredAssets: [],
        acceptanceCriteria: ['原生可编辑文字', '来源只来自本项目', '保留项目识别码'],
      })),
    }
    const deck = {
      version: 1,
      id: projectId,
      title: document.title,
      style,
      assets: [],
      claims: claims.map((claim, index) => ({
        id: claim.id,
        text: claim.statement,
        source: sources[index].uri,
        locator: sources[index].locator,
      })),
      slides: document.slideTitles.map((title, index) => ({
        id: `p${String(index + 1).padStart(2, '0')}`,
        title,
        claimIds: index >= 2 && index <= 6 ? [claims[index - 2].id] : [],
        notes: `${document.marker}；${spec.boundary}。`,
        elements: [
          { kind: 'text', id: 'title', x: 0.65, y: 0.48, w: 12, h: 0.7, text: title, fontSize: 26 },
          {
            kind: 'text',
            id: 'body',
            x: 0.9,
            y: 1.65,
            w: 11.4,
            h: spec.chart && index === 5 ? 0.9 : 3.8,
            text: spec.bodies[index],
            fontSize: 19,
          },
          {
            kind: 'text',
            id: 'marker',
            x: 0.9,
            y: 6.3,
            w: 3,
            h: 0.3,
            text: document.marker,
            fontSize: 11,
          },
          {
            kind: 'text',
            id: 'footer',
            role: 'decoration',
            x: 0.9,
            y: 6.62,
            w: 11.4,
            h: 0.25,
            text: `${document.marker}；来源：${document.source}；待专业审阅`,
            fontSize: 9,
          },
        ],
      })),
    }
    if (document.key === 'finance') {
      const source = sources[2]
      for (const [year, value] of [
        [2023, 383285],
        [2024, 391035],
      ]) {
        if (!source.excerpt.includes(value.toLocaleString('en-US')))
          throw new Error(`finance source missing FY${year} value`)
        const id = `finance-sales-${year}`
        const statement = `FY${year} 净销售额 ${value.toLocaleString('en-US')} 百万美元，取自审计年报。`
        plan.claims.push({
          id,
          statement,
          type: 'calculation',
          sourceIds: [source.id],
          confidence: 'low',
          reviewStatus: 'needs_review',
          calculation: {
            formula: 'amount',
            inputs: ['amount'],
            unit: 'USD millions',
            reproduction: {
              bindings: [{ name: 'amount', inputIndex: 0, value, sourceId: source.id }],
              expected: value,
            },
          },
        })
        deck.claims.push({ id, text: statement, source: source.uri, locator: source.locator })
        plan.slides[5].claimIds.push(id)
        deck.slides[5].claimIds.push(id)
      }
    }
    if (spec.chart) {
      const slide = deck.slides[5]
      const chart = {
        kind: 'chart',
        id: `${document.key}-chart`,
        x: 1,
        y: 2.85,
        w: 10.8,
        h: 3.25,
        chartType: 'bar',
        categories: spec.chart.categories,
        series: [{ name: spec.chart.name, values: spec.chart.values }],
      }
      slide.elements.push(chart)
      const source = sources[3]
      const points = spec.chart.values.map((value, index) => {
        if (document.key === 'finance')
          return {
            value,
            claimId: `finance-sales-${index === 0 ? 2023 : 2024}`,
            basis: { kind: 'calculation' },
          }
        const text = String(value)
        const offset = source.excerpt.indexOf(text)
        if (offset < 0) throw new Error(`${document.key} chart point lacks source: ${text}`)
        return {
          value,
          claimId: claims[3].id,
          basis: { kind: 'source', sourceId: source.id, excerptOffset: offset, excerptText: text },
        }
      })
      plan.slides[5].chartData = [
        {
          elementId: chart.id,
          categories: chart.categories,
          series: [{ name: spec.chart.name, points }],
          unit: document.key === 'science' ? 'people' : 'USD millions',
        },
      ]
    }
    const parsed = parsePresentationPlan(plan)
    assertDeckMatchesPresentationPlan(deck, parsed)
    if (
      spec.chart &&
      checkPresentationChartData(parsed, 'p06', [deck.slides[5].elements.at(-1)]).charts.some(
        (item) => item.findings.length,
      )
    )
      throw new Error(`${document.key} chart data mismatch`)
    const compiled = await compilePresentationDeck(deck)
    const format = (value) =>
      prettier.format(JSON.stringify(value), { parser: 'json', printWidth: 100 })
    const prefix = `reference-${document.key}`
    await writeFile(join(root, `${prefix}-plan.json`), await format(plan))
    await writeFile(join(root, `${prefix}-deck.json`), await format(deck))
    await writeFile(join(root, `${prefix}.pptx`), compiled.bytes)
    files.push(`${prefix}-plan.json`, `${prefix}-deck.json`, `${prefix}.pptx`)
  }
  const names = [...scenario.documents.map((doc) => doc.source), 'scenario.json', ...files]
  await writeFile(
    join(root, 'SHA256SUMS'),
    (
      await Promise.all(
        names.map(async (name) => `${sha(await readFile(join(root, name)))}  ${name}`),
      )
    ).join('\n') + '\n',
  )
  console.log('P0-17 three source-isolated eight-page reference decks generated')
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
