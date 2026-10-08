const { readFile, writeFile } = require('node:fs/promises')
const { createHash } = require('node:crypto')
const { join } = require('node:path')
const prettier = require('prettier')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')
const { checkPresentationChartData } = require('@wiswork/pptx-engine/presentation-chart-data')
const {
  reproducePresentationCalculation,
} = require('@wiswork/pptx-engine/presentation-calculation')
const { presentationSlideSourceLabels } = require('@wiswork/pptx-engine/presentation')
const {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
  presentationPlanClaims,
} = require('@wiswork/pptx-engine/presentation-plan')

const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')

async function main() {
  const root = __dirname
  const input = JSON.parse(await readFile(join(root, 'model-inputs.json'), 'utf8'))
  const rows = (await readFile(join(root, 'scenario-results.csv'), 'utf8'))
    .trim()
    .split('\n')
    .slice(1)
    .map((line) => line.split(','))
  const labels = rows.map((row) => row[1])
  const titles = input.requiredSlides
  const bodies = [
    'Apple 2024 财年历史披露 + WisWork 自制情景；情景不是 Apple 指引。',
    '历史期间：截至 2024-09-28 的十二个月；单位：百万美元；情景仅作演示。',
    '历史净销售额 391,035；营业利润 123,216（百万美元，未经审计）。',
    '下行 -5% / 29.5%；基准 +3% / 31.5%；上行 +8% / 33.5%。按输入表公式逐步取整。',
    '假设性演示，非 Apple 指引或预测。净销售额单位：百万美元。',
    '假设性演示，非 Apple 指引或预测。营业利润单位：百万美元。',
    '增长率和利润率均为敏感假设；模型未覆盖税费、营运资金、现金流及股价。',
    '历史来源：Apple 官方 2024-10-31 财务表；情景由 WisWork 自制，待财务审阅。',
  ]
  const pdfBytes = await readFile(join(root, input.source.file))
  const modelBytes = await readFile(join(root, 'model-inputs.json'))
  const csvBytes = await readFile(join(root, 'scenario-results.csv'))
  const { pdfToPages } = await import('../../../../packages/file-parse/src/pdf.ts')
  const pdfPage = (await pdfToPages(pdfBytes))[input.source.page - 1]
  const line = (needle) => {
    const found = pdfPage.split('\n').find((part) => part.includes(needle))
    if (!found) throw new Error(`historical line missing: ${needle}`)
    return found
  }
  if (input.caseId !== 'PPT-P0-09' || input.asOf !== '2024-10-31' || rows.length !== 3)
    throw new Error('P0-09 frozen scope changed')
  const sources = [
    {
      id: 'historical-sales',
      title: 'Apple FY2024 未审计净销售额',
      uri: input.source.url,
      snapshotAttachmentId: sha(pdfBytes),
      locator: 'PDF 1 FY2024 12 months',
      excerpt: line('Total net sales'),
      asOf: input.source.published,
    },
    {
      id: 'historical-income',
      title: 'Apple FY2024 未审计营业利润',
      uri: input.source.url,
      snapshotAttachmentId: sha(pdfBytes),
      locator: 'PDF 1 FY2024 12 months',
      excerpt: line('Operating income'),
      asOf: input.source.published,
    },
    {
      id: 'model',
      title: 'WisWork 自制假设与逐步取整公式',
      uri: 'local:model-inputs.json',
      snapshotAttachmentId: sha(modelBytes),
      locator: 'model.scenarios / model.formulas',
      excerpt: modelBytes.toString('utf8').slice(0, 600),
      asOf: input.asOf,
    },
    ...rows.map((scenario) => ({
      id: `${scenario[0]}-csv`,
      title: `${scenario[1]}独立复算`,
      uri: 'local:scenario-results.csv',
      snapshotAttachmentId: sha(csvBytes),
      locator: `CSV ${scenario[0]}`,
      excerpt: scenario.join(','),
      asOf: input.asOf,
    })),
  ]
  const context = (limitations, formula) => ({
    domain: 'finance',
    materialKind: 'financial_statement',
    reportingPeriod: 'FY2024 ended 2024-09-28; illustrative next twelve months',
    asOf: input.asOf,
    currency: 'USD',
    unit: 'USD millions',
    accountingBasis: 'unaudited consolidated U.S. GAAP',
    limitations,
    ...(formula ? { formula } : {}),
  })
  const claim = (id, statement, sourceIds, type, limitations, formula) => ({
    id,
    statement,
    sourceIds,
    type,
    confidence: 'low',
    reviewStatus: 'needs_review',
    asOf: input.asOf,
    professionalContext: context(limitations, formula),
  })
  const claims = [
    claim(
      'historical',
      'Apple FY2024 截至 2024-09-28 的十二个月，净销售额 391,035、营业利润 123,216 百万美元，官方表格未经审计。',
      ['historical-sales', 'historical-income'],
      'fact',
      '不可把未经审计表称为审计年报',
    ),
    claim(
      'model-boundary',
      '三种情景由 WisWork 自制，仅是假设性演示，非 Apple 指引或预测。',
      ['model'],
      'judgment',
      '未覆盖税费、营运资金、现金流或股价',
    ),
  ]
  const formula = {
    sales: 'round(historical*(10000+bps)/10000,0)',
    income: 'round(sales*margin/10000,0)',
  }
  const chartClaimIds = { sales: [], income: [] }
  for (const scenario of rows) {
    const [id, label, growth, margin, sales, income, evidenceType] = scenario
    if (evidenceType !== 'calculated_hypothesis') throw new Error('scenario provenance changed')
    for (const kind of ['sales', 'income']) {
      const amount = Number(kind === 'sales' ? sales : income)
      const claimId = `${id}-${kind}`
      chartClaimIds[kind].push(claimId)
      const bindings =
        kind === 'sales'
          ? [
              {
                name: 'historical',
                inputIndex: 0,
                value: input.historical.netSales,
                sourceId: 'model',
              },
              { name: 'bps', inputIndex: 1, value: Number(growth), sourceId: 'model' },
            ]
          : [
              { name: 'sales', inputIndex: 0, value: Number(sales), sourceId: `${id}-csv` },
              { name: 'margin', inputIndex: 1, value: Number(margin), sourceId: 'model' },
            ]
      const calculation = {
        formula: formula[kind],
        inputs: bindings.map((binding) => binding.name),
        unit: 'USD millions',
        currency: 'USD',
        reproduction: { bindings, expected: amount },
      }
      const next = {
        ...claim(
          claimId,
          `${label}${kind === 'sales' ? '净销售额' : '营业利润'} ${amount.toLocaleString('en-US')} 百万美元，仅为自制假设结果。`,
          ['model', `${id}-csv`],
          'calculation',
          '假设性演示，非 Apple 指引或预测；不推断未建模指标',
          formula[kind],
        ),
        calculation,
      }
      if (reproducePresentationCalculation(next).status !== 'reproduced')
        throw new Error(`scenario arithmetic mismatch: ${claimId}`)
      claims.push(next)
    }
  }
  const style = {
    fontFace: 'Noto Sans CJK SC',
    background: 'FFFFFF',
    textColor: '183247',
    accentColor: '087D83',
  }
  const projectId = 'p0-09-hypothetical-reference'
  const claimIds = [
    ['model-boundary'],
    ['historical', 'model-boundary'],
    ['historical'],
    ['model-boundary'],
    chartClaimIds.sales,
    chartClaimIds.income,
    ['model-boundary'],
    ['historical', 'model-boundary'],
  ]
  const sections = [
    'financial_question',
    'reporting_basis',
    'reporting_basis',
    'financial_results',
    'financial_risks_and_scenarios',
    'financial_risks_and_scenarios',
    'financial_risks_and_scenarios',
    'financial_actions',
  ]
  const plan = {
    version: 1,
    projectId,
    title: 'P0-09 情景分析参考稿：非 Apple 指引',
    domain: 'finance',
    brief: {
      objective: '严格区分 Apple FY2024 未审计历史披露与 WisWork 自制三种假设情景',
      audience: '财务专业审阅者',
      language: 'zh-CN',
      minutes: 8,
      requiredContent: [
        '历史期间和审计状态',
        '三情景参数与逐步取整公式',
        '两张原生图表',
        '非 Apple 指引提示',
      ],
      constraints: [...input.mustNotClaim, '财务审阅和真实 PowerPoint 验收待完成'],
    },
    sources,
    claims,
    style,
    slides: titles.map((title, index) => {
      const kind = index === 4 ? 'sales' : index === 5 ? 'income' : null
      return {
        id: `p${String(index + 1).padStart(2, '0')}`,
        title,
        purpose: '将历史披露与自制假设及受限公式分开呈现',
        claimIds: claimIds[index],
        domainSection: sections[index],
        layout: index === 0 ? 'cover' : index === 7 ? 'summary' : kind ? 'chart' : 'content',
        requiredAssets: [],
        acceptanceCriteria: ['历史数据有官方原文', '情景计算可复算', '原生图表保留假设标签'],
        ...(kind
          ? {
              chartData: [
                {
                  elementId: kind === 'sales' ? 'scenario-sales' : 'scenario-operating-income',
                  categories: labels,
                  currency: 'USD',
                  unit: 'USD millions',
                  series: [
                    {
                      name: kind === 'sales' ? '净销售额（百万美元）' : '营业利润（百万美元）',
                      points: rows.map((scenario) => ({
                        value: Number(scenario[kind === 'sales' ? 4 : 5]),
                        claimId: `${scenario[0]}-${kind}`,
                        basis: { kind: 'calculation' },
                      })),
                    },
                  ],
                },
              ],
            }
          : {}),
      }
    }),
  }
  const parsedPlan = parsePresentationPlan(plan)
  const deck = {
    version: 1,
    id: projectId,
    title: plan.title,
    style,
    assets: [],
    claims: presentationPlanClaims(parsedPlan),
    slides: titles.map((title, index) => ({
      id: `p${String(index + 1).padStart(2, '0')}`,
      title,
      claimIds: plan.slides[index].claimIds,
      notes: '参考稿：历史披露与自制假设分开；非 Apple 指引或预测。',
      elements: [
        { kind: 'text', id: 'title', x: 0.65, y: 0.48, w: 12, h: 0.7, text: title, fontSize: 26 },
        {
          kind: 'text',
          id: 'body',
          x: 0.9,
          y: 1.6,
          w: 11.4,
          h: index === 4 || index === 5 ? 0.65 : 2.7,
          text: bodies[index],
          fontSize: index === 3 ? 17 : 19,
        },
        ...(index === 4 || index === 5
          ? [
              {
                kind: 'chart',
                id: index === 4 ? 'scenario-sales' : 'scenario-operating-income',
                x: 1.0,
                y: 2.6,
                w: 10.8,
                h: 3.7,
                chartType: 'bar',
                categories: labels,
                series: [
                  {
                    name: index === 4 ? '净销售额（百万美元）' : '营业利润（百万美元）',
                    values: rows.map((row) => Number(row[index === 4 ? 4 : 5])),
                  },
                ],
              },
            ]
          : []),
        {
          kind: 'text',
          id: 'footer',
          role: 'decoration',
          x: 0.9,
          y: 6.62,
          w: 11.4,
          h: 0.25,
          text:
            index === 2
              ? '来源：Apple 官方 2024 财年财务表，第 1 页；未经审计'
              : '情景假设：WisWork 自制，仅供基准测试；非 Apple 指引',
          fontSize: 9,
        },
      ],
    })),
  }
  assertDeckMatchesPresentationPlan(deck, parsedPlan)
  for (const slide of deck.slides) {
    const length = presentationSlideSourceLabels(slide, deck.claims).join('；').length
    if (length > 500) throw new Error(`source footer overflow: ${slide.id} ${length}`)
  }
  for (const index of [4, 5]) {
    const chart = deck.slides[index].elements.find((element) => element.kind === 'chart')
    const report = checkPresentationChartData(parsedPlan, deck.slides[index].id, [chart])
    if (report.charts.some((item) => item.findings.length))
      throw new Error(`chart data invalid: ${index}`)
  }
  const result = await compilePresentationDeck(deck)
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
  await writeFile(join(root, 'p0-09-reference.pptx'), result.bytes)
  console.log(`P0-09 reference: ${result.bytes.length} bytes, ${deck.slides.length} pages`)
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
