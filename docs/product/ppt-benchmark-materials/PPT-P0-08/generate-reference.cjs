const { readFile, writeFile } = require('node:fs/promises')
const { createHash } = require('node:crypto')
const { join } = require('node:path')
const prettier = require('prettier')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')
const { presentationSlideSourceLabels } = require('@wiswork/pptx-engine/presentation')
const { checkPresentationChartData } = require('@wiswork/pptx-engine/presentation-chart-data')
const {
  reproducePresentationCalculation,
} = require('@wiswork/pptx-engine/presentation-calculation')
const {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
  presentationPlanClaims,
} = require('@wiswork/pptx-engine/presentation-plan')

const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')
function row(page, anchor) {
  const found = page.split('\n').find((line) => line.includes(anchor))
  if (!found) throw new Error(`official row missing: ${anchor}`)
  return found
}

async function main() {
  const root = __dirname
  const basis = JSON.parse(await readFile(join(root, 'basis.json'), 'utf8'))
  const apple = basis.companies[0]
  const toyota = basis.companies[1]
  const titles = [
    '跨公司收入可比性：先核对口径',
    '资料时点与两份年度报告',
    '财年、准则、币种和范围矩阵',
    'Apple：原币净销售额',
    'Toyota：原币销售及金融服务收入',
    'Toyota 汇率换算仅作示意',
    '不可比项与留空规则',
    '结论、审阅和宿主验收',
  ]
  const bodies = [
    '两家公司披露的财年、会计准则和收入范围不同；本稿不作收入差额或排名。',
    `as-of ${basis.asOf}。Apple 10-K 于 ${apple.filingDate} 披露；Toyota 20-F 于 ${toyota.filingDate} 披露，汇率汇总于 ${basis.illustrativeFx.publicationDate} 发布。`,
    'Apple：2024-09-28 / US GAAP / 百万美元 / 净销售额。Toyota：2024-03-31 / IFRS / 百万日元 / 销售及金融服务收入。',
    `${apple.valueMillions.toLocaleString('en-US')} 百万美元；Apple 2024 Form 10-K，PDF 第 ${apple.reportPdfPage} 页。不得与 Toyota 原币图共用数值轴。`,
    `${toyota.valueMillions.toLocaleString('en-US')} 百万日元；Toyota Form 20-F，PDF 第 ${toyota.valueSourcePdfPage} 页。包括金融服务收入。`,
    `${toyota.valueMillions.toLocaleString('en-US')} ÷ ${basis.illustrativeFx.jpyPerUsd} = ${basis.illustrativeFx.toyotaUsdMillionsRounded2.toLocaleString('en-US')} 百万美元。汇率取 Toyota FY2024 历史 12 个月栏，官方汇总 PDF 第 ${basis.illustrativeFx.sourcePdfPage} 页；不是 FY2025 预测栏。`,
    '同口径美元收入差额：空缺；跨公司收入排名：空缺。原因：财年结束日、准则及收入范围不同，货币换算不能消除差异。',
    '保留原币、原期间和原始口径；示意换算不支持相对规模结论。官方原件及本地摘要已冻结；仍待财务审阅及真实 PowerPoint 重开编辑证据。',
  ]
  const appleBytes = await readFile(join(root, apple.localFile))
  const toyotaBytes = await readFile(join(root, toyota.localFile))
  const fxBytes = await readFile(join(root, 'toyota-fy2024-financial-summary.pdf'))
  const basisBytes = await readFile(join(root, 'basis.json'))
  const csvBytes = await readFile(join(root, 'independent-recalc.csv'))
  const { pdfToPages } = await import('../../../../packages/file-parse/src/pdf.ts')
  const applePages = await pdfToPages(appleBytes)
  const toyotaPages = await pdfToPages(toyotaBytes)
  const fxPages = await pdfToPages(fxBytes)
  const fxPage = fxPages[basis.illustrativeFx.sourcePdfPage - 1]
  const csvRow = csvBytes.toString('utf8').trim().split(/\r?\n/)[1]
  if (
    basis.caseId !== 'PPT-P0-08' ||
    basis.asOf !== '2024-11-01' ||
    basis.comparison.likeForLikeUsdRevenueDifferenceMillions !== null ||
    basis.comparison.crossCompanyRevenueRanking !== null ||
    csvRow.split(',')[1] !== String(toyota.valueMillions) ||
    csvRow.split(',')[3] !== String(basis.illustrativeFx.jpyPerUsd)
  )
    throw new Error('P0-08 frozen comparison boundary changed')
  const sources = [
    {
      id: 'apple-report',
      title: 'Apple 审计年报净销售额',
      uri: apple.source,
      snapshotAttachmentId: sha(appleBytes),
      locator: `PDF ${apple.reportPdfPage}`,
      excerpt: row(applePages[apple.reportPdfPage - 1], 'Total net sales'),
      asOf: apple.filingDate,
    },
    {
      id: 'toyota-report',
      title: 'Toyota 20-F 销售收入',
      uri: toyota.source,
      snapshotAttachmentId: sha(toyotaBytes),
      locator: `PDF ${toyota.valueSourcePdfPage}`,
      excerpt: row(toyotaPages[toyota.valueSourcePdfPage - 1], 'Total sales revenues'),
      asOf: toyota.filingDate,
    },
    {
      id: 'toyota-finance',
      title: 'Toyota 20-F 金融服务收入',
      uri: toyota.source,
      snapshotAttachmentId: sha(toyotaBytes),
      locator: `PDF ${toyota.valueSourcePdfPage}`,
      excerpt: row(toyotaPages[toyota.valueSourcePdfPage - 1], 'Financial services'),
      asOf: toyota.filingDate,
    },
    {
      id: 'fx-history',
      title: 'Toyota FY2024 历史十二个月汇率栏',
      uri: basis.illustrativeFx.source,
      snapshotAttachmentId: sha(fxBytes),
      locator: `PDF ${basis.illustrativeFx.sourcePdfPage} FY2024 12 months`,
      excerpt: fxPage.slice(0, 430),
      asOf: basis.illustrativeFx.publicationDate,
    },
    ...[
      [apple, 'apple'],
      [toyota, 'toyota'],
    ].map(([company, id]) => ({
      id: `${id}-value`,
      title: `${company.name} ${company.fiscalYearEnd} 原币收入字段`,
      uri: 'local:basis.json',
      snapshotAttachmentId: sha(basisBytes),
      locator: `companies.${id}.valueMillions`,
      excerpt: String(company.valueMillions),
      asOf: basis.asOf,
    })),
    {
      id: 'fx-csv',
      title: '独立汇率示意复算行',
      uri: 'local:independent-recalc.csv',
      snapshotAttachmentId: sha(csvBytes),
      locator: 'CSV toyota_total_sales_revenues',
      excerpt: csvRow,
      asOf: basis.asOf,
    },
  ]
  const context = (company, limitations, formula) => ({
    domain: 'finance',
    materialKind: 'financial_statement',
    reportingPeriod: `FY2024 ended ${company.fiscalYearEnd}`,
    asOf: basis.asOf,
    currency: company.currency,
    unit: `${company.currency} millions`,
    accountingBasis: company.accounting,
    limitations,
    ...(formula ? { formula } : {}),
  })
  const claim = (id, statement, sourceIds, professionalContext, type = 'fact') => ({
    id,
    statement,
    sourceIds,
    professionalContext,
    type,
    confidence: 'low',
    reviewStatus: 'needs_review',
    asOf: basis.asOf,
  })
  const claims = [
    claim(
      'apple-sales',
      `Apple FY2024 净销售额为 ${apple.valueMillions.toLocaleString('en-US')} 百万美元。`,
      ['apple-report', 'apple-value'],
      context(apple, '与 Toyota 财年、准则和收入范围不同；不可直接排名'),
    ),
    claim(
      'toyota-sales',
      `Toyota FY2024 销售收入为 ${toyota.valueMillions.toLocaleString('en-US')} 百万日元，含金融服务收入。`,
      ['toyota-report', 'toyota-finance', 'toyota-value'],
      context(toyota, '销售及金融服务收入与 Apple 净销售额不可直接比较'),
    ),
    claim(
      'fx-period',
      'Toyota 官方汇总的 FY2024 已发生十二个月栏使用 1 美元兑 145 日元；右侧 FY2025 预测栏不能充当历史依据。',
      ['fx-history'],
      context(toyota, '历史年度汇率非 2024-11-01 即期汇率，也非 Apple 财年汇率'),
    ),
    {
      ...claim(
        'fx-illustrative',
        `Toyota 收入按其 FY2024 历史汇率示意换算为 ${basis.illustrativeFx.toyotaUsdMillionsRounded2.toLocaleString('en-US')} 百万美元。`,
        ['toyota-report', 'fx-history', 'fx-csv'],
        context(
          toyota,
          '仅示意换算，不能据此与 Apple 计算同口径差额或排名',
          basis.illustrativeFx.formula,
        ),
        'calculation',
      ),
      calculation: {
        formula: basis.illustrativeFx.reproductionFormula,
        inputs: ['jpy', 'rate'],
        unit: 'USD millions',
        currency: 'USD',
        reproduction: {
          bindings: [
            { name: 'jpy', inputIndex: 0, value: toyota.valueMillions, sourceId: 'fx-csv' },
            {
              name: 'rate',
              inputIndex: 1,
              value: basis.illustrativeFx.jpyPerUsd,
              sourceId: 'fx-csv',
            },
          ],
          expected: basis.illustrativeFx.toyotaUsdMillionsRounded2,
        },
      },
    },
    claim(
      'comparison-gap',
      '两公司财年结束日、会计准则和收入范围不同；同口径美元差额与跨公司收入排名均须留空。',
      ['apple-report', 'toyota-report', 'toyota-finance'],
      {
        domain: 'finance',
        materialKind: 'disclosure',
        reportingPeriod: 'non-aligned FY2024 periods',
        asOf: basis.asOf,
        currency: 'USD and JPY',
        unit: 'millions',
        accountingBasis: 'US GAAP versus IFRS',
        limitations: '无可靠同口径调整；不发布差额或排名',
      },
      'judgment',
    ),
  ]
  if (reproducePresentationCalculation(claims[3]).status !== 'reproduced')
    throw new Error('illustrative FX formula not reproduced')
  const style = {
    fontFace: 'Noto Sans CJK SC',
    background: 'FFFFFF',
    textColor: '173248',
    accentColor: '087D83',
  }
  const projectId = 'p0-08-cross-company-comparability-reference'
  const claimIds = [
    [],
    ['apple-sales', 'toyota-sales'],
    ['apple-sales', 'toyota-sales'],
    ['apple-sales'],
    ['toyota-sales'],
    ['fx-period', 'fx-illustrative'],
    ['comparison-gap'],
    ['comparison-gap'],
  ]
  const sections = [
    'financial_question',
    'reporting_basis',
    'reporting_basis',
    'financial_results',
    'financial_results',
    'financial_results',
    'financial_risks_and_scenarios',
    'financial_actions',
  ]
  const plan = {
    version: 1,
    projectId,
    title: 'Apple 与 Toyota 跨公司可比性候选稿',
    domain: 'finance',
    brief: {
      objective: '先核对两公司财年、会计准则、币种、单位和收入范围，仅展示有来源的历史汇率示意换算',
      audience: '财务专业审阅者',
      language: 'zh-CN',
      minutes: 8,
      requiredContent: ['口径矩阵', '两张各自原币图表', '历史 FY2024 汇率', '不可比项空白结果'],
      constraints: [
        '不得计算同口径美元差额',
        '不得排名',
        '不把 FY2025 预测汇率当历史汇率',
        '财务与真实 PowerPoint 验收待完成',
      ],
    },
    sources,
    claims,
    style,
    slides: titles.map((title, index) => {
      const company = index === 3 ? apple : index === 4 ? toyota : null
      const id = index === 3 ? 'apple' : 'toyota'
      const valueSource = sources.find((source) => source.id === `${id}-value`)
      return {
        id: `p${String(index + 1).padStart(2, '0')}`,
        title,
        purpose: '保留原币与不可比边界，示意换算独立呈现',
        claimIds: claimIds[index],
        domainSection: sections[index],
        layout: index === 0 ? 'cover' : index === 7 ? 'summary' : company ? 'chart' : 'content',
        requiredAssets: [],
        acceptanceCriteria: ['原生可编辑表格和图表', '源数值可追溯', '不可比结果保持空白'],
        ...(company
          ? {
              chartData: [
                {
                  elementId: `${id}-native-chart`,
                  categories: [company.name],
                  currency: company.currency,
                  unit: `${company.currency} millions`,
                  series: [
                    {
                      name: `${company.metric}（百万${company.currency === 'USD' ? '美元' : '日元'}）`,
                      points: [
                        {
                          value: company.valueMillions,
                          claimId: `${id}-sales`,
                          basis: {
                            kind: 'source',
                            sourceId: valueSource.id,
                            excerptOffset: 0,
                            excerptText: valueSource.excerpt,
                          },
                        },
                      ],
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
    slides: titles.map((title, index) => {
      const company = index === 3 ? apple : index === 4 ? toyota : null
      return {
        id: `p${String(index + 1).padStart(2, '0')}`,
        title,
        claimIds: plan.slides[index].claimIds,
        notes: '候选参考稿；官方原件已冻结，本稿未经财务审阅和 PowerPoint 宿主验收。',
        elements: [
          { kind: 'text', id: 'title', x: 0.65, y: 0.48, w: 12, h: 0.7, text: title, fontSize: 26 },
          {
            kind: 'text',
            id: 'body',
            x: 0.9,
            y: 1.6,
            w: 11.4,
            h: company || index === 2 || index === 6 ? 0.8 : 3.5,
            text: bodies[index],
            fontSize: 18,
          },
          ...(index === 2
            ? [
                {
                  kind: 'table',
                  id: 'basis-matrix',
                  x: 0.9,
                  y: 2.65,
                  w: 11.4,
                  h: 3.3,
                  fontSize: 15,
                  rows: [
                    ['核对项', 'Apple FY2024', 'Toyota FY2024'],
                    ['财年结束', apple.fiscalYearEnd, toyota.fiscalYearEnd],
                    ['会计准则', apple.accounting, toyota.accounting],
                    ['币种与单位', 'USD millions', 'JPY millions'],
                    ['收入范围', apple.metric, toyota.metric],
                  ],
                },
              ]
            : []),
          ...(index === 6
            ? [
                {
                  kind: 'table',
                  id: 'comparison-gaps',
                  x: 0.9,
                  y: 2.65,
                  w: 11.4,
                  h: 2.7,
                  fontSize: 16,
                  rows: [
                    ['比较项', '结果', '理由'],
                    ['同口径美元收入差额', '', '财年、准则和范围不同'],
                    ['跨公司收入排名', '', '换算不能消除口径差异'],
                  ],
                },
              ]
            : []),
          ...(company
            ? [
                {
                  kind: 'chart',
                  id: `${index === 3 ? 'apple' : 'toyota'}-native-chart`,
                  x: 1,
                  y: 2.8,
                  w: 10.8,
                  h: 3.2,
                  chartType: 'bar',
                  categories: [company.name],
                  series: [
                    {
                      name: `${company.metric}（百万${company.currency === 'USD' ? '美元' : '日元'}）`,
                      values: [company.valueMillions],
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
              index === 3
                ? '来源：Apple FY2024 Form 10-K，PDF 第 32 页'
                : index === 4
                  ? '来源：Toyota FY2024 Form 20-F，PDF 第 167 页；原件已冻结'
                  : '参考稿；资料与专业结论待审阅',
            fontSize: 9,
          },
        ],
      }
    }),
  }
  assertDeckMatchesPresentationPlan(deck, parsedPlan)
  for (const slide of deck.slides) {
    const sourceText = presentationSlideSourceLabels(slide, deck.claims).join('；')
    if (sourceText.length > 500) throw new Error(`source footer ${slide.id}: ${sourceText.length}`)
  }
  for (const index of [3, 4]) {
    const chart = deck.slides[index].elements.find((element) => element.kind === 'chart')
    const report = checkPresentationChartData(parsedPlan, plan.slides[index].id, [chart])
    if (report.charts.some((item) => item.findings.length))
      throw new Error(`declared chart data invalid: ${plan.slides[index].id}`)
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
  await writeFile(join(root, 'p0-08-reference.pptx'), result.bytes)
  console.log(
    `P0-08 source-backed reference: ${result.bytes.length} bytes, ${deck.slides.length} pages`,
  )
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
