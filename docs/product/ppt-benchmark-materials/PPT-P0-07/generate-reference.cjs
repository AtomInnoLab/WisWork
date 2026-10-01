const { readFile, writeFile } = require('node:fs/promises')
const { createHash } = require('node:crypto')
const { join } = require('node:path')
const prettier = require('prettier')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')
const { presentationSlideSourceLabels } = require('@wiswork/pptx-engine/presentation')
const { checkPresentationChartData } = require('@wiswork/pptx-engine/presentation-chart-data')
const {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
  presentationPlanClaims,
} = require('@wiswork/pptx-engine/presentation-plan')

const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')
function lineContaining(page, text, values = []) {
  const found = page
    .split('\n')
    .find((line) => line.includes(text) && values.every((value) => line.includes(value)))
  if (!found) throw new Error(`frozen financial row missing: ${text}`)
  return found
}

async function main() {
  const root = __dirname
  const dictionary = JSON.parse(await readFile(join(root, 'data-dictionary.json'), 'utf8'))
  const rows = (await readFile(join(root, 'independent-recalc.csv'), 'utf8'))
    .trim()
    .split('\n')
    .slice(1)
    .map((line) => line.split(','))
  const titles = dictionary.requiredSlides
  const bodies = [
    'Apple 2024 财年；审计年报为主，官方业绩披露仅交叉核对。',
    '2024 财年截至 9 月 28 日，2023 财年截至 9 月 30 日；合并口径，单位：百万美元。',
    `2023：383,285；2024：391,035；同比 +${rows[0][4]}%。审计年报第 29 页。`,
    `2023：96,995；2024：93,736；同比 ${rows[1][4]}%，为下降。审计年报第 29 页。`,
    `2023：110,543；2024：118,254；同比 +${rows[2][4]}%。审计年报第 33 页。`,
    '同比 = (2024 值－2023 值) ÷ 2023 值 × 100%；四舍五入到百分比小数点后两位。',
    '净利润同比下降；官方 Q4 业绩表未经审计，不能替代 Form 10-K 审计意见。',
    '资料时点 2024-11-01；指标和同比待财务审阅，真实 PowerPoint 编辑与重开待验收。',
  ]
  const dictionaryBytes = await readFile(join(root, 'data-dictionary.json'))
  const csvBytes = await readFile(join(root, 'independent-recalc.csv'))
  const annualBytes = await readFile(join(root, dictionary.annualReport.file))
  const releaseBytes = await readFile(join(root, dictionary.earningsRelease.file))
  const { pdfToPages } = await import('../../../../packages/file-parse/src/pdf.ts')
  const annualPages = await pdfToPages(annualBytes)
  const releasePages = await pdfToPages(releaseBytes)
  const csvLines = csvBytes.toString('utf8').trim().split(/\r?\n/).slice(1)
  const csvRows = new Map(csvLines.map((line) => [line.split(',')[0], line]))
  if (dictionary.caseId !== 'PPT-P0-07' || dictionary.asOf !== '2024-11-01' || csvRows.size !== 3)
    throw new Error('P0-07 frozen scope changed')
  const sources = [
    {
      id: 'annual-period',
      title: 'Apple FY2024 Form 10-K 合并报表期间与单位',
      uri: dictionary.annualReport.sourceUrl,
      snapshotAttachmentId: sha(annualBytes),
      locator: 'PDF 第 32 页／印刷第 29 页',
      excerpt: annualPages[31].slice(0, 220),
      asOf: dictionary.annualReport.filingDate,
    },
    {
      id: 'audit-opinion',
      title: 'Apple FY2024 Form 10-K 独立审计意见',
      uri: dictionary.annualReport.sourceUrl,
      snapshotAttachmentId: sha(annualBytes),
      locator: 'PDF 第 51 页／印刷第 48 页',
      excerpt: lineContaining(annualPages[50], 'We have audited the accompanying'),
      asOf: dictionary.annualReport.filingDate,
    },
    {
      id: 'release-unaudited',
      title: 'Apple FY2024 Q4 官方业绩表未审计标识',
      uri: dictionary.earningsRelease.sourceUrl,
      snapshotAttachmentId: sha(releaseBytes),
      locator: 'PDF 第 1 页',
      excerpt: lineContaining(releasePages[0], '(Unaudited)'),
      asOf: dictionary.earningsRelease.published,
    },
    {
      id: 'dictionary',
      title: '本包自制口径字典',
      uri: 'local:data-dictionary.json',
      snapshotAttachmentId: sha(dictionaryBytes),
      locator: 'asOf、currency、unit、basis',
      excerpt: dictionaryBytes.toString('utf8').slice(0, 430),
      asOf: dictionary.asOf,
    },
    ...dictionary.metrics.flatMap((metric) => {
      const values = [metric.fy2024, metric.fy2023].map((value) => value.toLocaleString('en-US'))
      const row = csvRows.get(metric.id)
      if (
        !row ||
        row.split(',')[2] !== String(metric.fy2024) ||
        row.split(',')[3] !== String(metric.fy2023)
      )
        throw new Error(`frozen independent calculation changed: ${metric.id}`)
      return [
        {
          id: `${metric.id}-annual`,
          title: `审计年报 ${metric.label} 原值`,
          uri: dictionary.annualReport.sourceUrl,
          snapshotAttachmentId: sha(annualBytes),
          locator: `PDF ${metric.annualPdfPage} / 印 ${metric.annualPrintedPage}`,
          excerpt: lineContaining(annualPages[metric.annualPdfPage - 1], metric.annualRow, values),
          asOf: dictionary.annualReport.filingDate,
        },
        {
          id: `${metric.id}-release`,
          title: `未经审计业绩表 ${metric.label} 交叉核对`,
          uri: dictionary.earningsRelease.sourceUrl,
          snapshotAttachmentId: sha(releaseBytes),
          locator: `PDF 第 ${metric.releasePdfPage} 页`,
          excerpt: lineContaining(
            releasePages[metric.releasePdfPage - 1],
            metric.annualRow,
            values,
          ),
          asOf: dictionary.earningsRelease.published,
        },
        {
          id: `${metric.id}-csv`,
          title: `独立复算表 ${metric.label}`,
          uri: 'local:independent-recalc.csv',
          snapshotAttachmentId: sha(csvBytes),
          locator: `CSV ${metric.id} 行`,
          excerpt: row,
          asOf: dictionary.asOf,
        },
        ...[2023, 2024].map((year) => ({
          id: `${metric.id}-${year}-csv`,
          title: `独立复算表 ${metric.label} ${year} 原值`,
          uri: 'local:independent-recalc.csv',
          snapshotAttachmentId: sha(csvBytes),
          locator: `CSV ${metric.id} FY${year}`,
          excerpt: String(metric[`fy${year}`]),
          asOf: dictionary.asOf,
        })),
      ]
    }),
  ]
  const financeContext = (limitations, extra = {}) => ({
    domain: 'finance',
    materialKind: 'financial_statement',
    reportingPeriod: 'fiscal years ended 2024-09-28 and 2023-09-30',
    asOf: dictionary.asOf,
    currency: 'USD',
    unit: 'USD millions',
    accountingBasis: 'consolidated U.S. GAAP',
    limitations,
    ...extra,
  })
  const fact = (id, statement, sourceIds, limitations, extra = {}) => ({
    id,
    statement,
    sourceIds,
    type: 'fact',
    confidence: 'low',
    reviewStatus: 'needs_review',
    asOf: dictionary.asOf,
    professionalContext: financeContext(limitations, extra),
  })
  const claims = [
    fact(
      'basis',
      '2024 与 2023 财年结束日分别为 2024-09-28、2023-09-30；合并报表单位为百万美元。',
      ['annual-period', 'dictionary'],
      '两财年结束日不同；重述口径待财务审阅',
    ),
    fact(
      'audit',
      '主要年度数值取自包含独立审计意见的 Apple FY2024 Form 10-K。',
      ['audit-opinion'],
      '审计意见身份不认证本稿财务解释',
    ),
    fact(
      'release-status',
      '2024-10-31 官方 Q4 业绩表标为 Unaudited，仅用于交叉核对。',
      ['release-unaudited'],
      '不能把业绩表称为审计年报',
      { materialKind: 'disclosure' },
    ),
  ]
  for (const metric of dictionary.metrics) {
    const percent = Number(csvRows.get(metric.id).split(',')[4])
    for (const year of [2023, 2024]) {
      const value = metric[`fy${year}`]
      claims.push(
        fact(
          `${metric.id}-${year}`,
          `${metric.label} ${year} 财年为 ${value.toLocaleString('en-US')} 百万美元。`,
          [`${metric.id}-annual`, `${metric.id}-${year}-csv`],
          '审计年报原值为主，未经审计业绩表仅交叉核对；重述口径待审阅',
        ),
      )
    }
    claims.push({
      ...fact(
        `${metric.id}-yoy`,
        `${metric.label} 2024 相对 2023 财年同比 ${percent > 0 ? '+' : ''}${percent.toFixed(2)}%，为本包独立复算。`,
        [`${metric.id}-csv`],
        '独立复算而非 Apple 原文百分比；不作未来预测',
        { formula: '(2024−2023)÷2023×100%，小数点后两位四舍五入' },
      ),
      type: 'calculation',
      calculation: {
        formula: 'round((fy2024-fy2023)/fy2023*100,2)',
        inputs: ['fy2024', 'fy2023'],
        unit: '%',
        reproduction: {
          bindings: [
            { name: 'fy2024', inputIndex: 0, value: metric.fy2024, sourceId: `${metric.id}-csv` },
            { name: 'fy2023', inputIndex: 1, value: metric.fy2023, sourceId: `${metric.id}-csv` },
          ],
          expected: percent,
        },
      },
    })
  }
  const claimIds = [
    ['basis', 'audit', 'release-status'],
    ['basis', 'audit', 'release-status'],
    ...dictionary.metrics.map((metric) => [
      `${metric.id}-2023`,
      `${metric.id}-2024`,
      `${metric.id}-yoy`,
    ]),
    dictionary.metrics.map((metric) => `${metric.id}-yoy`),
    ['net_income-yoy', 'release-status'],
    ['basis', 'audit', 'release-status'],
  ]
  const style = {
    fontFace: 'Noto Sans CJK SC',
    background: 'FFFFFF',
    textColor: '173248',
    accentColor: '087D83',
  }
  const projectId = 'p0-07-annual-performance-reference'
  const sections = [
    'financial_question',
    'reporting_basis',
    'financial_results',
    'financial_results',
    'financial_results',
    'financial_results',
    'financial_risks_and_scenarios',
    'financial_actions',
  ]
  const plan = {
    version: 1,
    projectId,
    title: 'Apple 2024 财年表现候选稿',
    domain: 'finance',
    brief: {
      objective: '依据审计年报制作八页年度表现汇报，独立复算同比并保留未经审计业绩表的交叉核对身份',
      audience: '财务专业审阅者',
      language: 'zh-CN',
      minutes: 8,
      requiredContent: [
        '不同财年结束日',
        '百万美元单位',
        '三指标年度原值与同比',
        '三张原生图表',
        '审计身份与限制',
      ],
      constraints: [
        '净利润为下降',
        '不把 Q4 业绩表称为审计报告',
        '重述口径待审阅',
        '不得提供投资建议',
        '真实 PowerPoint 验收待完成',
      ],
    },
    sources,
    claims,
    style,
    slides: titles.map((title, index) => {
      const metric = dictionary.metrics[index - 2]
      return {
        id: `p${String(index + 1).padStart(2, '0')}`,
        title,
        purpose: '展示已冻结年度数据、独立复算与审计边界',
        claimIds: claimIds[index],
        domainSection: sections[index],
        layout: index === 0 ? 'cover' : index === 7 ? 'summary' : metric ? 'chart' : 'content',
        requiredAssets: [],
        acceptanceCriteria: [
          '原生可编辑图表或文字',
          '原值与公式可追溯',
          '保留报告期、单位和审计身份',
        ],
        ...(metric
          ? {
              chartData: [
                {
                  elementId: `${metric.id}-chart`,
                  categories: ['2023 财年', '2024 财年'],
                  currency: 'USD',
                  unit: 'USD millions',
                  series: [
                    {
                      name: `${metric.label}（百万美元）`,
                      points: [2023, 2024].map((year) => {
                        const value = metric[`fy${year}`]
                        const text = String(value)
                        const csvSource = sources.find(
                          (source) => source.id === `${metric.id}-${year}-csv`,
                        )
                        if (csvSource.excerpt !== text)
                          throw new Error(`numeric CSV basis missing: ${metric.id}-${year}`)
                        return {
                          value,
                          claimId: `${metric.id}-${year}`,
                          basis: {
                            kind: 'source',
                            sourceId: csvSource.id,
                            excerptOffset: 0,
                            excerptText: text,
                          },
                        }
                      }),
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
      const metric = dictionary.metrics[index - 2]
      return {
        id: `p${String(index + 1).padStart(2, '0')}`,
        title,
        claimIds: plan.slides[index].claimIds,
        notes: '参考稿：年度审计数值与未经审计业绩披露分开；同比为独立计算。',
        elements: [
          { kind: 'text', id: 'title', x: 0.65, y: 0.48, w: 12, h: 0.7, text: title, fontSize: 26 },
          {
            kind: 'text',
            id: 'body',
            x: 0.9,
            y: 1.6,
            w: 11.4,
            h: metric ? 0.65 : 2.7,
            text: bodies[index],
            fontSize: 18,
          },
          ...(metric
            ? [
                {
                  kind: 'chart',
                  id: `${metric.id}-chart`,
                  x: 1.0,
                  y: 2.6,
                  w: 10.8,
                  h: 3.65,
                  chartType: 'bar',
                  categories: ['2023 财年', '2024 财年'],
                  series: [
                    {
                      name: `${metric.label}（百万美元）`,
                      values: [metric.fy2023, metric.fy2024],
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
            text: metric
              ? `来源：Apple FY2024 Form 10-K，印刷第 ${metric.annualPrintedPage} 页；同比为独立计算`
              : '来源：Apple FY2024 Form 10-K；待财务审阅',
            fontSize: 9,
          },
        ],
      }
    }),
  }
  assertDeckMatchesPresentationPlan(deck, parsedPlan)
  for (const slide of deck.slides) {
    const length = presentationSlideSourceLabels(slide, deck.claims).join('；').length
    if (length > 500) throw new Error(`source footer overflow: ${slide.id} (${length})`)
  }
  for (let index = 2; index <= 4; index++) {
    const chart = deck.slides[index].elements.find((element) => element.kind === 'chart')
    const report = checkPresentationChartData(parsedPlan, plan.slides[index].id, [chart])
    if (report.charts.some((item) => item.findings.length))
      throw new Error(
        `declared chart data invalid: ${plan.slides[index].id} ${JSON.stringify(report.charts)}`,
      )
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
  await writeFile(join(root, 'p0-07-reference.pptx'), result.bytes)
  console.log(
    `P0-07 source-backed reference: ${result.bytes.length} bytes, ${deck.slides.length} pages`,
  )
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
