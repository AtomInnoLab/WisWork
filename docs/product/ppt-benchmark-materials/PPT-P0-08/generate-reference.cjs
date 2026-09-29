const { readFile, writeFile } = require('node:fs/promises')
const { join } = require('node:path')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')

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
  const deck = {
    version: 1,
    id: 'p0-08-cross-company-comparability-reference',
    title: 'Apple 与 Toyota 跨公司可比性参考稿',
    style: {
      fontFace: 'Noto Sans CJK SC',
      background: 'FFFFFF',
      textColor: '173248',
      accentColor: '087D83',
    },
    assets: [],
    claims: [
      {
        id: 'apple-sales',
        text: `Apple 净销售额 ${apple.valueMillions} 百万美元`,
        source: apple.source,
        locator: `PDF 第 ${apple.reportPdfPage} 页`,
      },
      {
        id: 'toyota-sales',
        text: `Toyota 销售及金融服务收入 ${toyota.valueMillions} 百万日元`,
        source: toyota.valueSource,
        locator: `PDF 第 ${toyota.valueSourcePdfPage} 页`,
      },
    ],
    slides: titles.map((title, index) => {
      const company = index === 3 ? apple : index === 4 ? toyota : null
      return {
        id: `p${String(index + 1).padStart(2, '0')}`,
        title,
        claimIds: index === 3 ? ['apple-sales'] : index === 4 ? ['toyota-sales'] : [],
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
            y: 6.95,
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
  const result = await compilePresentationDeck(deck)
  await writeFile(join(root, 'p0-08-reference.pptx'), result.bytes)
  console.log(`P0-08 reference: ${result.bytes.length} bytes, ${deck.slides.length} pages`)
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
