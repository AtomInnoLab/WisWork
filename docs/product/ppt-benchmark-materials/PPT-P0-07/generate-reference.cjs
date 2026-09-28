const { readFile, writeFile } = require('node:fs/promises')
const { join } = require('node:path')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')

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
    `2023：383,285；2024：391,035；同比 ${rows[0][4]}%。审计年报第 29 页。`,
    `2023：96,995；2024：93,736；同比 ${rows[1][4]}%，为下降。审计年报第 29 页。`,
    `2023：110,543；2024：118,254；同比 +${rows[2][4]}%。审计年报第 33 页。`,
    '同比 = (2024 值－2023 值) ÷ 2023 值 × 100%；四舍五入到百分比小数点后两位。',
    '净利润同比下降；官方 Q4 业绩表未经审计，不能替代 Form 10-K 审计意见。',
    '资料时点 2024-11-01；指标和同比待财务审阅，真实 PowerPoint 编辑与重开待验收。',
  ]
  const deck = {
    version: 1,
    id: 'p0-07-annual-performance-reference',
    title: 'Apple 2024 财年表现参考稿',
    style: {
      fontFace: 'Noto Sans CJK SC',
      background: 'FFFFFF',
      textColor: '173248',
      accentColor: '087D83',
    },
    assets: [],
    claims: dictionary.metrics.map((metric) => ({
      id: metric.id,
      text: `${metric.label}：2024 ${metric.fy2024}，2023 ${metric.fy2023}，百万美元`,
      source: 'Apple Inc. FY2024 Form 10-K, audited consolidated statements',
      locator: `印刷第 ${metric.annualPrintedPage} 页`,
    })),
    slides: titles.map((title, index) => {
      const metric = dictionary.metrics[index - 2]
      return {
        id: `p${String(index + 1).padStart(2, '0')}`,
        title,
        claimIds: metric ? [metric.id] : [],
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
            y: 6.95,
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
  const result = await compilePresentationDeck(deck)
  await writeFile(join(root, 'p0-07-reference.pptx'), result.bytes)
  console.log(`P0-07 reference: ${result.bytes.length} bytes, ${deck.slides.length} pages`)
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
