const { readFile, writeFile } = require('node:fs/promises')
const { join } = require('node:path')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')

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
  const deck = {
    version: 1,
    id: 'p0-09-hypothetical-reference',
    title: 'P0-09 情景分析参考稿：非 Apple 指引',
    style: {
      fontFace: 'Noto Sans CJK SC',
      background: 'FFFFFF',
      textColor: '183247',
      accentColor: '087D83',
    },
    assets: [],
    claims: [
      {
        id: 'historical-2024',
        text: '2024 财年净销售额 391,035、营业利润 123,216，单位百万美元',
        source: 'Apple Inc. 2024-10-31 Consolidated Financial Statements (Unaudited)',
        locator: '第 1 页；截至 2024-09-28 的十二个月',
      },
    ],
    slides: titles.map((title, index) => ({
      id: `p${String(index + 1).padStart(2, '0')}`,
      title,
      claimIds: index === 2 ? ['historical-2024'] : [],
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
          y: 6.95,
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
  const result = await compilePresentationDeck(deck)
  await writeFile(join(root, 'p0-09-reference.pptx'), result.bytes)
  console.log(`P0-09 reference: ${result.bytes.length} bytes, ${deck.slides.length} pages`)
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
