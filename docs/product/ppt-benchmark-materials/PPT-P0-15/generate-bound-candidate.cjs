const { readFile, writeFile } = require('node:fs/promises')
const { join } = require('node:path')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')

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
  const deck = {
    version: 1,
    id: 'p0-15-generated-candidate',
    title: '计算可复现性工作坊：同内容生成绑定稿候选',
    style: {
      fontFace: 'Noto Sans CJK SC',
      background: 'FFFFFF',
      textColor: '122B43',
      accentColor: '008086',
    },
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
    claims: [
      {
        id: 'source-1',
        text: '论文研究事实仍待人工审阅',
        source: 'Deardorff et al., PLOS ONE 2020, doi:10.1371/journal.pone.0230697',
        locator: '第 1 页',
      },
    ],
    slides: titles.map((title, index) => ({
      id: `p${String(index + 1).padStart(2, '0')}`,
      title,
      claimIds: ['source-1'],
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
          y: 6.95,
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
  const compiled = await compilePresentationDeck(deck)
  await writeFile(join(root, 'wiswork-generated-candidate.pptx'), compiled.bytes)
  console.log(
    `PPT-P0-15 generated candidate: ${compiled.bytes.length} bytes, ${deck.slides.length} pages`,
  )
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
