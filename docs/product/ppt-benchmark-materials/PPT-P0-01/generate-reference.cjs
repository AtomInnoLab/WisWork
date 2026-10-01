const { readFile, writeFile } = require('node:fs/promises')
const { join } = require('node:path')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')

async function main() {
  const root = __dirname
  const basis = JSON.parse(await readFile(join(root, 'basis.json'), 'utf8'))
  const titles = [
    '科研工作流可复现性：编程培训观察',
    '问题与研究范围',
    '访谈设计与样本流失',
    '六项清单：测量什么',
    '清单均分：不可夸大变化',
    'Table 1：开源工具使用',
    '定性观察与研究局限',
    '结论与待核验事项',
  ]
  const bodies = [
    'Deardorff（2020）基于 UCSF 生物医学研究人员的编程培训访谈与清单；这是待专业审阅的候选参考稿。',
    '问题：入门编程培训后，研究人员的可复现工作流实践出现了哪些自述变化？研究对象和结论限于这项小样本观察。',
    `培训前访谈 ${basis.sample.before} 人，三个月后 ${basis.sample.after} 人完成后访谈。分母不同，不能把两次计数直接视为同一批人的个体转变。`,
    'S1 Checklist 是六项可复现实践的空白测量表，不提供参与者个体原始分数。报告只使用论文已发表的汇总值。',
    `六项清单均分从 ${basis.score.before}/${basis.score.maximum} 到 ${basis.score.after}/${basis.score.maximum}；论文报告差异未达统计显著（p = ${basis.score.p}）。这不是培训因果效果的证明。`,
    `Table 1 中“使用开源软件”为培训前 ${basis.table1OpenSource.before}/${basis.table1OpenSource.beforeDenominator}、三个月后 ${basis.table1OpenSource.after}/${basis.table1OpenSource.afterDenominator}。图中展示计数，横轴同时标出不同分母。`,
    '访谈提示受访者对“分享代码”的理解可能改变。作者列出小样本、招募/应答偏差、单人编码与定量分析功效不足等局限。',
    '可以报告受访者自述与清单汇总的变化；不能声称已证明普遍或因果效果。仍需科研审阅、真实 PowerPoint 的图表编辑与保存重开验收。',
  ]
  const footer = [
    '来源：Deardorff 2020，PLOS ONE；候选参考稿',
    '来源：论文 PDF 第 1、3 页；仅限本研究样本',
    '来源：论文 PDF 第 3、5 页；前后测分母不同',
    '来源：S1 Checklist PDF 第 1 页；论文 PDF 第 5 页',
    '来源：论文 PDF 第 5 页；p = 0.318',
    '来源：论文 PDF 第 5 页 Table 1；前测 n=14、后测 n=12',
    '来源：论文 PDF 第 5、9 页；局限见第 9 页',
    '原件已冻结；科研审阅与 PowerPoint 宿主验收待完成',
  ]
  const source = 'Deardorff 2020，PLOS ONE e0230697'
  const claims = [
    { id: 'sample', text: '培训前 14 人、三个月后 12 人', source, locator: 'PDF 第 5 页' },
    { id: 'score', text: '均分 1.6/6 到 2.2/6；p=0.318', source, locator: 'PDF 第 5 页' },
    {
      id: 'open-source',
      text: '开源软件使用 7/14 到 10/12',
      source,
      locator: 'PDF 第 5 页 Table 1',
    },
    {
      id: 'limitations',
      text: '小样本、选择/应答、单人编码和功效限制',
      source,
      locator: 'PDF 第 9 页',
    },
  ]
  const deck = {
    version: 1,
    id: 'p0-01-research-reference',
    title: '编程培训与生物医学工作流可复现性：候选科研汇报',
    style: {
      fontFace: 'Noto Sans CJK SC',
      background: 'FFFFFF',
      textColor: '173248',
      accentColor: '087D83',
    },
    assets: [],
    claims,
    slides: titles.map((title, index) => ({
      id: `p${String(index + 1).padStart(2, '0')}`,
      title,
      claimIds:
        index === 2
          ? ['sample']
          : index === 4
            ? ['score']
            : index === 5
              ? ['open-source']
              : index === 6
                ? ['limitations']
                : [],
      notes: '待科研专业审阅；所有主张仅限冻结原文，不代表真实 PowerPoint 宿主验收。',
      elements: [
        { kind: 'text', id: 'title', x: 0.65, y: 0.48, w: 12, h: 0.7, text: title, fontSize: 26 },
        {
          kind: 'text',
          id: 'body',
          x: 0.9,
          y: 1.65,
          w: 11.4,
          h: index === 5 ? 0.9 : 3.8,
          text: bodies[index],
          fontSize: 19,
        },
        ...(index === 5
          ? [
              {
                kind: 'chart',
                id: 'table1-open-source-native',
                x: 1.0,
                y: 2.85,
                w: 10.8,
                h: 3.25,
                chartType: 'bar',
                categories: [
                  `前测 n=${basis.table1OpenSource.beforeDenominator}`,
                  `三个月后 n=${basis.table1OpenSource.afterDenominator}`,
                ],
                series: [
                  {
                    name: '报告使用开源软件的人数',
                    values: [basis.table1OpenSource.before, basis.table1OpenSource.after],
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
          y: 6.63,
          w: 11.4,
          h: 0.25,
          text: footer[index],
          fontSize: 9,
        },
      ],
    })),
  }
  const result = await compilePresentationDeck(deck)
  await writeFile(join(root, 'p0-01-reference.pptx'), result.bytes)
  console.log(
    `P0-01 candidate reference: ${result.bytes.length} bytes, ${deck.slides.length} native slides`,
  )
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
