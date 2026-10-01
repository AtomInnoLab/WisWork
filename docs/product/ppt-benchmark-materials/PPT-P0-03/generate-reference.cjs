const { createHash } = require('node:crypto')
const { readFile, writeFile } = require('node:fs/promises')
const { join } = require('node:path')
const prettier = require('prettier')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')
const {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
  presentationPlanClaims,
} = require('@wiswork/pptx-engine/presentation-plan')

const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')
function between(text, first, last) {
  const start = text.indexOf(first)
  const end = text.indexOf(last, start)
  if (start < 0 || end < 0) throw new Error(`source anchor missing: ${first}`)
  return text.slice(start, end + last.length)
}

async function main() {
  const root = __dirname
  const { pdfToPages } = await import('../../../../packages/file-parse/src/pdf.ts')
  const articleBytes = await readFile(join(root, 'hess-peterson-2015-article.pdf'))
  const dictionaryBytes = await readFile(join(root, 'hess-peterson-2015-dictionary.pdf'))
  const csvBytes = await readFile(join(root, 'treatment-outcomes.csv'))
  const article = await pdfToPages(articleBytes)
  const dictionary = await pdfToPages(dictionaryBytes)
  const csvLines = csvBytes.toString('utf8').trim().split(/\r?\n/)
  if (
    csvLines[0] !== 'treatment,n,permitted_2_agree,safe_2_agree,permitted_4_agree,safe_4_agree' ||
    csvLines.length !== 5
  )
    throw new Error('frozen aggregate CSV shape changed')
  const treatments = csvLines.slice(1).map((line) => {
    const [id, ...numbers] = line.split(',')
    const [n, permitted, safe] = numbers.map(Number)
    if (
      !id ||
      numbers.length !== 5 ||
      numbers.some((v) => !/^\d+$/.test(v)) ||
      !Number.isSafeInteger(n) ||
      permitted > n ||
      safe > n
    )
      throw new Error('frozen aggregate CSV row invalid')
    return { id, n, permitted, safe, line }
  })
  if (treatments.reduce((sum, row) => sum + row.n, 0) !== 1824)
    throw new Error('frozen aggregate denominator changed')
  const sources = [
    {
      id: 'article-sample',
      title: 'Hess 与 Peterson 2015，PDF 第 8 页 Data Analysis',
      uri: 'doi:10.1371/journal.pone.0136973',
      snapshotAttachmentId: sha(articleBytes),
      locator: '第 8 页',
      excerpt: between(article[7], 'We obtained 1,978 responses', 'remaining 1,824 responses'),
    },
    {
      id: 'article-scope',
      title: 'Hess 与 Peterson 2015，PDF 第 1 页 Abstract',
      uri: 'doi:10.1371/journal.pone.0136973',
      snapshotAttachmentId: sha(articleBytes),
      locator: '第 1 页',
      excerpt: between(article[0], 'We administered a web-based survey', 'perceptions of safety.'),
    },
    {
      id: 'dictionary-fields',
      title: 'Hess 与 Peterson 2015，S1 Data PDF 第 1 页',
      uri: 'doi:10.1371/journal.pone.0136973.s001',
      snapshotAttachmentId: sha(dictionaryBytes),
      locator: '第 1 页',
      excerpt: between(
        dictionary[0],
        'Permitted2: [2-lane roadway]',
        '0_Disagree; 1_Agree\nBikeMoveRight4:',
      ),
    },
    ...treatments.map((row, index) => ({
      id: `csv-row-${index + 1}`,
      title: `汇总 CSV 第 ${index + 2} 行：${row.id}`,
      uri: 'local:treatment-outcomes.csv',
      snapshotAttachmentId: sha(csvBytes),
      locator: `CSV 第 ${index + 2} 行`,
      excerpt: row.line,
    })),
  ]
  const calculationClaims = treatments.flatMap((row, index) =>
    ['permitted', 'safe'].map((field) => {
      const agree = row[field]
      const percent = Number(((agree / row.n) * 100).toFixed(2))
      return {
        id: `${field}-${index + 1}`,
        statement: `${row.id}：${field === 'permitted' ? 'Permitted2' : 'Safe2'} 同意 ${agree}/${row.n} = ${percent}%`,
        type: 'calculation',
        sourceIds: [`csv-row-${index + 1}`],
        confidence: 'low',
        reviewStatus: 'needs_review',
        calculation: {
          formula: 'round(agree/n*100,2)',
          inputs: [`${field}_agree`, 'n'],
          unit: '%',
          reproduction: {
            bindings: [
              { name: 'agree', inputIndex: 0, value: agree, sourceId: `csv-row-${index + 1}` },
              { name: 'n', inputIndex: 1, value: row.n, sourceId: `csv-row-${index + 1}` },
            ],
            expected: percent,
          },
        },
      }
    }),
  )
  const claims = [
    {
      id: 'sample',
      statement: '论文收集 1,978 份问卷，排除 154 份非美国答卷后分析 1,824 份。',
      type: 'fact',
      sourceIds: ['article-sample'],
      confidence: 'low',
      reviewStatus: 'needs_review',
    },
    {
      id: 'scope',
      statement: '网络问卷考察标志理解与主观安全感，并未直接测量交通事故变化。',
      type: 'fact',
      sourceIds: ['article-scope'],
      confidence: 'low',
      reviewStatus: 'needs_review',
    },
    {
      id: 'dictionary',
      statement: 'Permitted2 与 Safe2 分别是两车道场景的通行许可理解和主观安全感同意项。',
      type: 'fact',
      sourceIds: ['dictionary-fields'],
      confidence: 'low',
      reviewStatus: 'needs_review',
    },
    ...calculationClaims,
  ]
  const chartClaimIds = calculationClaims.map((claim) => claim.id)
  const rows = [
    [
      '美国道路标志：理解与安全感',
      '来自论文、S1 字典与去个体化汇总 CSV 的八页候选研究汇报；交通研究审阅仍待完成。',
      [],
      '来源：Hess 与 Peterson 2015；数据为候选复算',
    ],
    [
      '研究问题与测量边界',
      '网络问卷考察美国道路标志是否传达骑车人可使用整条车道的规则，以及受访者的主观安全感；它没有直接观测事故率。',
      ['scope'],
      '论文 PDF 第 1 页 Abstract；不外推为事故减少',
    ],
    [
      '样本筛选：不能混用分母',
      '收到 1,978 份问卷，排除 154 份非美国答卷后分析 1,824 份。四组汇总人数为 489、422、454、459，总和为 1,824。',
      ['sample'],
      '论文 PDF 第 8 页 Data Analysis；汇总 CSV 的 n 列',
    ],
    [
      '四种标志与两车道指标',
      '四组：无标志、Share the Road、Shared Lane Markings、Bicycles May Use Full Lane。Permitted2 是对车道中央骑行许可的同意，Safe2 是对安全感的同意；两项均为主观问卷回答。',
      ['dictionary'],
      'S1 Data 字典 PDF 第 1 页；按组各有自己的分母',
    ],
    [
      '原生图表：两项同意比例',
      '计算式：各组同意人数 ÷ 该组 n × 100%，四舍五入至两位小数。图表横轴写明 n；柱形为原生可编辑数据。',
      chartClaimIds,
      '汇总 CSV：四组 Permitted2、Safe2 计数；不代表回归调整结果',
    ],
    [
      '如何阅读组间差异',
      '两项指标均以各组 n 为分母。候选图只给未经调整的描述性比例；不能据此声称已独立复现论文回归、置信区间或因果效应。',
      chartClaimIds,
      '汇总 CSV 与 S1 字典；统计解释待交通研究审阅',
    ],
    [
      '局限：主观问卷不是事故实验',
      '美国受访者、Twitter 招募、道路标志理解与感知安全的调查范围，需要与真实道路行为和事故数据分开。论文建议进一步做虚拟现实与道路实验。',
      ['scope'],
      '论文 PDF 第 1 页 Abstract；地域与方法边界',
    ],
    [
      '结论与待核验事项',
      '可报告各组描述性同意比例和论文的研究问题；不得将主观安全感写成事故率下降。交通研究审阅、真实 PowerPoint 图表数据编辑与保存重开仍待完成。',
      ['scope'],
      '三份冻结材料；候选稿尚未通过专业和宿主验收',
    ],
  ]
  const categories = treatments.map((row) => `${row.id} n=${row.n}`)
  const series = ['permitted', 'safe'].map((field) => ({
    name: field === 'permitted' ? 'Permitted2 同意率' : 'Safe2 同意率',
    values: treatments.map((row) => Number(((row[field] / row.n) * 100).toFixed(2))),
  }))
  const style = {
    fontFace: 'Noto Sans CJK SC',
    background: 'FFFFFF',
    textColor: '173248',
    accentColor: '087D83',
  }
  const projectId = 'p0-03-road-sign-reference'
  const plan = {
    version: 1,
    projectId,
    title: '美国道路标志：理解与安全感候选汇报',
    brief: {
      objective: '基于论文、字典和去个体化汇总 CSV 制作八页交通研究候选汇报',
      audience: '交通研究专业审阅者',
      language: 'zh-CN',
      minutes: 8,
      requiredContent: ['样本筛选', '四组标志', '两车道两项同意率原生图表', '数据口径和局限'],
      constraints: [
        '未经调整的汇总比例不得当作回归结果',
        '主观安全感不得写成事故率',
        '真实 PowerPoint 保存重开尚未验收',
      ],
    },
    sources,
    claims,
    style,
    slides: rows.map(([title, , claimIds], index) => ({
      id: `p${String(index + 1).padStart(2, '0')}`,
      title,
      purpose: '展示可复算汇总与研究边界',
      claimIds,
      layout: index === 0 ? 'cover' : index === 4 ? 'chart' : index === 7 ? 'summary' : 'content',
      requiredAssets: [],
      acceptanceCriteria: ['原生可编辑对象', '数据与来源可回读', '保留分母和局限'],
      ...(index === 4
        ? {
            chartData: [
              {
                elementId: 'two-lane-agreement',
                categories,
                unit: '%',
                series: series.map((item, seriesIndex) => ({
                  name: item.name,
                  points: item.values.map((value, pointIndex) => ({
                    value,
                    claimId: `${seriesIndex === 0 ? 'permitted' : 'safe'}-${pointIndex + 1}`,
                    basis: { kind: 'calculation' },
                  })),
                })),
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
    assets: [],
    claims: presentationPlanClaims(parsedPlan),
    slides: rows.map(([title, body, claimIds, footer], index) => ({
      id: plan.slides[index].id,
      title,
      claimIds,
      notes: '汇总比例为待专业审阅的描述性复算；真实 PowerPoint 宿主验收未完成。',
      elements: [
        { kind: 'text', id: 'title', x: 0.65, y: 0.48, w: 12, h: 0.7, text: title, fontSize: 26 },
        {
          kind: 'text',
          id: 'body',
          x: 0.9,
          y: 1.5,
          w: 11.4,
          h: index === 4 ? 0.9 : 4.75,
          text: body,
          fontSize: 19,
        },
        ...(index === 4
          ? [
              {
                kind: 'chart',
                id: 'two-lane-agreement',
                x: 0.95,
                y: 2.55,
                w: 11.35,
                h: 3.65,
                chartType: 'bar',
                categories,
                series,
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
          text: footer,
          fontSize: 9,
        },
      ],
    })),
  }
  assertDeckMatchesPresentationPlan(deck, parsedPlan)
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
  await writeFile(join(root, 'p0-03-reference.pptx'), result.bytes)
  console.log(
    `P0-03 candidate reference: ${result.bytes.length} bytes, ${deck.slides.length} native slides`,
  )
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
