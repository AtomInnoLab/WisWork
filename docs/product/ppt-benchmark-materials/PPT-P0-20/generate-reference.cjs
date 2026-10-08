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

async function main() {
  const root = __dirname
  const template = join(root, '../PPT-P0-01')
  const scenario = JSON.parse(await readFile(join(root, 'scenario.json'), 'utf8'))
  if (
    scenario.caseId !== 'PPT-P0-20' ||
    scenario.pages.length !== 8 ||
    scenario.faultSchedule.length !== 3
  )
    throw new Error('P0-20 scenario changed')
  const plan = JSON.parse(await readFile(join(template, 'reference-plan.json'), 'utf8'))
  const deck = JSON.parse(await readFile(join(template, 'reference-deck.json'), 'utf8'))
  const article = await readFile(join(root, scenario.documentInputs[0]))
  const checklist = await readFile(join(root, scenario.documentInputs[1]))
  if (plan.sources.some((source) => source.snapshotAttachmentId !== sha(article)))
    throw new Error('P0-20 article differs from source-backed template')
  const { pdfToPages } = await import('../../../../packages/file-parse/src/pdf.ts')
  const checklistPages = await pdfToPages(checklist)
  if (checklistPages.length !== 1) throw new Error('P0-20 checklist page count changed')
  const checklistExcerpt = checklistPages[0].trim().slice(0, 400)
  if (!checklistExcerpt) throw new Error('P0-20 checklist text missing')
  plan.projectId = deck.id = 'p0-20-fault-recovery-reference'
  plan.title = deck.title = '计算可复现性研究：瞬时错误与诚实交付候选'
  plan.brief.objective = '用两份冻结 PDF 制作八页原生稿，逐页保留来源与恢复边界'
  plan.brief.requiredContent = [
    '14/12 样本边界',
    '六项清单',
    'p=0.318 未达显著',
    'Table 1 原生图表',
    '三处故障后诚实交付',
  ]
  plan.brief.constraints.push('瞬时错误、截图失败与用户结束不等于已完成验收')
  plan.sources.push({
    id: 'supplement-checklist',
    title: 'Deardorff 2020 S1 六项清单 PDF',
    uri: 'doi:10.1371/journal.pone.0230697.s001',
    snapshotAttachmentId: sha(checklist),
    locator: 'PDF 第 1 页',
    excerpt: checklistExcerpt,
  })
  plan.claims.push({
    id: 'checklist-design',
    statement: '六项清单的项目见论文 S1 补充材料；它不是参与者原始分数表。',
    type: 'fact',
    sourceIds: ['supplement-checklist'],
    confidence: 'low',
    reviewStatus: 'needs_review',
  })
  deck.claims.push({
    id: 'checklist-design',
    text: plan.claims.at(-1).statement,
    source: 'doi:10.1371/journal.pone.0230697.s001',
    locator: 'PDF 第 1 页',
  })
  const bodies = [
    'Deardorff（2020）生物医学研究人员编程工作坊研究；本稿为待科研审阅候选。',
    '研究问题聚焦计算工作流可复现性；观察性结果不能解释为工作坊因果效果。',
    '论文采用工作坊前后访谈与六项清单；前访谈 14 人，三个月后完成后访谈 12 人。S1 是空白清单。',
    '研究路径为招募与前测、编程工作坊、三个月后访谈；两时点分母不同。',
    '六项清单均值由 1.6/6 到 2.2/6，论文报告 p=0.318，未达统计显著。',
    'Table 1 开源软件使用计数前测 7/14、三个月后 10/12；图中保留不同分母。',
    '小样本、招募或应答偏差、单人编码与统计功效不足限制解释。',
    '只交付已核验的页面和回执；故障或结束时不把部分成果称为完成。',
  ]
  for (const [index, page] of scenario.pages.entries()) {
    const slide = deck.slides[index]
    const planSlide = plan.slides[index]
    slide.title = planSlide.title = page.purpose
    slide.elements.find((element) => element.id === 'title').text = page.purpose
    slide.elements.find((element) => element.id === 'body').text = bodies[index]
    if (index === 2) {
      slide.claimIds.push('checklist-design')
      planSlide.claimIds.push('checklist-design')
    }
    if (index === 3) {
      slide.claimIds.push('sample')
      planSlide.claimIds.push('sample')
    }
    slide.notes = '候选研究稿；来源和故障后状态需逐页核对，真实 PowerPoint 验收待完成。'
  }
  const parsed = parsePresentationPlan(plan)
  assertDeckMatchesPresentationPlan(deck, parsed)
  const chart = deck.slides[5].elements.find((element) => element.kind === 'chart')
  if (
    !chart ||
    checkPresentationChartData(parsed, deck.slides[5].id, [chart]).charts.some(
      (item) => item.findings.length,
    )
  )
    throw new Error('P0-20 Table 1 chart mismatch')
  const compiled = await compilePresentationDeck(deck)
  const format = (value) =>
    prettier.format(JSON.stringify(value), { parser: 'json', printWidth: 100 })
  await writeFile(join(root, 'reference-plan.json'), await format(plan))
  await writeFile(join(root, 'reference-deck.json'), await format(deck))
  await writeFile(join(root, 'wiswork-generated-candidate.pptx'), compiled.bytes)
  const names = [
    ...scenario.documentInputs,
    'scenario.json',
    'reference-plan.json',
    'reference-deck.json',
    'wiswork-generated-candidate.pptx',
  ]
  await writeFile(
    join(root, 'SHA256SUMS'),
    (
      await Promise.all(
        names.map(async (name) => `${sha(await readFile(join(root, name)))}  ${name}`),
      )
    ).join('\n') + '\n',
  )
  console.log('P0-20 source-bound eight-page candidate generated')
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
