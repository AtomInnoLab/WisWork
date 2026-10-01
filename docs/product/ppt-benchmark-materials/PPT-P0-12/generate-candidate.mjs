import { createHash } from 'node:crypto'
import { readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = dirname(fileURLToPath(import.meta.url))
const sourceRoot = join(root, '../PPT-P0-01')
const readJson = async (path) => JSON.parse(await readFile(path, 'utf8'))
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')
const [originalPlan, originalDeck, brandKit, article, checklist] = await Promise.all([
  readJson(join(sourceRoot, 'reference-plan.json')),
  readJson(join(sourceRoot, 'reference-deck.json')),
  readJson(join(root, 'brand-kit.json')),
  readFile(join(root, 'deardorff-2020-article.pdf')),
  readFile(join(root, 'deardorff-2020-checklist.pdf')),
])
if (originalPlan.sources.some((source) => source.snapshotAttachmentId !== sha(article)))
  throw new Error('P0-01 article source changed')
const title = '编程培训与生物医学工作流可复现性：品牌汇报候选'
const projectId = 'p0-12-branded-research-candidate'
const style = {
  fontFace: 'WisWork Benchmark Display 2026',
  fontFallbacks: ['Noto Sans CJK SC'],
  background: 'FFFFFF',
  textColor: '102A43',
  accentColor: '007F86',
}
const plan = structuredClone(originalPlan)
const deck = structuredClone(originalDeck)
plan.projectId = projectId
deck.id = projectId
plan.title = title
deck.title = title
plan.style = style
deck.style = style
plan.brandKit = brandKit
plan.brief.objective = '按自制品牌模板和 Brand Kit 制作八页可编辑科研汇报，验证缺失字体的显式回退'
plan.brief.requiredContent.push('品牌布局组件', '缺失字体回退', '两份科研 PDF 来源')
plan.brief.constraints.push('品牌与科学解释待人工审阅', '真实 PowerPoint 宿主字形待验收')
const checklistSource = {
  id: 'checklist-scorecard',
  title: 'Deardorff 2020 S1 Checklist',
  uri: 'local:deardorff-2020-checklist.pdf',
  snapshotAttachmentId: sha(checklist),
  locator: 'PDF 第 1 页 / Reproducibility Score Card',
  excerpt: 'Reproducibility Score Card',
  asOf: '2020-07-01',
}
plan.sources.push(checklistSource)
plan.claims.push({
  id: 'checklist-six',
  statement: 'S1 Checklist 列出六项可复现实践，空白表本身不提供参与者个体原始分数。',
  type: 'fact',
  sourceIds: [checklistSource.id],
  confidence: 'low',
  reviewStatus: 'needs_review',
  asOf: '2020-07-01',
})
deck.claims.push({
  id: 'checklist-six',
  text: plan.claims.at(-1).statement,
  source: checklistSource.uri,
  locator: checklistSource.locator,
})
plan.slides[3].claimIds = ['checklist-six']
deck.slides[3].claimIds = ['checklist-six']
const contentPages = new Set(['p02', 'p03', 'p04', 'p05', 'p07'])
for (const [index, slide] of deck.slides.entries()) {
  const task = plan.slides[index]
  slide.elements.push({
    id: 'brand-bar',
    kind: 'shape',
    shape: 'rect',
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.16,
    fill: '007F86',
    lineColor: '007F86',
    role: 'decoration',
  })
  slide.elements.push({
    id: 'brand-label',
    kind: 'text',
    text: 'WW  /  RESEARCH',
    x: 0.62,
    y: 0.3,
    w: 3,
    h: 0.3,
    fontSize: 10,
    color: '007F86',
    role: 'decoration',
  })
  const titleElement = slide.elements.find((element) => element.id === 'title')
  const body = slide.elements.find((element) => element.id === 'body')
  if (slide.id === 'p01') {
    slide.elements.push({
      id: 'cover-rule',
      kind: 'shape',
      shape: 'rect',
      x: 0.62,
      y: 1.18,
      w: 0.16,
      h: 4.6,
      fill: '007F86',
      lineColor: '007F86',
      role: 'decoration',
    })
    task.layoutComponentId = 'benchmark-cover'
    titleElement.id = 'cover-title'
    Object.assign(titleElement, { x: 1.03, y: 1.65, w: 10.8, h: 1, fontSize: 29, bold: true })
    body.id = 'cover-subtitle'
    Object.assign(body, {
      text: 'PLOS ONE 原文与 S1 Checklist · 品牌模板约束 · 科学审阅待完成',
      x: 1.03,
      y: 2.85,
      w: 10.8,
      h: 0.5,
      fontSize: 18,
    })
  } else if (contentPages.has(slide.id)) {
    task.layoutComponentId = 'benchmark-content'
    titleElement.id = 'content-title'
    Object.assign(titleElement, { x: 0.62, y: 0.94, w: 11.6, h: 0.7, fontSize: 26, bold: true })
    body.id = 'content-body'
    Object.assign(body, { x: 0.9, y: 2.17, w: 6.4, h: 3.8, fontSize: 18 })
    slide.elements.unshift({
      id: 'content-panel',
      kind: 'shape',
      shape: 'rect',
      x: 0.62,
      y: 1.9,
      w: 7.05,
      h: 4.55,
      fill: 'E8F5F3',
      lineColor: 'E8F5F3',
      role: 'background',
    })
    slide.elements.push({
      id: 'evidence-panel',
      kind: 'shape',
      shape: 'rect',
      x: 8.05,
      y: 1.9,
      w: 4.63,
      h: 4.55,
      fill: 'FFFFFF',
      lineColor: 'D8E2E9',
      role: 'background',
    })
    slide.elements.push({
      id: 'evidence-label',
      kind: 'text',
      text: '原文证据',
      x: 8.38,
      y: 2.35,
      w: 3.95,
      h: 0.42,
      fontSize: 18,
      color: '007F86',
      bold: true,
    })
    slide.elements.push({
      id: 'evidence-key',
      kind: 'text',
      text:
        slide.id === 'p03'
          ? '14 → 12'
          : slide.id === 'p04'
            ? '6 项实践'
            : slide.id === 'p05'
              ? 'p = 0.318'
              : slide.id === 'p07'
                ? '小样本'
                : '观察范围',
      x: 8.38,
      y: 3.25,
      w: 3.75,
      h: 0.9,
      fontSize: 29,
      color: '007F86',
      bold: true,
    })
    slide.elements.push({
      id: 'content-source',
      kind: 'text',
      text:
        slide.id === 'p04'
          ? 'S1 Checklist · PDF 第 1 页\n六项空白量表；无个人原始分数'
          : slide.id === 'p05'
            ? '论文 PDF 第 5 页\n均分 1.6/6 → 2.2/6；p=0.318'
            : slide.id === 'p07'
              ? '论文 PDF 第 9 页\n作者列出小样本与分析局限'
              : slide.id === 'p03'
                ? '论文 PDF 第 5 页\n前测 14 人，后访 12 人'
                : 'PLOS ONE 原文\n只讨论本项小样本观察',
      x: 8.38,
      y: 5.45,
      w: 3.95,
      h: 0.55,
      fontSize: 11,
    })
  } else {
    Object.assign(titleElement, { x: 0.62, y: 0.94, w: 11.6, h: 0.7, fontSize: 26, bold: true })
  }
  const footer = slide.elements.find((element) => element.id === 'footer')
  footer.text = 'Deardorff 2020 · PLOS ONE · CC BY 4.0 · 研究与品牌审阅待完成'
  footer.x = 0.9
  footer.y = 6.65
  footer.w = 11.4
  footer.h = 0.24
  footer.fontSize = 9
  if (slide.id === 'p06') {
    slide.elements.unshift({
      id: 'chart-panel',
      kind: 'shape',
      shape: 'rect',
      x: 0.62,
      y: 1.9,
      w: 12.05,
      h: 4.55,
      fill: 'FFFFFF',
      lineColor: 'D8E2E9',
      role: 'background',
    })
    Object.assign(body, { y: 1.82, h: 0.72, fontSize: 17 })
    const chart = slide.elements.find((element) => element.kind === 'chart')
    Object.assign(chart, { x: 1.05, y: 2.85, w: 10.9, h: 3.2 })
  }
  if (slide.id === 'p08') {
    Object.assign(body, { x: 0.9, y: 2.05, w: 11.4, h: 0.9, fontSize: 20 })
    for (const [id, label, value, x] of [
      ['conclusion', '可陈述', '自述与清单汇总出现变化；研究只支持受限描述。', 0.9],
      ['review', '待核验', '科学解释、真实宿主字体、图表编辑与保存重开。', 6.75],
    ]) {
      slide.elements.push({
        id: `${id}-panel`,
        kind: 'shape',
        shape: 'roundRect',
        x,
        y: 3.45,
        w: 5.55,
        h: 2.2,
        fill: 'E8F5F3',
        lineColor: 'D8E2E9',
        role: 'background',
      })
      slide.elements.push({
        id: `${id}-label`,
        kind: 'text',
        text: label,
        x: x + 0.2,
        y: 3.72,
        w: 5.1,
        h: 0.4,
        fontSize: 17,
        color: '007F86',
        bold: true,
      })
      slide.elements.push({
        id: `${id}-body`,
        kind: 'text',
        text: value,
        x: x + 0.2,
        y: 4.3,
        w: 5.1,
        h: 1.03,
        fontSize: 16,
      })
    }
  }
}
await writeFile(join(root, 'reference-plan.json'), `${JSON.stringify(plan, null, 2)}\n`)
await writeFile(join(root, 'reference-deck.json'), `${JSON.stringify(deck, null, 2)}\n`)
