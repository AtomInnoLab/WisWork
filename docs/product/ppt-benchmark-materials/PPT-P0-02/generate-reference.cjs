const { createHash } = require('node:crypto')
const { readFile, writeFile } = require('node:fs/promises')
const { join } = require('node:path')
const prettier = require('prettier')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')
const {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
} = require('@wiswork/pptx-engine/presentation-plan')

const papers = [
  {
    id: 'helps',
    file: 'helps-2014-white-noise.pdf',
    uri: 'doi:10.1371/journal.pone.0112768',
    title: 'Helps 2014，PDF 第 1 页 Results',
    first: 'Results: There were different effects of WN on performance',
    last: 'normal-attentive children’s performance was unaffected by WN exposure.',
    claim:
      '教师评定的 8–10 岁儿童中，中等白噪声与高注意力组表现变差、低注意力组部分执行功能任务改善有关；常注意力组未见影响。',
  },
  {
    id: 'han',
    file: 'han-2013-speech-noise.pdf',
    uri: 'doi:10.1371/journal.pone.0076261',
    title: 'Han 2013，PDF 第 1 页 Abstract',
    first: 'We found that the memory performance in low arousal condition',
    last: 'between high arousal and silent conditions.',
    claim:
      'N-back 任务中低唤醒语音噪声条件较静音和高唤醒条件反应更快；高唤醒与静音间未见显著差异。',
  },
  {
    id: 'mohanathasan',
    file: 'mohanathasan-2025-conversation-noise.pdf',
    uri: 'doi:10.1371/journal.pone.0318821',
    title: 'Mohanathasan 2025，PDF 第 1 页 Abstract',
    first: 'In Experiments 1 and 2, short-term memory of running speech',
    last: 'while performance in the vibrotactile secondary task was unaffected.',
    claim:
      '两人对话任务中，+10 dB 信噪比的轻度噪声未见短时记忆影响；−3 dB 条件下主听力任务表现受损。',
  },
]

function excerptBetween(text, first, last) {
  const start = text.indexOf(first)
  const end = text.indexOf(last, start)
  if (start < 0 || end < 0) throw new Error(`original PDF anchor missing: ${first}`)
  return text.slice(start, end + last.length)
}

async function main() {
  const root = __dirname
  const { pdfToPages } = await import('../../../../packages/file-parse/src/pdf.ts')
  const sources = []
  const claims = []
  for (const paper of papers) {
    const bytes = await readFile(join(root, paper.file))
    const pages = await pdfToPages(bytes)
    const excerpt = excerptBetween(pages[0], paper.first, paper.last)
    sources.push({
      id: paper.id,
      title: paper.title,
      uri: paper.uri,
      snapshotAttachmentId: createHash('sha256').update(bytes).digest('hex'),
      locator: '第 1 页',
      excerpt,
    })
    claims.push({
      id: paper.id,
      text: paper.claim,
      source: paper.uri,
      locator: '第 1 页',
    })
  }
  const rows = [
    {
      title: '背景噪声会帮助认知任务吗？',
      body: '三项原始研究的参与者、噪声、任务与指标不同。本稿为待领域审阅的候选证据比较，不给通用学习或医疗建议。',
      claimIds: [],
      footer: '来源：Helps 2014；Han 2013；Mohanathasan 2025，原始论文 PDF',
    },
    {
      title: '先定义可比边界',
      body: '比较的是各研究报告的方向及条件，不合并效果量。白噪声、语音噪声和宽带背景噪声不是同一种干预；儿童分组、N-back 和对话记忆也不是同一种任务。',
      claimIds: [],
      footer: '三篇论文 PDF 第 1 页 Abstract；不能做跨研究剂量曲线',
    },
    {
      title: 'Helps：组别决定白噪声方向',
      body: '人群：教师评定的 8–10 岁儿童，三种注意力组。条件：不同白噪声水平。指标：记忆和执行功能任务。中等白噪声下，低注意力组部分执行功能任务改善，高注意力组表现变差，常注意力组未见影响。',
      claimIds: ['helps'],
      footer: 'Helps 2014，PDF 第 1 页 Results；注意力分组不等于临床诊断',
    },
    {
      title: 'Han：低唤醒语音噪声条件',
      body: '任务：N-back 工作记忆。比较：低唤醒语音噪声、静音、高唤醒语音噪声。低唤醒条件反应较快；高唤醒与静音之间未见显著差异。语音唤醒条件不能直接代换为白噪声强度。',
      claimIds: ['han'],
      footer: 'Han 2013，PDF 第 1 页 Abstract；反应时间与 ERP 均为研究指标',
    },
    {
      title: 'Mohanathasan：轻度与中度条件',
      body: '任务：听取两人对话后回忆，并执行第二任务。轻度背景噪声为 +10 dB 信噪比，实验 1/2 未见短时记忆影响；实验 3 的 −3 dB 条件使主听力任务表现受损。信噪比不是绝对音量。',
      claimIds: ['mohanathasan'],
      footer: 'Mohanathasan 2025，PDF 第 1 页 Abstract；实验条件不可互换',
    },
    {
      title: '同页并列：有利、未见显著差异、不利',
      body: '有利｜Helps：低注意力儿童的部分执行功能任务；Han：低唤醒语音噪声下 N-back 反应。\n未见显著差异｜Helps：常注意力组；Han：高唤醒与静音；Mohanathasan：+10 dB 对话短时记忆。\n不利｜Helps：高注意力组；Mohanathasan：−3 dB 主听力任务。',
      claimIds: ['helps', 'han', 'mohanathasan'],
      footer: '三篇论文 PDF 第 1 页；“未见显著差异”不等于证明无效',
    },
    {
      title: '可比性矩阵：四个关键差异',
      body: '人群｜儿童注意力分组 / N-back 受试者 / 对话听者。\n噪声｜白噪声水平 / 语音情绪唤醒 / 宽带噪声信噪比。\n任务｜记忆与执行功能 / N-back / 对话回忆与双任务。\n指标｜组别任务表现 / 反应时间与 ERP / 主听力及第二任务。',
      claimIds: ['helps', 'han', 'mohanathasan'],
      footer: '三篇论文 PDF 第 1 页；不可直接汇总为单一总体效应',
    },
    {
      title: '结论与审阅边界',
      body: '证据方向随人群、噪声和任务而变；保留有利、未见显著差异与不利结果。Helps 披露作者与白噪声产品开发公司的关联。科研审阅仍须核对原文、冲突措辞和每页来源；真实 PowerPoint 保存重开尚未验收。',
      claimIds: ['helps', 'han', 'mohanathasan'],
      footer: 'Helps PDF 第 1 页 Competing Interests；三项研究均不能支持通用建议',
    },
  ]
  const deck = {
    version: 1,
    id: 'p0-02-noise-evidence-reference',
    title: '背景噪声与认知任务：三项研究候选比较',
    style: {
      fontFace: 'Noto Sans CJK SC',
      background: 'FFFFFF',
      textColor: '173248',
      accentColor: '087D83',
    },
    assets: [],
    claims,
    slides: rows.map((row, index) => ({
      id: `p${String(index + 1).padStart(2, '0')}`,
      title: row.title,
      claimIds: row.claimIds,
      notes: '候选比较，待认知/心理学领域审阅；不得当作通用建议。',
      elements: [
        {
          kind: 'text',
          id: 'title',
          x: 0.65,
          y: 0.48,
          w: 12,
          h: 0.7,
          text: row.title,
          fontSize: 26,
        },
        ...row.body.split('\n').map((text, line) => ({
          kind: 'text',
          id: line ? `body-${line + 1}` : 'body',
          x: 0.9,
          y: 1.55 + line * 1.2,
          w: 11.4,
          h: row.body.includes('\n') ? 1.1 : 4.75,
          text,
          fontSize: index === 5 || index === 6 ? 17 : 19,
        })),
        {
          kind: 'text',
          id: 'footer',
          role: 'decoration',
          x: 0.9,
          y: 6.63,
          w: 11.4,
          h: 0.25,
          text: row.footer,
          fontSize: 9,
        },
      ],
    })),
  }
  const plan = {
    version: 1,
    projectId: deck.id,
    title: deck.title,
    brief: {
      objective: '基于三篇冻结原文制作八页科研证据比较，保留方向冲突与可比性边界',
      audience: '科研专业审阅者',
      language: 'zh-CN',
      minutes: 8,
      requiredContent: ['三项研究的条件与结果', '同页三方向证据', '可比性矩阵', '利益冲突与局限'],
      constraints: ['不得合并效果量', '科学解释待专业审阅', '真实 PowerPoint 保存重开尚未验收'],
    },
    sources,
    claims: claims.map((claim) => ({
      id: claim.id,
      statement: claim.text,
      type: 'fact',
      sourceIds: [claim.id],
      confidence: 'low',
      reviewStatus: 'needs_review',
    })),
    style: deck.style,
    slides: deck.slides.map((slide, index) => ({
      id: slide.id,
      title: slide.title,
      purpose: '并列展示研究条件、结果方向和不可比边界',
      claimIds: slide.claimIds,
      layout: index === 0 ? 'cover' : index === 7 ? 'summary' : 'content',
      requiredAssets: [],
      acceptanceCriteria: ['原生可编辑文字', '保留三方向证据及来源', '不合并跨研究效果量'],
    })),
  }
  assertDeckMatchesPresentationPlan(deck, parsePresentationPlan(plan))
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
  await writeFile(join(root, 'p0-02-reference.pptx'), result.bytes)
  console.log(
    `P0-02 candidate reference: ${result.bytes.length} bytes, ${deck.slides.length} native slides`,
  )
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
