const { readFile, writeFile } = require('node:fs/promises')
const { createHash } = require('node:crypto')
const { join } = require('node:path')
const prettier = require('prettier')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')
const {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
  presentationPlanClaims,
} = require('@wiswork/pptx-engine/presentation-plan')

function sliceOriginal(page, start, end) {
  const from = page.indexOf(start)
  const through = page.indexOf(end, from)
  if (from < 0 || through < 0) throw new Error(`official PDF anchor missing: ${start}`)
  return page.slice(from, through + end.length)
}

async function main() {
  const root = __dirname
  const manifest = JSON.parse(await readFile(join(root, 'materials-manifest.json')))
  const { pdfToPages } = await import('../../../../packages/file-parse/src/pdf.ts')
  const originals = new Map()
  for (const source of manifest.sources) {
    const bytes = await readFile(join(root, source.path))
    const hash = createHash('sha256').update(bytes).digest('hex')
    if (hash !== source.sha256 || bytes.length !== source.fileSize)
      throw new Error(`official original changed: ${source.id}`)
    originals.set(source.id, { source, pages: await pdfToPages(bytes) })
  }
  const sourceSpecs = [
    [
      'order-effective',
      'amendment',
      194,
      'December 1, 2023, and shall govern',
      'all proceedings then pending.',
    ],
    ['old-rule', 'old', 29, 'opinion or otherwise if:', 'ods to the facts of the case.'],
    [
      'new-opening',
      'historical',
      30,
      'opinion or otherwise if the proponent',
      'that it is more likely than not that:',
    ],
    [
      'new-d',
      'historical',
      31,
      '(d) the expert’s opinion reflects',
      'principles and methods to the facts of the case.',
    ],
    [
      'note-standard',
      'amendment',
      210,
      'Rule 702 has been amended in two respects:',
      'requirements set forth in the rule. See Rule 104(a).',
    ],
    [
      'note-weight',
      'amendment',
      211,
      'Some challenges to expert testimony',
      'matters of weight rather than admissibility',
    ],
    [
      'note-scope',
      'amendment',
      212,
      'that each expert opinion must stay',
      'basis and methodology.',
    ],
    [
      'note-certainty',
      'amendment',
      213,
      'This amendment does not, however, bar testimony',
      'particular degree of certainty.',
    ],
    [
      'note-procedure',
      'amendment',
      213,
      'Nothing in the amendment imposes any new,',
      'specific procedures.',
    ],
    [
      'rule-scope',
      'historical',
      43,
      'Rule 1101. Applicability of the Rules',
      '(a) TO COURTS AND JUDGES.',
    ],
  ]
  const sources = sourceSpecs.map(([id, originalId, pageNumber, start, end]) => {
    const { source, pages } = originals.get(originalId)
    return {
      id,
      title: `${source.editionOrOrderDate} 官方原件 PDF 第 ${pageNumber} 页`,
      uri: source.url,
      snapshotAttachmentId: source.sha256,
      locator: `第 ${pageNumber} 页`,
      excerpt: sliceOriginal(pages[pageNumber - 1], start, end),
      asOf: source.editionOrOrderDate,
    }
  })
  const context = (id, location, limitations, date = '2024-12-01') => ({
    domain: 'law',
    ...(['old', 'new-opening', 'new-d', 'rule-1101'].includes(id)
      ? { materialKind: 'statute' }
      : {}),
    jurisdiction: 'United States federal courts, subject to Rule 1101',
    effectLevel:
      id === 'effective'
        ? 'Supreme Court amendment order'
        : ['standard', 'weight', 'scope', 'certainty', 'procedure'].includes(id)
          ? 'official explanatory committee note'
          : 'federal procedural evidence rule',
    ...(['effective', 'new-opening', 'new-d'].includes(id) ? { effectiveFrom: '2023-12-01' } : {}),
    applicabilityDate: date,
    originalLocation: location,
    limitations,
  })
  const claimSpecs = [
    [
      'effective',
      '2023 命令设定 2023-12-01 生效；待决程序仅在公正且可行范围内适用。',
      ['order-effective'],
      'amendment PDF 194',
      '命令范围受待决程序限定',
    ],
    [
      'old',
      '2022 历史汇编 Rule 702 开头列出四项条件；仅供与后版文字比较。',
      ['old-rule'],
      'old PDF 29',
      '旧文未写新增短语不意味着旧法无可靠性门槛',
    ],
    [
      'new-opening',
      '2024 历史汇编开头增加由提出方向法院作更可能为真的展示。',
      ['new-opening'],
      'historical PDF 30',
      '历史文字差异，不作今天适用或个案判断',
    ],
    [
      'new-d',
      '2024 历史汇编 (d) 使用专家意见反映可靠应用的表述。',
      ['new-d'],
      'historical PDF 31',
      '须与 (a)–(c) 一起阅读',
    ],
    [
      'standard',
      '官方 Note 说明 Rule 104(a) 门槛适用于规则中的可靠性要求。',
      ['note-standard'],
      'amendment PDF 210',
      '解释性 Note 不是新增规则条款',
    ],
    [
      'weight',
      '官方 Note 保留部分专家证言异议属于重量而非可采性的情形。',
      ['note-weight'],
      'amendment PDF 211',
      '不能概括为全部异议仅关乎重量',
    ],
    [
      'scope',
      '官方 Note 强调专家意见不得超出其基础与方法可靠支持的范围。',
      ['note-scope'],
      'amendment PDF 212',
      '未评价具体专家意见',
    ],
    [
      'certainty',
      '官方 Note 对实体法要求特定确定程度的证言保留例外。',
      ['note-certainty'],
      'amendment PDF 213',
      '不能概括为全面禁止确定性表述',
    ],
    [
      'procedure',
      '官方 Note 表明修订不新增特定程序。',
      ['note-procedure'],
      'amendment PDF 213',
      '不据此省略可靠性审查',
    ],
    [
      'rule-1101',
      '联邦证据规则的适用范围须按 Rule 1101 的法院和例外条款判断。',
      ['rule-scope'],
      'historical PDF 43',
      '不延伸至所有州法院或全部程序',
    ],
  ]
  const claims = claimSpecs.map(([id, statement, sourceIds, location, limitations]) => ({
    id,
    statement,
    sourceIds,
    type: 'fact',
    confidence: 'low',
    reviewStatus: 'needs_review',
    asOf: id === 'old' ? '2022-12-01' : '2024-12-01',
    jurisdiction: 'United States federal courts, subject to Rule 1101',
    professionalContext: context(
      id,
      location,
      limitations,
      id === 'old' ? '2022-12-01' : '2024-12-01',
    ),
  }))
  const rows = [
    {
      title: 'Rule 702：历史文字与官方说明',
      body: '固定历史时点 2024-12-01。美国联邦证据规则的版本比较候选稿；不证明今天仍适用，也不提供个案法律意见。',
      claims: [],
      footer: '2022、2024 官方汇编及 2023 最高法院修订包；法律审阅待完成',
    },
    {
      title: '命令与时间线',
      body: '2023-04-24 最高法院命令；2023-12-01 生效。其后开始的程序按命令适用；当时待决程序仅在公正且可行范围内适用。此处不推断任何案件的具体适用性。',
      claims: ['effective'],
      footer: '2023 官方修订包 PDF 194；待决程序限定必须保留',
    },
    {
      title: '旧版：仅作文字对照',
      body: '2022 历史汇编 Rule 702 开头列出 (a)–(d) 四项条件；(d) 以专家可靠应用原则和方法到案件事实表述。旧版未含新增短语，不代表可绕开可靠性门槛。',
      claims: ['old'],
      footer: '2022 规则汇编 PDF 29 / 印刷 15；旧文不是当前适用证明',
    },
    {
      title: '新版：开头与 (d) 的差异',
      body: '开头｜2024 历史汇编明写提出方向法院作更可能为真的展示。\n(d)｜专家意见反映原则和方法对案件事实的可靠应用。\n(a)–(c) 条件仍在规则中；黑线稿的删除和新增须对照清洁正文。',
      claims: ['old', 'new-opening', 'new-d'],
      footer: '旧 PDF 29；历史 PDF 30–31；修订包清洁文本 PDF 198',
    },
    {
      title: '官方 Note：门槛与重量',
      body: 'Note 强调 Rule 104(a) 门槛适用于可靠性要求；同时保留部分异议属于证言重量而非可采性的情形。不能把任何一端概括为“所有争议只看重量”或“所有事实争议都排除专家”。',
      claims: ['standard', 'weight'],
      footer: '2023 修订包 Note PDF 210–211；Note 是官方解释',
    },
    {
      title: '官方 Note：意见范围与确定性',
      body: '每项专家意见须留在其基础和方法可靠支持的范围内。对绝对确定性的警示取决于方法和出错风险；官方 Note 同时保留实体法要求特定确定程度的例外。',
      claims: ['scope', 'certainty'],
      footer: '2023 修订包 Note PDF 212–213；不评价具体专家',
    },
    {
      title: '程序与联邦规则适用范围',
      body: '官方 Note 不新增特定程序。Rule 1101 规定联邦证据规则适用的法院、案件及例外；本稿不延伸至所有州法院，也不对任何待决案件作准入结论。',
      claims: ['procedure', 'rule-1101'],
      footer: '修订包 Note PDF 213；2024 历史汇编 Rule 1101 PDF 43–44',
    },
    {
      title: '来源与未完成审阅',
      body: '三份官方 PDF 原件与页码、摘要已冻结；此稿只展示可核对的历史文字与官方解释。法律适用、版权范围、版面与真实 PowerPoint 编辑保存重开仍待复核。',
      claims: [],
      footer: '材料清单与原始 PDF；不得把字面一致当作法律审阅通过',
    },
  ]
  const style = {
    fontFace: 'Noto Sans CJK SC',
    background: 'FFFFFF',
    textColor: '173248',
    accentColor: '087D83',
  }
  const projectId = 'p0-04-rule-702-historical-reference'
  const sections = [
    'legal_question',
    'applicable_materials',
    'applicable_materials',
    'analysis_and_alternatives',
    'analysis_and_alternatives',
    'analysis_and_alternatives',
    'legal_risks',
    'legal_conclusion',
  ]
  const plan = {
    version: 1,
    projectId,
    title: 'Rule 702 历史规则比较：候选汇报',
    domain: 'law',
    brief: {
      objective: '比较 Rule 702 修订前后文字与官方说明，并保留历史时点、法域和程序边界',
      audience: '法律专业审阅者',
      language: 'zh-CN',
      minutes: 8,
      requiredContent: ['命令时间线', '旧新文字差异', '官方 Note', 'Rule 1101 适用边界'],
      constraints: [
        '历史 as-of 2024-12-01',
        '不得提供个案意见',
        '法律专业审阅待完成',
        '真实 PowerPoint 验收待完成',
      ],
    },
    sources,
    claims,
    style,
    slides: rows.map((row, index) => ({
      id: `p${String(index + 1).padStart(2, '0')}`,
      title: row.title,
      purpose: '分层呈现历史规则文字、官方解释和适用边界',
      claimIds: row.claims,
      domainSection: sections[index],
      layout: index === 0 ? 'cover' : index === 7 ? 'summary' : 'content',
      requiredAssets: [],
      acceptanceCriteria: ['原生可编辑文字', '历史版本和来源清楚', '不形成个案结论'],
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
    slides: rows.map((row, index) => ({
      id: plan.slides[index].id,
      title: row.title,
      claimIds: row.claims,
      notes: '中文为待法律专业审阅的组织释义，非官方译本；固定历史 as-of 2024-12-01。',
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
          fontSize: 19,
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
  await writeFile(join(root, 'p0-04-reference.pptx'), result.bytes)
  console.log(
    `P0-04 historical candidate: ${result.bytes.length} bytes, ${deck.slides.length} native slides`,
  )
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
