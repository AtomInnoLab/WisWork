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

async function main() {
  const root = __dirname
  const manifest = JSON.parse(await readFile(join(root, 'manifest.json')))
  const basis = JSON.parse(await readFile(join(root, 'basis.json')))
  const { pdfToPages } = await import('../../../../packages/file-parse/src/pdf.ts')
  const originals = new Map()
  for (const source of manifest.sources) {
    const bytes = await readFile(join(root, source.file))
    if (
      bytes.length !== source.bytes ||
      createHash('sha256').update(bytes).digest('hex') !== source.sha256
    )
      throw new Error(`official original changed: ${source.id}`)
    originals.set(source.id, { source, pages: await pdfToPages(bytes) })
  }
  const sourceSpecs = [
    ['g-facts', 'google', 5, 'Oracle America, Inc.'],
    ['g-issue', 'google', 18, 'poses two questions'],
    ['g-assume', 'google', 19, 'We shall assume'],
    ['g-use', 'google', 26, 'This is the declaring code'],
    ['g-result', 'google', 39, 'We reach the conclusion'],
    ['g-remand', 'google', 40, 'and the case is remanded'],
    ['w-facts', 'warhol', 7, 'In 1984, Vanity Fair'],
    ['w-scope', 'warhol', 18, 'the only question before this Court'],
    ['w-factor', 'warhol', 19, 'Even though Orange Prince'],
    ['w-limit', 'warhol', 27, 'expresses no opinion'],
    ['w-comparison', 'warhol', 29, 'specific purposes of the origi-'],
    ['w-result', 'warhol', 44, 'Because this Court agrees'],
  ]
  const sources = sourceSpecs.map(([id, originalId, page, anchor]) => {
    const { source, pages } = originals.get(originalId)
    const text = pages[page - 1]
    const at = text.indexOf(anchor)
    if (at < 0) throw new Error(`official majority opinion anchor missing: ${id}`)
    return {
      id,
      title: `${source.caseName} 多数意见 PDF 第 ${page} 页`,
      uri: source.url,
      snapshotAttachmentId: source.sha256,
      locator: `第 ${page} 页`,
      excerpt: text.slice(at, Math.min(text.length, at + 220)),
      asOf: source.decisionDate,
    }
  })
  const claimSpecs = [
    [
      'g-facts',
      'Google 案涉及 Android 对 Java 接口声明代码的重新实现；此页将事实背景与法律评价分开。',
      ['g-facts', 'g-use'],
      '事实背景；不推断所有 API 的版权状态',
    ],
    [
      'g-question',
      'Google 案提出可版权性与合理使用两个争点；最高法院仅为本案分析假定 API 可版权。',
      ['g-issue', 'g-assume'],
      '假定不是可版权性终局裁判',
    ],
    [
      'g-holding',
      '法院在该重新实现情境下认定合理使用，撤销并发回。',
      ['g-result', 'g-remand'],
      '限于裁判所述使用情境',
    ],
    [
      'w-facts',
      'Warhol 案涉及 Goldsmith 摄影作品和 Prince Series；本稿聚焦被争议的杂志许可用途。',
      ['w-facts', 'w-scope'],
      '不评价其他创作、展示或销售用途',
    ],
    [
      'w-holding',
      '法院就该特定商业许可用途认定合理使用第一因素有利于 Goldsmith，并维持原判。',
      ['w-scope', 'w-factor', 'w-result'],
      '第一因素和本案程序范围不等于所有用途最终侵权结论',
    ],
    [
      'w-limit',
      '多数意见明确不评价原 Prince Series 的创作、展示或出售。',
      ['w-limit'],
      '不同用途须分别分析',
    ],
    [
      'comparison',
      '两案的比较须分别核对作品性质、具体使用目的和市场关系；这是待审研究推论，不是第三案规则。',
      ['g-use', 'w-comparison', 'w-limit'],
      '研究推论；不得用于第三案胜负预测',
    ],
  ]
  const sourceById = new Map(sources.map((source) => [source.id, source]))
  const claims = claimSpecs.map(([id, statement, sourceIds, limitations]) => {
    const originalId = id.startsWith('g-') ? 'google' : id.startsWith('w-') ? 'warhol' : undefined
    const original = originalId ? originals.get(originalId).source : undefined
    return {
      id,
      statement,
      sourceIds,
      type: id === 'comparison' ? 'judgment' : 'fact',
      confidence: 'low',
      reviewStatus: 'needs_review',
      asOf: basis.asOf,
      jurisdiction: basis.jurisdiction,
      professionalContext: {
        domain: 'law',
        ...(original ? { materialKind: 'case', caseNumber: original.docket } : {}),
        jurisdiction: basis.jurisdiction,
        effectLevel: original
          ? 'Supreme Court majority opinion'
          : 'research comparison, not a court holding',
        applicabilityDate: basis.asOf,
        originalLocation: sourceIds.map((sourceId) => sourceById.get(sourceId).title).join('; '),
        limitations,
      },
    }
  })
  const rows = [
    [
      '范围与来源',
      '美国联邦版权法｜历史分析截至 2024-12-01。比较 Google v. Oracle（18-956）与 Warhol v. Goldsmith（21-869）两份官方多数意见；非个案法律意见，法律专业审阅待完成。',
      '两份最高法院 slip opinion；官方发布不等于最终装订版',
    ],
    [
      'Google：事实与程序',
      'Android 重新实现 Java 接口声明代码。先记录事实与争议对象，再看法院在此情境下的法律评价。',
      'Google 多数意见 PDF 5、26；事实与判断分开',
    ],
    [
      'Google：争点与受限判断',
      '争点｜API 可版权性与合理使用。\n处理｜法院为分析假定可版权性，认定具体重新实现构成合理使用；撤销并发回。',
      'Google 多数意见 PDF 18–19、39–40；不概括所有 API',
    ],
    [
      'Warhol：事实与程序',
      'Goldsmith 的 Prince 摄影与 Warhol 的 Prince Series。此处聚焦被诉的杂志商业许可用途，区分作品创作与具体使用。',
      'Warhol 多数意见 PDF 7、18；不复用原件图像',
    ],
    [
      'Warhol：判断与明确限定',
      '第一因素｜就该杂志许可用途有利于 Goldsmith；最高法院维持原判。\n明确限定｜多数意见不评价原系列作品的创作、展示或出售。',
      'Warhol 多数意见 PDF 18–19、27、44；仅本案第一因素',
    ],
    [
      '两案比较矩阵',
      '对象｜接口声明代码／摄影衍生图像\n用途｜软件重新实现／特定杂志许可\n范围｜假定可版权性后的合理使用／具体许可的第一因素\n结果｜撤销并发回／维持',
      '研究比较，非两案对第三案的共同裁判规则',
    ],
    [
      '差异归因与不得泛化',
      '比较推论须核对作品性质、使用目的和市场关系。不得推出“所有 API 可复制”“所有改编均侵权”“商用必不合理使用”或“新含义单独决定合理使用”。',
      'Google PDF 26、39；Warhol PDF 27、29；法律审阅待完成',
    ],
    [
      '核验与人工审阅',
      '原件摘要、正文页码与逐字摘录已固定。须继续核对中文归纳、程序状态、隐私与图像权利、版面及真实 PowerPoint 保存重开；不得将工程产物标为专业验收通过。',
      '官方 PDF 原件；syllabus 与多数意见分开',
    ],
  ]
  const sections = [
    'legal_question',
    'applicable_materials',
    'analysis_and_alternatives',
    'applicable_materials',
    'analysis_and_alternatives',
    'analysis_and_alternatives',
    'legal_risks',
    'legal_conclusion',
  ]
  const style = {
    fontFace: 'Noto Sans CJK SC',
    background: 'FFFFFF',
    textColor: '173248',
    accentColor: '087D83',
  }
  const projectId = 'p0-05-copyright-case-comparison-reference'
  const plan = {
    version: 1,
    projectId,
    title: '最高法院版权裁判比较：候选汇报',
    domain: 'law',
    brief: {
      objective: '按具体争点、使用和限定比较两案多数意见，保留历史时点与研究推论边界',
      audience: basis.audience,
      language: 'zh-CN',
      minutes: 8,
      requiredContent: ['事实与程序', 'Google 假定与判断', 'Warhol 特定用途与限定', '比较矩阵'],
      constraints: [
        '历史 as-of 2024-12-01',
        '非个案意见',
        '不复用原件图像',
        '法律专业及真实 PowerPoint 验收待完成',
      ],
    },
    sources,
    claims,
    style,
    slides: basis.slides.map((slide, index) => ({
      id: `p${String(index + 1).padStart(2, '0')}`,
      title: rows[index][0],
      purpose: slide.purpose,
      claimIds: slide.statementIds,
      domainSection: sections[index],
      layout: index === 0 ? 'cover' : index === 7 ? 'summary' : 'content',
      requiredAssets: [],
      acceptanceCriteria: ['原生可编辑文字', '具体用途与来源清楚', '研究推论和裁判判断分开'],
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
    slides: rows.map(([title, body, footer], index) => ({
      id: plan.slides[index].id,
      title,
      claimIds: plan.slides[index].claimIds,
      notes: '中文为待法律专业审阅的研究归纳，非官方译文；历史 as-of 2024-12-01。',
      elements: [
        { kind: 'text', id: 'title', x: 0.65, y: 0.48, w: 12, h: 0.7, text: title, fontSize: 26 },
        ...body.split('\n').map((text, line) => ({
          kind: 'text',
          id: line ? `body-${line + 1}` : 'body',
          x: 0.9,
          y: 1.55 + line * (index === 5 ? 1.0 : 1.25),
          w: 11.4,
          h: body.includes('\n') ? (index === 5 ? 0.85 : 1.1) : 4.75,
          text,
          fontSize: index === 5 ? 17 : 19,
        })),
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
  await writeFile(join(root, 'p0-05-reference.pptx'), result.bytes)
  console.log(
    `P0-05 case comparison candidate: ${result.bytes.length} bytes, ${deck.slides.length} native slides`,
  )
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
