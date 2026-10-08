const { readFile, writeFile } = require('node:fs/promises')
const { createHash } = require('node:crypto')
const { join } = require('node:path')
const prettier = require('prettier')
const JSZip = require('jszip')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')
const {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
  presentationPlanClaims,
} = require('@wiswork/pptx-engine/presentation-plan')

async function main() {
  const root = __dirname
  const scenario = JSON.parse(await readFile(join(root, 'scenario.json'), 'utf8'))
  if (scenario.caseId !== 'PPT-P0-06' || scenario.pages.length !== 8)
    throw new Error('P0-06 scope changed')
  const inputs = new Map()
  for (const file of scenario.documentInputs) {
    const bytes = await readFile(join(root, file))
    inputs.set(file, {
      text: bytes.toString('utf8'),
      hash: createHash('sha256').update(bytes).digest('hex'),
    })
  }
  const contract = inputs.get('synthetic-nda.txt').text
  const probe = /^测试追踪码：([^\r\n]+)$/m.exec(contract)?.[1]
  if (!probe || !/^PRIVATE-SIM-P0-06-[A-Z0-9]+$/.test(probe))
    throw new Error('synthetic disclosure probe missing')
  const sourceSpecs = [
    ['status', 'synthetic-nda.txt', '本文件由 WisWork 为产品测试自行编写', '合同状态与本机授权'],
    ...Array.from({ length: 8 }, (_, index) => [
      `c${index + 1}`,
      'synthetic-nda.txt',
      `第 ${index + 1} 条｜`,
      `合同第 ${index + 1} 条`,
    ]),
    ['policy-use', 'review-policy.txt', '材料性质：', '测试审阅口径'],
    ['policy-risk', 'review-policy.txt', '风险分级：', '审阅风险分级'],
    ...Array.from({ length: 5 }, (_, index) => [
      `r${index + 1}`,
      'revision-notes.txt',
      `R${index + 1}｜`,
      `修订意见 R${index + 1}`,
    ]),
  ]
  const sources = sourceSpecs.map(([id, file, anchor, locator]) => {
    const { text, hash } = inputs.get(file)
    const lines = text.split(/\r?\n/).filter((line) => line.startsWith(anchor))
    if (lines.length !== 1 || lines[0].includes(probe))
      throw new Error(`unsafe or missing source: ${id}`)
    return {
      id,
      title: `${file}：${locator}`,
      uri: `attachment:${hash}`,
      snapshotAttachmentId: hash,
      locator,
      excerpt: lines[0],
      asOf: '2026-09-28',
    }
  })
  const claimSpecs = [
    [
      'status',
      '材料为 WisWork 自有虚构且未签署的测试合同；仅限获准本机项目制作内部审阅稿。',
      ['status', 'policy-use'],
      'fact',
      '未签署、无真实合同效力',
    ],
    [
      'c1',
      '第 1 条定义保密信息，并列出已合法知悉、非违约公开、第三方合法提供及独立开发四类例外。',
      ['c1'],
      'fact',
      '四类例外不可删除',
    ],
    [
      'c2',
      '第 2 条只允许为示意联合设计试验评估使用信息，其他项目、模型训练、公开搜索或营销须先获书面同意。',
      ['c2'],
      'fact',
      '不得扩写授权用途',
    ],
    [
      'c3',
      '第 3 条限制必要知悉人员并要求保护和记录；获准人员、离岗撤权、分包方范围仍待确认。',
      ['c3', 'r1'],
      'judgment',
      '后半是修订建议，未成为合同义务',
    ],
    [
      'c4',
      '第 4 条允许依法必要披露；仅在法律允许时提前通知，通知例外和披露范围待专业审阅。',
      ['c4', 'r2'],
      'judgment',
      '不能改写为一律提前通知',
    ],
    [
      'c5',
      '第 5 条写明双方书面签署才生效；草稿约定终止后三年保密延续，当前尚未签署。',
      ['c5'],
      'fact',
      '期限是未生效草稿文字',
    ],
    [
      'c6',
      '第 6 条草稿约定三十日返还或删除，并保留法律强制保存例外；备份和删除证明待确认。',
      ['c6', 'r3'],
      'judgment',
      '不能宣称系统已实现自动删除',
    ],
    [
      'c7',
      '第 7 条仅写先协商补救，损失、举证、救济和适用法律仍未确定。',
      ['c7'],
      'fact',
      '不得补写赔偿数字或上限',
    ],
    [
      'c8',
      '第 8 条的通知地址、适用法律、争议机构和签署人均留待确认。',
      ['c8', 'r4'],
      'judgment',
      '不得推断法域、法院或仲裁机构',
    ],
    [
      'followup',
      '修订意见 R1–R5 仅是待专业审阅的建议，应由合同责任人与法律/隐私审阅人确认。',
      ['r1', 'r4', 'r5'],
      'judgment',
      '建议未写入或生效；不得作为法律意见',
    ],
  ]
  const sourcesById = new Map(sources.map((source) => [source.id, source]))
  const claims = claimSpecs.map(([id, statement, sourceIds, type, limitations]) => ({
    id,
    statement,
    sourceIds,
    type,
    confidence: 'low',
    reviewStatus: 'needs_review',
    asOf: '2026-09-28',
    jurisdiction: 'unspecified in unsigned synthetic draft',
    professionalContext: {
      domain: 'law',
      materialKind: 'contract',
      jurisdiction: 'unspecified in unsigned synthetic draft',
      effectLevel: 'unsigned synthetic test draft; no legal effect',
      applicabilityDate: '2026-09-28',
      originalLocation: sourceIds.map((sourceId) => sourcesById.get(sourceId).locator).join('; '),
      limitations,
    },
  }))
  const rows = [
    [
      '虚构保密合同：内部审阅',
      'WisWork 自有虚构材料｜双方未签署｜仅限获准本机测试项目。此稿是条款审阅候选，不构成真实合同或法律意见。',
      [],
      '内部测试；法律与隐私审阅待完成',
    ],
    [
      '审阅范围与授权',
      '范围｜八条测试合同、内部审阅口径和五条修订意见。\n边界｜禁止原始材料外发、跨项目复用及未经授权的模型请求；实际发送记录须另行核验。',
      ['status'],
      '合同状态与 review-policy.txt；不包含禁止披露字段',
    ],
    [
      '条款地图：定义、例外与期限',
      '第 1 条｜保留四类例外：既有合法知悉、非违约公开、合法第三方提供、独立开发。\n第 5 条｜书面签署后才生效；终止后三年保密延续只是未签署草稿的约定。',
      ['c1', 'c5'],
      '合同第 1、5 条；不把草稿当已生效义务',
    ],
    [
      '目的限制与访问控制',
      '合同原文｜第 2 条限定使用目的；第 3 条只允许必要知悉人员访问并要求保护和记录。\n审阅建议｜R1 要求确认人员登记、撤权、分包方及记录保留；尚未生效。',
      ['c2', 'c3'],
      '合同第 2、3 条；R1 为待审建议',
    ],
    [
      '依法披露与通知限定',
      '合同原文｜第 4 条仅允许在必要范围内依法披露，且只在法律允许时提前通知。\n待确认｜R2 建议专业审阅通知例外和披露范围责任；不得写成一律提前通知。',
      ['c4'],
      '合同第 4 条；R2 为待审建议',
    ],
    [
      '返还、删除与留存例外',
      '合同原文｜第 6 条草稿约定三十日返还或删除；法律强制保存副本可在必要范围内保留并继续受约束。\n待确认｜R3 提出备份、灾备与删除证明；不宣称系统已自动删除。',
      ['c6'],
      '合同第 6 条；R3 为待审建议',
    ],
    [
      '救济、通知与争议未定',
      '第 7 条｜先协商补救；损失、举证、救济与适用法律未定。\n第 8 条｜通知地址、法律、争议机构和签署人均未定；R4 建议签署前确认。',
      ['c7', 'c8'],
      '合同第 7、8 条；不得推断法域或赔偿数额',
    ],
    [
      '修订与人工复核',
      'R1–R5 是建议，不是已生效义务。合同责任人、法律与隐私审阅人须核对逐条映射、授权范围及实际发送记录，并检查 PPTX、截图、来源账本和诊断中的禁止字段。',
      ['followup'],
      '法律/隐私、实际发送与真实 PowerPoint 验收待完成',
    ],
  ]
  const sections = [
    'legal_question',
    'applicable_materials',
    'applicable_materials',
    'analysis_and_alternatives',
    'legal_risks',
    'legal_risks',
    'analysis_and_alternatives',
    'legal_conclusion',
  ]
  const style = {
    fontFace: 'Noto Sans CJK SC',
    background: 'FFFFFF',
    textColor: '173248',
    accentColor: '087D83',
  }
  const projectId = 'p0-06-synthetic-nda-review-reference'
  const plan = {
    version: 1,
    projectId,
    title: '虚构保密合同风险汇报：候选稿',
    domain: 'law',
    brief: {
      objective: '在授权本机测试范围内逐条审阅虚构未签合同，区分原文、建议和待确认问题',
      audience: '内部法律与隐私审阅者',
      language: 'zh-CN',
      minutes: 8,
      requiredContent: ['八条款', '第 1、4、5、6 条限定', 'R1–R5 建议', '禁止披露复核'],
      constraints: [
        '不调用模型服务',
        '不外发禁止披露字段',
        '不推断法域或损失数字',
        '专业与真实 PowerPoint 验收待完成',
      ],
    },
    sources,
    claims,
    style,
    slides: rows.map(([title, , ,], index) => ({
      id: `p${String(index + 1).padStart(2, '0')}`,
      title,
      purpose: scenario.pages[index].purpose,
      claimIds: rows[index][2],
      domainSection: sections[index],
      layout: index === 0 ? 'cover' : index === 7 ? 'summary' : 'content',
      requiredAssets: [],
      acceptanceCriteria: [
        '原生可编辑文字',
        '条款号与原文准确',
        '原文、建议和待确认分开',
        '不包含禁止披露字段',
      ],
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
      notes: '仅供自有虚构合同的本机测试；未签署、无法律效力；法律和隐私审阅待完成。',
      elements: [
        { kind: 'text', id: 'title', x: 0.65, y: 0.48, w: 12, h: 0.7, text: title, fontSize: 26 },
        ...body.split('\n').map((text, line) => ({
          kind: 'text',
          id: line ? `body-${line + 1}` : 'body',
          x: 0.9,
          y: 1.55 + line * 1.5,
          w: 11.4,
          h: body.includes('\n') ? 1.35 : 4.75,
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
          text: footer,
          fontSize: 9,
        },
      ],
    })),
  }
  assertDeckMatchesPresentationPlan(deck, parsedPlan)
  const result = await compilePresentationDeck(deck)
  if ([JSON.stringify(plan), JSON.stringify(deck)].some((value) => value.includes(probe)))
    throw new Error('disclosure probe in plan or deck')
  const pptx = await JSZip.loadAsync(result.bytes)
  for (const [name, entry] of Object.entries(pptx.files)) {
    if (!entry.dir && (await entry.async('nodebuffer')).includes(Buffer.from(probe)))
      throw new Error(`disclosure probe in PPTX: ${name}`)
  }
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
  await writeFile(join(root, 'p0-06-reference.pptx'), result.bytes)
  console.log(
    `P0-06 synthetic NDA candidate: ${result.bytes.length} bytes, ${deck.slides.length} native slides; disclosure probe excluded`,
  )
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
