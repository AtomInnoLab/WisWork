import type { PresentationClaim, PresentationDeck, PresentationStyle } from './presentation'
import {
  PROFESSIONAL_CONTEXT_SCHEMA,
  parsePresentationProfessionalContext,
  type PresentationProfessionalContext,
} from '@wiswork/project-store/presentation-professional-context'
import {
  PRESENTATION_DECK_SCHEMA,
  PRESENTATION_HEIGHT,
  PRESENTATION_TEXT_BUDGET,
  PRESENTATION_WIDTH,
} from './presentation'
import {
  type Schema,
  text,
  number,
  choice,
  array,
  object,
  id,
  color,
  valid,
} from './presentation-schema'

/** Mandatory deck JSON may consume 192 KiB of the 256 KiB compile transport.
 * The remaining 64 KiB accommodates content/geometry and the request envelope;
 * optional content and assets still require the normal compile transport checks.
 */
export const PRESENTATION_PLAN_COMPILED_BYTE_BUDGET = 192 * 1024

/** Optional domain story structure; labels guide planning, not factual verification. */
export const PRESENTATION_DOMAIN_PROFILES = {
  pitch: {
    title: '路演',
    sections: ['audience_problem', 'solution', 'evidence', 'business_case', 'ask'],
    labels: ['目标用户问题', '方案', '验证证据', '商业论证', '明确请求'],
    questions: [
      '问题与目标受众是否清楚？',
      '关键数字是否有来源或明确标为假设？',
      '最后是否有明确请求？',
    ],
  },
  report: {
    title: '汇报',
    sections: ['executive_summary', 'findings', 'supporting_evidence', 'risks', 'actions'],
    labels: ['执行摘要', '主要发现', '支撑证据', '风险', '行动建议'],
    questions: ['结论能否回溯到证据？', '风险与限制是否可见？', '行动项是否有负责人和时间？'],
  },
  training: {
    title: '培训',
    sections: ['objectives', 'concept', 'demonstration', 'practice', 'recap'],
    labels: ['学习目标', '核心概念', '演示', '练习', '回顾'],
    questions: ['学习目标是否可检查？', '是否有练习或应用环节？', '总结是否回扣目标？'],
  },
  research: {
    title: '研究报告',
    sections: ['question', 'method', 'results', 'limitations', 'references'],
    labels: ['研究问题', '方法', '结果', '局限', '参考来源'],
    questions: ['方法和样本口径是否可追溯？', '结果与推断是否区分？', '局限与原始来源是否完整？'],
  },
  sales: {
    title: '销售方案',
    sections: ['customer_problem', 'offer', 'proof', 'value', 'next_step'],
    labels: ['客户问题', '方案内容', '效果证据', '客户价值', '下一步'],
    questions: ['客户问题是否来自已知材料？', '产品效果主张是否有证据？', '下一步是否具体？'],
  },
  science: {
    title: '科研制作',
    sections: [
      'research_question',
      'methods_and_sample',
      'results_and_data',
      'scope_and_limitations',
      'research_references',
    ],
    labels: ['研究问题', '方法与样本', '结果与原始数据', '范围与局限', '研究引用'],
    questions: [
      '原论文/数据集和版本是否可追溯？',
      '统计范围、方法、样本与图表原数据是否一致？',
      '结果、推断与局限是否完整保留？',
    ],
  },
  law: {
    title: '法律制作',
    sections: [
      'legal_question',
      'applicable_materials',
      'analysis_and_alternatives',
      'legal_risks',
      'legal_conclusion',
    ],
    labels: ['法律问题', '适用材料', '分析与替代观点', '法律风险', '限定结论'],
    questions: [
      '法域、效力层级和适用日期是否声明？',
      '法条/案例原位置与相反观点是否保留？',
      '限定语和适用边界是否由专业人员核验？',
    ],
  },
  finance: {
    title: '金融制作',
    sections: [
      'financial_question',
      'reporting_basis',
      'financial_results',
      'financial_risks_and_scenarios',
      'financial_actions',
    ],
    labels: ['金融问题', '报告口径', '财务结果', '风险与情景', '行动建议'],
    questions: [
      '报告期、时点、币种、单位与会计口径是否可比较？',
      '数字与公式能否回溯原披露？',
      '预测依据、假设与风险是否由专业人员核验？',
    ],
  },
} as const
export type PresentationDomainKind = keyof typeof PRESENTATION_DOMAIN_PROFILES
export type PresentationDomainSection =
  (typeof PRESENTATION_DOMAIN_PROFILES)[PresentationDomainKind]['sections'][number]
const domainKinds = Object.keys(PRESENTATION_DOMAIN_PROFILES)
const domainSections = Object.values(PRESENTATION_DOMAIN_PROFILES).flatMap((profile) => [
  ...profile.sections,
])

export interface PresentationProfessionalWorkflow {
  readonly version: 1
  readonly domain: 'science' | 'law' | 'finance'
  readonly sourcePriority: readonly string[]
  readonly contextFields: readonly string[]
  readonly reviewSteps: readonly {
    readonly id: string
    readonly title: string
    readonly tools: readonly string[]
    readonly instruction: string
  }[]
  readonly manualChecks: readonly string[]
  readonly disclosure: string
}
const professionalGuidance = {
  science: {
    sourcePriority: ['原始学术论文', '官方数据集', '学术机构与标准制定机构'],
    contextFields: [
      'materialKind',
      'publicationId',
      'version',
      'sample',
      'method',
      'statisticalBasis',
      'limitations',
    ],
    manualChecks: [
      '核对样本、方法和统计范围，区分显著性与实际意义。',
      '逐图核对原始数据、单位、样本与误差表达。',
      '核对结论和推断的适用范围，保留局限及相反证据。',
    ],
    reviewInstruction:
      '审阅论文/数据集/标准标识与版本、样本、方法、统计依据和局限。字段缺口与 secondary/unverified 声明来源需人工检查；原文出现不能证明科学结论正确。',
  },
  law: {
    sourcePriority: ['官方法律法规', '司法裁判案例', '监管机构材料', '正式合同'],
    contextFields: [
      'materialKind',
      'jurisdiction',
      'effectLevel',
      'effectiveFrom',
      'effectiveUntil',
      'applicabilityDate',
      'caseNumber',
      'originalLocation',
      'limitations',
    ],
    manualChecks: [
      '由专业人员核对法域、效力层级及适用日期的真实法律效力。',
      '逐字检查法条、案例或合同原位置与限定语，保留替代分析。',
      '核对适用对象、例外和风险；无失效日期不能推断仍有效。',
    ],
    reviewInstruction:
      '检查专业法域与通用法域差异、适用日期明确超出已声明生效区间的警告、案例编号和原文位置。日期比较只是范围提示，不证明规则仍有效或适用于具体案件。',
  },
  finance: {
    sourcePriority: ['监管或交易所披露', '经审计财务报表', '公司投资者关系材料', '权威市场数据'],
    contextFields: [
      'materialKind',
      'reportingPeriod',
      'asOf',
      'currency',
      'unit',
      'accountingBasis',
      'formula',
      'limitations',
    ],
    manualChecks: [
      '核对期间、时点、币种、单位和会计口径的可比性。',
      '回溯数字、公式、输入与原披露，检查图表是否保留原数据口径。',
      '逐项审阅预测依据、假设和情景风险；计算复现不证明预测可信。',
    ],
    reviewInstruction:
      '检查专业时点与通用 asOf 差异、计算单位/币种不一致以及专业字段缺口。保留原报告期和会计口径，不将声明来源等级、审计标签或计算成功当作真实性认证。',
  },
} as const
/** Local guidance is cloned; following it never certifies professional accuracy. */
export function presentationProfessionalWorkflow(
  domain?: string,
): PresentationProfessionalWorkflow | undefined {
  if (domain !== 'science' && domain !== 'law' && domain !== 'finance') return undefined
  const guide = professionalGuidance[domain]
  return structuredClone({
    version: 1,
    domain,
    sourcePriority: guide.sourcePriority,
    contextFields: guide.contextFields,
    reviewSteps: [
      {
        id: 'original_research',
        title: '读取指定原研究',
        tools: ['read_research_ledger'],
        instruction:
          '读取完整指定研究版本及原来源、双方冲突与局限，不能用最新版本代替。优先专业原始/官方来源，也可保留用户材料；缺时点或法域保持 unknown，不猜填。',
      },
      {
        id: 'freeze_plan',
        title: '保存领域章节与冻结映射',
        tools: ['save_presentation_plan'],
        instruction:
          '设置对应 domain 和全部必需章节，绑定实际研究 ID/sequence/draftDigest，逐字段保留原事实 professionalContext 与通用标签。章节约束组织结构，不认证内容。',
      },
      {
        id: 'literal_evidence',
        title: '读取完整冻结原文窗口',
        tools: ['read_presentation_claim_evidence'],
        instruction: guide.reviewInstruction,
      },
      {
        id: 'scoped_review',
        title: '记录有范围的 Agent 判断',
        tools: ['record_presentation_claim_review'],
        instruction:
          '在实际读过的原文窗口内以 notes 保存支持判断。可选 source_assessment 保存 historical Agent 意见，其 basis 写入 source_assessment.basis 的绝对 UTF-16 原文位置；不要提交 top-level basis。若评估权威性/时效性，声明比较框架、时点和范围；这些维度仍未认证，不为叙事删掉相反材料或限定。',
      },
      {
        id: 'page_content',
        title: '检查页面内容',
        tools: ['check_presentation_page_content'],
        instruction:
          '检查冻结主张与页面表达、数字、引用和限定是否保留；内容匹配不等于事实支持或专业语义通过。',
      },
      {
        id: 'delivery_review',
        title: '读取问题与人工门槛',
        tools: ['read_presentation_delivery_report'],
        instruction:
          '审阅全部缺上下文、专业范围、来源和研究冲突警告。保留未决项；manualChecks 未实际完成不能宣称专业验收、项目完成或已交付。',
      },
    ],
    manualChecks: guide.manualChecks,
    disclosure:
      '这是本次读取的当前版本专业制作与审阅指引，冻结任务仅保留所选领域，不保证跨代码版本的指引不变。字段完整、章节齐备、原文字面匹配或执行流程不认证事实、来源权威性、时效、法律适用或金融预测；专业语义和实际 PowerPoint 验收仍需人工及真实任务检查。',
  })
}

/** Durable planning metadata; source URIs are never fetched or treated as verified evidence. */
export interface PresentationPlan {
  version: 1
  projectId: string
  title: string
  research?: {
    ledgerId: string
    sequence: number
    draftDigest: string
    sources: { sourceId: string; researchSourceId: string }[]
    claims: { claimId: string; researchClaimId: string }[]
  }
  brief: {
    objective: string
    audience: string
    language: string
    minutes: number
    requiredContent: string[]
    constraints: string[]
  }
  sources: {
    id: string
    title: string
    uri: string
    snapshotAttachmentId?: string
    locator?: string
    excerpt: string
    asOf?: string
  }[]
  claims: {
    id: string
    statement: string
    type: 'fact' | 'quote' | 'calculation' | 'judgment' | 'assumption'
    sourceIds: string[]
    confidence: 'high' | 'medium' | 'low'
    reviewStatus: 'needs_review'
    asOf?: string
    jurisdiction?: string
    professionalContext?: PresentationProfessionalContext
    calculation?: {
      formula: string
      inputs: string[]
      unit?: string
      currency?: string
      reproduction?: {
        bindings: { name: string; inputIndex: number; value: number; sourceId: string }[]
        expected: number
      }
    }
  }[]
  style: PresentationStyle
  domain?: PresentationDomainKind
  parallelism?: 1 | 2
  brandKit?: {
    id: string
    revision: number
    name: string
    allowedColors: string[]
    logo?: { assetId: string; assetDigest: string; placement: 'cover' | 'all' }
    layoutComponents?: {
      id: string
      name: string
      layout: PresentationPlan['slides'][number]['layout']
      slots: {
        id: string
        kind: 'text' | 'shape' | 'image' | 'table' | 'chart'
        x: number
        y: number
        w: number
        h: number
      }[]
    }[]
  }
  slides: {
    id: string
    title: string
    purpose: string
    locked?: boolean
    claimIds: string[]
    layout: 'cover' | 'content' | 'comparison' | 'process' | 'chart' | 'summary'
    domainSection?: PresentationDomainSection
    dependsOn?: string[]
    layoutComponentId?: string
    requiredAssets: string[]
    acceptanceCriteria: string[]
  }[]
}
export type PresentationBrandKit = NonNullable<PresentationPlan['brandKit']>

/** Tool discovery and runtime validation share bounded structural rules. */
export const PRESENTATION_PLAN_SCHEMA: Schema = object(
  {
    version: { type: 'number', enum: [1] },
    projectId: id,
    title: text(300, 1),
    research: object({
      ledgerId: { ...id, maxLength: 128 },
      sequence: number(1, Number.MAX_SAFE_INTEGER),
      draftDigest: { ...text(64, 64), pattern: '^[a-f0-9]{64}$' },
      sources: array(object({ sourceId: id, researchSourceId: { ...id, maxLength: 128 } }), 256),
      claims: array(object({ claimId: id, researchClaimId: { ...id, maxLength: 128 } }), 256),
    }),
    brief: object({
      objective: text(4000, 1),
      audience: text(1000, 1),
      language: text(80, 1),
      minutes: number(0.1, 1440),
      requiredContent: array(text(2000, 1), 100),
      constraints: array(text(2000, 1), 100),
    }),
    sources: array(
      object(
        {
          id,
          title: text(300, 1),
          uri: text(500, 1),
          snapshotAttachmentId: { ...text(64, 64), pattern: '^[a-f0-9]{64}$' },
          locator: text(200),
          excerpt: text(12000),
          asOf: text(100, 1),
        },
        ['id', 'title', 'uri', 'excerpt'],
      ),
      256,
    ),
    claims: array(
      object(
        {
          id,
          statement: text(12000, 1),
          type: choice('fact', 'quote', 'calculation', 'judgment', 'assumption'),
          sourceIds: array(id, 3),
          confidence: choice('high', 'medium', 'low'),
          reviewStatus: choice('needs_review'),
          asOf: text(100, 1),
          jurisdiction: text(300, 1),
          professionalContext: PROFESSIONAL_CONTEXT_SCHEMA as Schema,
          calculation: object(
            {
              formula: text(2000, 1),
              inputs: array(text(1000, 1), 32, 1),
              unit: text(100, 1),
              currency: text(100, 1),
              reproduction: object({
                bindings: array(
                  object({
                    name: { ...text(32, 1), pattern: '^[A-Za-z][A-Za-z0-9_]*$' },
                    inputIndex: number(0, 31),
                    value: number(-1e12, 1e12),
                    sourceId: id,
                  }),
                  32,
                  1,
                ),
                expected: number(-1e12, 1e12),
              }),
            },
            ['formula', 'inputs'],
          ),
        },
        ['id', 'statement', 'type', 'sourceIds', 'confidence', 'reviewStatus'],
      ),
      256,
    ),
    style: PRESENTATION_DECK_SCHEMA.properties!.style!,
    domain: choice(...domainKinds),
    parallelism: { type: 'number', enum: [1, 2] },
    brandKit: object(
      {
        id,
        revision: number(1, 1_000_000),
        name: text(160, 1),
        allowedColors: array(color, 32, 1),
        logo: object({
          assetId: id,
          assetDigest: { ...text(64, 64), pattern: '^[a-f0-9]{64}$' },
          placement: choice('cover', 'all'),
        }),
        layoutComponents: array(
          object({
            id,
            name: text(160, 1),
            layout: choice('cover', 'content', 'comparison', 'process', 'chart', 'summary'),
            slots: array(
              object({
                id,
                kind: choice('text', 'shape', 'image', 'table', 'chart'),
                x: number(0, PRESENTATION_WIDTH),
                y: number(0, PRESENTATION_HEIGHT),
                w: number(0.01, PRESENTATION_WIDTH),
                h: number(0.01, PRESENTATION_HEIGHT),
              }),
              32,
              1,
            ),
          }),
          32,
        ),
      },
      ['id', 'revision', 'name', 'allowedColors'],
    ),
    slides: array(
      object(
        {
          id,
          title: text(300, 1),
          purpose: text(2000, 1),
          locked: { type: 'boolean' },
          claimIds: array(id, 32),
          layout: choice('cover', 'content', 'comparison', 'process', 'chart', 'summary'),
          domainSection: choice(...domainSections),
          dependsOn: array(id, 31),
          layoutComponentId: id,
          requiredAssets: array(text(1000, 1), 32),
          acceptanceCriteria: array(text(2000, 1), 32),
        },
        ['id', 'title', 'purpose', 'claimIds', 'layout', 'requiredAssets', 'acceptanceCriteria'],
      ),
      32,
      1,
    ),
  },
  ['version', 'projectId', 'title', 'brief', 'sources', 'claims', 'style', 'slides'],
)

export function parsePresentationBrandKit(input: unknown): PresentationBrandKit {
  if (!valid(input, PRESENTATION_PLAN_SCHEMA.properties!.brandKit!))
    throw new Error('presentation_brand_kit_invalid')
  const kit = input as PresentationBrandKit
  const colors = kit.allowedColors.map((value) => value.toUpperCase())
  if (!Number.isSafeInteger(kit.revision) || new Set(colors).size !== colors.length)
    throw new Error('presentation_brand_kit_invalid')
  const components = kit.layoutComponents ?? []
  if (new Set(components.map((item) => item.id)).size !== components.length)
    throw new Error('presentation_brand_kit_invalid')
  for (const component of components) {
    if (
      new Set(component.slots.map((slot) => slot.id)).size !== component.slots.length ||
      component.slots.some(
        (slot) => slot.x + slot.w > PRESENTATION_WIDTH || slot.y + slot.h > PRESENTATION_HEIGHT,
      )
    )
      throw new Error('presentation_brand_kit_invalid')
  }
  return structuredClone(kit)
}

function reject(reason: string): never {
  throw new Error(`presentation_plan_invalid:${reason}`)
}
function unique(values: string[], label: string): void {
  if (new Set(values).size !== values.length) reject(`duplicate_${label}`)
}

/** Resolves local evidence without replacing or fetching the original source URI. */
export function presentationSourceAttachmentId(
  source: Pick<PresentationPlan['sources'][number], 'uri' | 'snapshotAttachmentId'>,
): string | undefined {
  const legacy = /^attachment:([a-f0-9]{64})$/.exec(source.uri)?.[1]
  if (legacy && source.snapshotAttachmentId && legacy !== source.snapshotAttachmentId)
    reject('source_snapshot_conflict')
  return source.snapshotAttachmentId ?? legacy
}

export function parsePresentationPlan(input: unknown): PresentationPlan {
  if (!valid(input, PRESENTATION_PLAN_SCHEMA)) reject('schema')
  const plan = input as PresentationPlan
  for (const claim of plan.claims) {
    if (Object.hasOwn(claim, 'professionalContext')) {
      try {
        parsePresentationProfessionalContext(claim.professionalContext)
      } catch {
        reject('professional_context')
      }
    }
  }
  if (
    plan.style.fontFallbacks &&
    new Set([plan.style.fontFace, ...plan.style.fontFallbacks]).size !==
      plan.style.fontFallbacks.length + 1
  )
    reject('font_fallback')
  if (plan.brandKit) {
    let brandKit: PresentationBrandKit
    try {
      brandKit = parsePresentationBrandKit(plan.brandKit)
    } catch {
      reject('brand_kit')
    }
    const colors = brandKit.allowedColors.map((value) => value.toUpperCase())
    if (
      [plan.style.background, plan.style.textColor, plan.style.accentColor].some(
        (value) => !colors.includes(value.toUpperCase()),
      )
    )
      reject('brand_kit')
    if (
      plan.brandKit.logo?.placement === 'cover' &&
      !plan.slides.some((slide) => slide.layout === 'cover')
    )
      reject('brand_kit_logo_scope')
  }
  const components = new Map(plan.brandKit?.layoutComponents?.map((item) => [item.id, item]) ?? [])
  if (plan.domain) {
    const sections: readonly string[] = PRESENTATION_DOMAIN_PROFILES[plan.domain].sections
    if (
      plan.slides.some(
        (slide) => !slide.domainSection || !sections.includes(slide.domainSection),
      ) ||
      sections.some((section) => !plan.slides.some((slide) => slide.domainSection === section))
    )
      reject('domain_section')
  } else if (plan.slides.some((slide) => slide.domainSection)) reject('domain_section')
  for (const slide of plan.slides) {
    if (slide.layoutComponentId && components.get(slide.layoutComponentId)?.layout !== slide.layout)
      reject('layout_component')
  }
  if (new TextEncoder().encode(JSON.stringify(plan)).byteLength > 192 * 1024) reject('size_budget')
  unique(
    plan.sources.map((source) => source.id),
    'source',
  )
  unique(
    plan.claims.map((claim) => claim.id),
    'claim',
  )
  unique(
    plan.slides.map((slide) => slide.id),
    'slide',
  )
  for (const source of plan.sources) presentationSourceAttachmentId(source)
  const sourceIds = new Set(plan.sources.map((source) => source.id))
  const claimIds = new Set(plan.claims.map((claim) => claim.id))
  if (plan.research) {
    if (!Number.isSafeInteger(plan.research.sequence)) reject('research_sequence')
    unique(
      plan.research.sources.map((mapping) => mapping.sourceId),
      'research_source_mapping',
    )
    unique(
      plan.research.claims.map((mapping) => mapping.claimId),
      'research_claim_mapping',
    )
    if (
      plan.research.sources.some((mapping) => !sourceIds.has(mapping.sourceId)) ||
      plan.research.claims.some((mapping) => !claimIds.has(mapping.claimId))
    )
      reject('research_mapping_target')
  }
  for (const claim of plan.claims) {
    unique(claim.sourceIds, 'source_reference')
    if (claim.sourceIds.some((source) => !sourceIds.has(source))) reject('source_reference')
    if (['fact', 'quote', 'calculation'].includes(claim.type) && !claim.sourceIds.length)
      reject('source_required')
    if (claim.type === 'calculation' && !claim.calculation) reject('calculation_required')
    const reproduction = claim.calculation?.reproduction
    if (reproduction) {
      if (
        claim.type !== 'calculation' ||
        reproduction.bindings.length !== claim.calculation!.inputs.length
      )
        reject('reproduction_inputs')
      unique(
        reproduction.bindings.map((binding) => binding.name),
        'binding_name',
      )
      unique(
        reproduction.bindings.map((binding) => String(binding.inputIndex)),
        'binding_index',
      )
      for (const binding of reproduction.bindings) {
        if (
          !Number.isInteger(binding.inputIndex) ||
          binding.inputIndex >= claim.calculation!.inputs.length ||
          ['prototype', 'constructor', '__proto__'].includes(binding.name) ||
          !claim.sourceIds.includes(binding.sourceId)
        )
          reject('reproduction_binding')
      }
    }
  }
  const previousSlideIds = new Set<string>()
  for (const slide of plan.slides) {
    if (plan.parallelism === 2 && slide.dependsOn === undefined) reject('page_dependency')
    unique(slide.claimIds, 'claim_reference')
    if (slide.claimIds.some((claim) => !claimIds.has(claim))) reject('claim_reference')
    unique(slide.dependsOn ?? [], 'page_dependency')
    if (slide.dependsOn?.some((dependency) => !previousSlideIds.has(dependency)))
      reject('page_dependency')
    previousSlideIds.add(slide.id)
  }
  const claims = mappedClaims(plan)
  const requiredText = claims.reduce(
    (sum, claim) => sum + claim.text.length + claim.source.length + (claim.locator?.length ?? 0),
    plan.slides.reduce((sum, slide) => sum + slide.title.length, 0),
  )
  if (requiredText > PRESENTATION_TEXT_BUDGET) reject('text_budget')
  const mandatoryDeck = {
    version: 1,
    id: plan.projectId,
    title: plan.title,
    style: plan.style,
    assets: [],
    claims,
    slides: plan.slides.map(({ id, title, claimIds }) => ({ id, title, claimIds, elements: [] })),
  }
  if (
    new TextEncoder().encode(JSON.stringify(mandatoryDeck)).byteLength >
    PRESENTATION_PLAN_COMPILED_BYTE_BUDGET
  )
    reject('compiled_byte_budget')
  return structuredClone(plan)
}

export function presentationPlanClaims(plan: PresentationPlan): PresentationClaim[] {
  return mappedClaims(parsePresentationPlan(plan))
}

/** A changed rule set under the same brand ID must advance its own revision. */
export function assertBrandKitRevision(previous: PresentationPlan, next: PresentationPlan): void {
  const before = previous.brandKit
  const after = next.brandKit
  if (!before || !after || before.id !== after.id) return
  const rules = (kit: NonNullable<PresentationPlan['brandKit']>) =>
    JSON.stringify({
      name: kit.name,
      allowedColors: kit.allowedColors.map((color) => color.toUpperCase()),
      logo: kit.logo,
      layoutComponents: kit.layoutComponents,
    })
  if (rules(before) !== rules(after) && after.revision <= before.revision)
    throw new Error('presentation_plan_invalid:brand_kit_revision')
  if (after.revision < before.revision)
    throw new Error('presentation_plan_invalid:brand_kit_revision')
}

function mappedClaims(plan: PresentationPlan): PresentationClaim[] {
  const sources = new Map(plan.sources.map((source) => [source.id, source]))
  return plan.claims.map((claim) => {
    const evidence = claim.sourceIds.map((id) => sources.get(id)!)
    const locator = evidence
      .map((source) => source.locator)
      .filter((value) => value?.trim())
      .join(' ; ')
    return {
      id: claim.id,
      text: claim.statement,
      source: evidence.length
        ? evidence.map((source) => source.uri).join(' ; ')
        : `未核验：${claim.type}`,
      ...(locator ? { locator } : {}),
    }
  })
}

/** Deterministic identity/content binding only; this makes no visual or factual QA claim. */
export function assertDeckMatchesPresentationPlan(
  deck: PresentationDeck,
  plan: PresentationPlan,
): void {
  const expectedClaims = presentationPlanClaims(plan)
  const mismatch = (reason: string): never => {
    throw new Error(`presentation_plan_mismatch:${reason}`)
  }
  if (deck.id !== plan.projectId || deck.title !== plan.title) mismatch('project')
  for (const key of ['fontFace', 'background', 'textColor', 'accentColor'] as const)
    if (deck.style[key] !== plan.style[key]) mismatch('style')
  if (
    JSON.stringify(deck.style.fontFallbacks ?? []) !==
    JSON.stringify(plan.style.fontFallbacks ?? [])
  )
    mismatch('style')
  if (plan.brandKit) {
    const colors = new Set(plan.brandKit.allowedColors.map((value) => value.toUpperCase()))
    for (const slide of deck.slides)
      for (const element of slide.elements) {
        const used =
          element.kind === 'text'
            ? [element.color]
            : element.kind === 'shape'
              ? [element.fill, element.lineColor]
              : []
        if (used.some((value) => value && !colors.has(value.toUpperCase()))) mismatch('brand_color')
      }
    if (plan.brandKit.logo) {
      if (!deck.assets.some((asset) => asset.id === plan.brandKit!.logo!.assetId))
        mismatch('brand_logo_asset')
      for (const [index, slide] of deck.slides.entries())
        if (
          (plan.brandKit.logo.placement === 'all' || plan.slides[index]!.layout === 'cover') &&
          !slide.elements.some(
            (element) =>
              element.kind === 'image' && element.assetId === plan.brandKit!.logo!.assetId,
          )
        )
          mismatch('brand_logo')
    }
  }
  if (deck.slides.length !== plan.slides.length) mismatch('slides')
  for (const [index, slide] of deck.slides.entries()) {
    const planned = plan.slides[index]!
    if (
      slide.id !== planned.id ||
      slide.title !== planned.title ||
      JSON.stringify(slide.claimIds ?? []) !== JSON.stringify(planned.claimIds)
    )
      mismatch('slide')
    if (planned.layoutComponentId) {
      const component = plan.brandKit?.layoutComponents?.find(
        (item) => item.id === planned.layoutComponentId,
      )
      if (!component) throw new Error('presentation_plan_mismatch:layout_component')
      for (const slot of component.slots) {
        const element = slide.elements.find((item) => item.id === slot.id)
        if (
          !element ||
          element.kind !== slot.kind ||
          element.x !== slot.x ||
          element.y !== slot.y ||
          element.w !== slot.w ||
          element.h !== slot.h
        )
          mismatch('layout_component')
      }
    }
  }
  if (deck.claims.length !== expectedClaims.length) mismatch('claims')
  for (const [index, claim] of deck.claims.entries()) {
    const expected = expectedClaims[index]!
    if (
      claim.id !== expected.id ||
      claim.text !== expected.text ||
      claim.source !== expected.source ||
      claim.locator !== expected.locator
    )
      mismatch('claim')
  }
}
