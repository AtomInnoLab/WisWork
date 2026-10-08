import { createHash } from 'node:crypto'
import { readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = dirname(fileURLToPath(import.meta.url))
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')
const rights = JSON.parse(await readFile(join(root, 'asset-rights.json'), 'utf8'))
const sourceFiles = [
  [
    'article',
    'nasa-2024-article.html',
    'https://science.nasa.gov/earth/earth-observatory/2024-was-the-warmest-year-on-record-153806/',
  ],
  ['svs', 'nasa-svs-5450.html', 'https://svs.gsfc.nasa.gov/5450'],
  [
    'rights',
    'nasa-media-guidelines.html',
    'https://www.nasa.gov/nasa-brand-center/images-and-media/',
  ],
]
const snapshots = await Promise.all(
  sourceFiles.map(async ([id, file, uri]) => {
    const bytes = await readFile(join(root, file))
    return { id, file, uri, digest: sha(bytes) }
  }),
)
const source = (id, title, locator, excerpt) => {
  const snapshot = snapshots.find((item) => item.id === id)
  return {
    id: `${id}-${locator.replace(/[^a-z0-9]/gi, '-').toLowerCase()}`.slice(0, 100),
    title,
    uri: snapshot.uri,
    snapshotAttachmentId: snapshot.digest,
    locator,
    excerpt,
    asOf: '2025-01-11',
  }
}
const sources = [
  source(
    'article',
    'NASA Earth Observatory 2024 温度文章',
    'description',
    'global temperatures in 2024 were 1.28 degrees Celsius (2.30 degrees Fahrenheit) above the agency’s 20th-century baseline',
  ),
  source(
    'svs',
    'NASA SVS 5450 地图说明',
    'description',
    'It does not show absolute temperatures; instead, it shows how much warmer or cooler each region of Earth was compared to the average from 1951 to 1980.',
  ),
  source(
    'svs',
    'NASA SVS 5450 图例说明',
    'map-colors',
    'Average temperatures are shown in white, with higher-than-normal temperatures shown in red and lower-than-normal temperatures in blue.',
  ),
  source(
    'svs',
    'NASA SVS 5450 全球均值',
    'global-mean',
    'Earth’s global surface temperatures in 2024 were the warmest on record -- 1.28 degrees Celsius (2.30 degrees Fahrenheit) above the agency’s 20th-century baseline (1951–1980).',
  ),
  source(
    'rights',
    'NASA Images and Media Usage Guidelines',
    'media-guidelines',
    'NASA Images and Media Usage Guidelines',
  ),
]
const claim = (id, statement, sourceIds, type = 'fact') => ({
  id,
  statement,
  sourceIds,
  type,
  confidence: 'low',
  reviewStatus: 'needs_review',
  asOf: '2025-01-11',
})
const claims = [
  claim('global-mean', 'NASA 报告 2024 年全球地表温度相对其 1951–1980 基线高 1.28 °C。', [
    sources[0].id,
    sources[3].id,
  ]),
  claim('anomaly-not-absolute', '图像显示相对 1951–1980 基线的温度异常，不表示绝对气温。', [
    sources[1].id,
  ]),
  claim('map-colors', '地图中白色为接近基线，红色为高于基线，蓝色为低于基线。', [sources[2].id]),
  claim(
    'scope-limit',
    '单幅全球异常图不能推出任意城市的天气或未来预测。',
    [sources[1].id],
    'judgment',
  ),
  claim('usage-review', '图像署名与 NASA 媒体使用例外需要人工复核。', [sources[4].id], 'judgment'),
]
const style = {
  fontFace: 'Noto Sans CJK SC',
  background: 'FFFFFF',
  textColor: '142D43',
  accentColor: '087D83',
}
const imageId = rights.assetId
const imageDigest = rights.candidates.find((candidate) => candidate.role === 'fallback').sha256
const title = '2024 全球地表温度异常：NASA 来源候选汇报'
const text = (id, value, x, y, w, h, fontSize = 18, extra = {}) => ({
  id,
  kind: 'text',
  text: value,
  x,
  y,
  w,
  h,
  fontSize,
  ...extra,
})
const base = (id, heading, body, claimIds = [], domainSection = 'results_and_data') => ({
  id,
  title: heading,
  claimIds,
  notes: 'NASA 原始资料截至 2025-01-11；科学解释、媒体使用条件与真实 PowerPoint 效果待审阅。',
  elements: [
    text('title', heading, 0.65, 0.45, 12, 0.7, 27, { bold: true }),
    text('body', body, 0.8, 1.55, 11.7, 4.5, 20),
    text(
      'source',
      'NASA Earth Observatory / NASA SVS 5450 · 1951–1980 基线 · 待审阅',
      0.75,
      6.55,
      11.6,
      0.25,
      9,
      { role: 'decoration' },
    ),
  ],
  domainSection,
})
const pages = [
  base(
    'p01',
    '2024 全球地表温度异常',
    'NASA：全球平均异常 +1.28 °C。图片为相对基线的空间分布，来源为 NASA SVS 5450。',
    ['global-mean', 'anomaly-not-absolute'],
    'research_question',
  ),
  base(
    'p02',
    '先明确比较基线',
    '这里的 +1.28 °C 是相对于 NASA 1951–1980 年平均值的异常量；不是 2024 年的绝对气温。',
    ['global-mean', 'anomaly-not-absolute'],
    'methods_and_sample',
  ),
  base(
    'p03',
    '年度结论的证据',
    'NASA Earth Observatory 与 SVS 5450 均报告 2024 年全球平均异常 +1.28 °C；只保留同一基线和时间口径。',
    ['global-mean'],
  ),
  base(
    'p04',
    '原图与图例',
    '红色高于基线，蓝色低于基线，白色接近基线；颜色不支持推算未经标注的地区精确数值。',
    ['anomaly-not-absolute', 'map-colors'],
  ),
  base(
    'p05',
    '哪些结论不能由此推出',
    '这张图不是绝对温度图，也不能单独证明某城市天气、单次事件归因或未来预测。',
    ['anomaly-not-absolute', 'scope-limit'],
    'scope_and_limitations',
  ),
  base(
    'p06',
    '素材回退与生产记录',
    '首选 4K 受控失败后，核对官方 2K 图的摘要、尺寸、来源和缓存，再进行页面生产。',
    [],
    'scope_and_limitations',
  ),
  base(
    'p07',
    '来源、署名和使用条件',
    '图像：NASA Scientific Visualization Studio；数据贡献者：NASA/GSFC GISS。NASA 媒体说明的例外仍需人工审核。',
    ['usage-review'],
    'scope_and_limitations',
  ),
  base(
    'p08',
    '结论与待核查',
    '可陈述：2024 年相对 NASA 1951–1980 基线的全球平均异常为 +1.28 °C。下一步核对图片许可、科学解释和真实 PowerPoint 可编辑性。',
    ['global-mean', 'usage-review'],
    'research_references',
  ),
]
const replaceBody = (page, element) => {
  page.elements[1] = element
}
const card = (page, id, label, value, x, y, w, h, fontSize = 16) => {
  page.elements.push({
    id: `${id}-background`,
    kind: 'shape',
    shape: 'roundRect',
    x,
    y,
    w,
    h,
    fill: 'E8F2F4',
    lineColor: 'C7DEE2',
    role: 'background',
  })
  page.elements.push(
    text(`${id}-label`, label, x + 0.18, y + 0.14, w - 0.36, 0.34, 15, {
      bold: true,
      color: '087D83',
    }),
  )
  page.elements.push(text(`${id}-value`, value, x + 0.18, y + 0.62, w - 0.36, h - 0.75, fontSize))
}
for (const page of pages)
  page.elements.push({
    id: 'accent',
    kind: 'shape',
    shape: 'rect',
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.12,
    fill: '087D83',
    lineColor: '087D83',
    role: 'decoration',
  })
pages[0].elements[1] = text(
  'body',
  '2024 年全球平均异常 +1.28 °C\n相对 NASA 1951–1980 年基线',
  0.75,
  1.7,
  6.1,
  3.2,
  29,
  { bold: true },
)
pages[0].elements.push({
  id: 'map',
  kind: 'image',
  assetId: imageId,
  x: 7.1,
  y: 1.65,
  w: 5.4,
  h: 3.05,
  fit: 'contain',
  altText: 'NASA 2024 全球地表温度异常地图，红色高于基线、蓝色低于基线',
})
card(pages[0], 'metric', '年度指标', 'NASA 全球平均异常 +1.28 °C', 0.75, 5.02, 5.85, 1.17, 15)
card(pages[0], 'basis', '比较口径', '相对 1951–1980 年平均值', 6.73, 5.02, 5.85, 1.17, 15)
replaceBody(
  pages[1],
  text('body', '+1.28 °C', 0.85, 2.0, 5.5, 1.25, 54, { bold: true, color: '087D83' }),
)
card(pages[1], 'baseline', '比较基线', 'NASA 1951–1980 年平均值', 6.67, 1.62, 5.9, 1.85)
card(pages[1], 'not-absolute', '避免误读', '异常量不是 2024 年的绝对气温', 6.67, 3.72, 5.9, 1.85)
replaceBody(
  pages[2],
  text('body', '两份 NASA 原始资料采用同一年度和基线口径。', 0.78, 5.03, 11.8, 0.65, 18),
)
card(
  pages[2],
  'article',
  'Earth Observatory',
  '2024 年全球平均异常：+1.28 °C',
  0.75,
  1.6,
  5.82,
  2.8,
  19,
)
card(
  pages[2],
  'svs',
  'Scientific Visualization Studio',
  '地图与文字说明共同指向 1951–1980 基线',
  6.75,
  1.6,
  5.82,
  2.8,
  18,
)
pages[3].elements[1] = text(
  'body',
  '红：高于基线\n白：接近基线\n蓝：低于基线',
  9.5,
  1.8,
  2.7,
  3.2,
  19,
)
pages[3].elements.push({
  id: 'map',
  kind: 'image',
  assetId: imageId,
  x: 0.72,
  y: 1.55,
  w: 8.35,
  h: 4.7,
  fit: 'contain',
  altText: 'NASA 2024 全球地表温度异常地图，1951–1980 基线',
})
replaceBody(
  pages[4],
  text('body', '解释边界', 0.78, 1.4, 11.6, 0.5, 18, { color: '087D83', bold: true }),
)
card(pages[4], 'absolute', '不是绝对温度', '颜色只表达相对基线的异常值', 0.75, 2.13, 3.75, 3.3, 17)
card(pages[4], 'weather', '不是城市天气', '全球分布图不能给出城市日天气', 4.79, 2.13, 3.75, 3.3, 17)
card(
  pages[4],
  'forecast',
  '不是未来预测',
  '这张历史地图不能单独预测未来',
  8.83,
  2.13,
  3.75,
  3.3,
  17,
)
pages[5].elements[1] = {
  id: 'flow',
  kind: 'table',
  x: 0.75,
  y: 1.55,
  w: 11.8,
  h: 4.5,
  fontSize: 15,
  rows: [
    ['阶段', '受控输入', '核对结果'],
    ['首选 4K', '注入一次失败', '尚未写页'],
    ['官方 2K', '下载并核对', '来源、摘要、尺寸'],
    ['重复取用', '请求同一候选', '命中缓存'],
    ['继续生产', '引用已验证附件', '逐页回读'],
  ],
}
replaceBody(
  pages[6],
  text('body', '使用前保留来源和署名；权利例外等待人工核对。', 0.8, 5.1, 11.7, 0.6, 18),
)
card(
  pages[6],
  'credit',
  '图像与数据',
  'NASA Scientific Visualization Studio；NASA/GSFC GISS',
  0.75,
  1.62,
  5.82,
  2.8,
  17,
)
card(
  pages[6],
  'rights',
  '使用条件',
  'NASA 媒体使用说明有标识、人物及第三方素材例外',
  6.75,
  1.62,
  5.82,
  2.8,
  17,
)
replaceBody(
  pages[7],
  text('body', '2024 年全球平均异常 +1.28 °C', 0.82, 1.48, 11.5, 0.75, 26, {
    bold: true,
    color: '087D83',
  }),
)
card(
  pages[7],
  'conclusion',
  '可陈述',
  '相对 NASA 1951–1980 年基线；地图只表达异常分布',
  0.75,
  2.55,
  5.82,
  2.8,
  17,
)
card(
  pages[7],
  'review',
  '待核查',
  '科学解释、图像使用条件和真实 PowerPoint 可编辑性',
  6.75,
  2.55,
  5.82,
  2.8,
  17,
)
const layout = [
  'cover',
  'content',
  'content',
  'content',
  'content',
  'process',
  'content',
  'summary',
]
const plan = {
  version: 1,
  projectId: 'p0-14-nasa-fallback-candidate',
  title,
  domain: 'science',
  brief: {
    objective:
      '用 NASA 原始资料解释 2024 全球地表温度异常，并在首选图片失败后使用经过验证的官方备用图继续八页生产',
    audience: '内部科学与媒体使用审阅者',
    language: 'zh-CN',
    minutes: 8,
    requiredContent: [
      '+1.28 °C 与 1951–1980 基线',
      '官方图例',
      '故障回退记录',
      'NASA 署名与使用条件',
    ],
    constraints: [
      '不将异常误写为绝对气温',
      '不从地图推出城市天气或未来预测',
      '人工审阅和真实 PowerPoint 验收待完成',
    ],
  },
  sources,
  claims,
  style,
  slides: pages.map((page, index) => ({
    id: page.id,
    title: page.title,
    purpose: page.elements.find((element) => element.id === 'body')?.text ?? page.title,
    claimIds: page.claimIds,
    layout: layout[index],
    domainSection: page.domainSection,
    requiredAssets: [0, 3].includes(index) ? [imageId] : [],
    acceptanceCriteria: ['结论与来源一致', '文本、表格与图片均为可编辑页面对象', '图片及图例可见'],
  })),
}
const deck = {
  version: 1,
  id: plan.projectId,
  title,
  style,
  assets: [{ id: imageId, attachmentId: imageDigest }],
  claims: claims.map((item) => ({
    id: item.id,
    text: item.statement,
    source: item.sourceIds
      .map((sourceId) => sources.find((entry) => entry.id === sourceId).uri)
      .join(' ; '),
    locator: item.sourceIds
      .map((sourceId) => sources.find((entry) => entry.id === sourceId).locator)
      .join(' ; '),
  })),
  slides: pages.map(({ domainSection: _domainSection, ...page }) => page),
}
await writeFile(join(root, 'reference-plan.json'), `${JSON.stringify(plan, null, 2)}\n`)
await writeFile(join(root, 'reference-deck.json'), `${JSON.stringify(deck, null, 2)}\n`)
