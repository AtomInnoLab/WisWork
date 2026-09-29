import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import PptxGenJS from 'pptxgenjs'

const root = dirname(fileURLToPath(import.meta.url))
const image = join(root, 'images/nasa-2024-temperature-anomaly-2k.png')
const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_WIDE'
pptx.author = 'WisWork benchmark materials'
pptx.subject = 'PPT-P0-14 external asset failure and controlled fallback reference'
pptx.title = '2024 全球地表温度异常：外部素材恢复示例'
pptx.lang = 'zh-CN'

const C = {
  navy: '142D43',
  teal: '087D83',
  ice: 'E8F2F4',
  white: 'FFFFFF',
  gray: '516575',
  amber: 'E49C50',
}
const source = 'NASA SVS 5450（2025-01-10）；1951–1980 基线；仅固定 2024 数据口径'
function page(title, number) {
  const s = pptx.addSlide()
  s.background = { color: C.white }
  s.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.14,
    line: { color: C.teal },
    fill: { color: C.teal },
  })
  s.addText(title, {
    x: 0.62,
    y: 0.42,
    w: 12,
    h: 0.62,
    fontFace: 'Arial',
    fontSize: 26,
    bold: true,
    color: C.navy,
    margin: 0,
    breakLine: false,
  })
  s.addText(source, {
    x: 0.64,
    y: 7.11,
    w: 11.9,
    h: 0.18,
    fontFace: 'Arial',
    fontSize: 8,
    color: C.gray,
    margin: 0,
  })
  s.addText(String(number).padStart(2, '0'), {
    x: 12.5,
    y: 7.08,
    w: 0.28,
    h: 0.19,
    fontFace: 'Arial',
    fontSize: 8,
    color: C.gray,
    align: 'right',
    margin: 0,
  })
  return s
}
function text(s, value, x, y, w, h, options = {}) {
  s.addText(value, {
    x,
    y,
    w,
    h,
    fontFace: 'Arial',
    fontSize: 18,
    color: C.navy,
    margin: 0.08,
    valign: 'mid',
    breakLine: false,
    ...options,
  })
}
function card(s, heading, body, x, y, w, h) {
  s.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    line: { color: 'C7DEE2', width: 1 },
    fill: { color: C.ice },
  })
  text(s, heading, x + 0.18, y + 0.17, w - 0.36, 0.42, { fontSize: 18, bold: true, color: C.teal })
  text(s, body, x + 0.18, y + 0.64, w - 0.36, h - 0.78, { fontSize: 14, valign: 'top' })
}

let s = page('2024 全球地表温度异常', 1)
text(s, '外部素材失败与受控降级 · PPT-P0-14 参考稿', 0.68, 1.28, 7.3, 0.6, {
  fontSize: 21,
  color: C.gray,
})
s.addImage({ path: image, x: 7.7, y: 1.7, w: 4.7, h: 2.65 })
card(
  s,
  '固定时点',
  'NASA 于 2025-01-10 发布该可视化；本示例只讨论 2024 年及其原报告基线。',
  0.69,
  2.1,
  6.2,
  1.5,
)
card(
  s,
  '素材规则',
  '先获取并验证授权图片，确认缓存或降级方案，再写入演示文稿。',
  0.69,
  4.02,
  6.2,
  1.5,
)

s = page('问题与证据边界', 2)
card(s, '科学问题', '如何说明“异常值”相对于基线，而不把颜色误读成绝对温度？', 0.7, 1.3, 5.9, 1.6)
card(
  s,
  '来源',
  'NASA Earth Observatory 文章与 NASA SVS 5450 原始可视化；图像署名按 SVS 页面保留。',
  6.74,
  1.3,
  5.9,
  1.6,
)
card(
  s,
  '不推断',
  '单幅全球分布图不能证明任意城市的天气、个体事件原因或未来预测。',
  0.7,
  3.35,
  5.9,
  1.6,
)
card(
  s,
  '待人工复核',
  '固定图例、来源版权例外、专业解释及真实 PowerPoint 的可读性与编辑性。',
  6.74,
  3.35,
  5.9,
  1.6,
)

s = page('异常值的含义', 3)
text(s, '2024 全球平均地表温度', 0.72, 1.3, 5.7, 0.58, { fontSize: 21, bold: true })
text(s, '+1.28 °C', 0.72, 2.05, 5.2, 1.13, { fontSize: 54, bold: true, color: C.teal })
text(s, '相对 NASA 1951–1980 年平均值；不是 2024 年的绝对气温。', 0.74, 3.47, 5.75, 1.15, {
  fontSize: 18,
})
card(
  s,
  '数值出处',
  'NASA SVS 5450 的文字说明与 NASA Earth Observatory 2025-01-11 文章；两个来源采用同一基线。',
  6.75,
  1.42,
  5.78,
  2.0,
)
card(
  s,
  '表达限制',
  '不把年度异常与单月异常、地区平均或 1850–1900 年基线混用。',
  6.75,
  3.82,
  5.78,
  1.55,
)

s = page('原图与图例', 4)
s.addImage({ path: image, x: 0.66, y: 1.22, w: 8.8, h: 4.95 })
card(
  s,
  '阅读提示',
  '红色表示高于 1951–1980 平均，蓝色表示低于该平均。图中采用 °C / °F 色标；不要以颜色推算未经标注的精确地区数值。',
  9.63,
  1.38,
  2.98,
  3.0,
)
text(
  s,
  '图像：NASA Scientific Visualization Studio；数据：NASA/GSFC GISS。',
  0.72,
  6.27,
  11.7,
  0.36,
  { fontSize: 12, color: C.gray },
)

s = page('素材获取与写页分离', 5)
const flow = [
  ['1', '核来源', 'NASA SVS 页面、固定 URL 与署名'],
  ['2', '下载', '首选 4K；失败时尝试 2K 候选'],
  ['3', '验字节', 'MIME、尺寸、摘要与缓存回执'],
  ['4', '再写页', '把已验证的本地素材交给页面生产'],
]
flow.forEach(([n, heading, body], i) => {
  const x = 0.7 + (i % 2) * 6.1,
    y = 1.35 + Math.floor(i / 2) * 2.25
  card(s, `${n}  ${heading}`, body, x, y, 5.82, 1.8)
})

s = page('一次受控超时的恢复', 6)
const rows = [
  [{ text: '阶段' }, { text: '受控输入' }, { text: '应记录的结果' }],
  ['首选 URL', '4K 图片响应超时一次', '超时、候选序号、未写入页面'],
  ['第二候选', '2K 图片成功', '摘要、尺寸、来源与许可记录'],
  ['重复取用', '同一候选再次请求', '命中持久缓存，不重新下载'],
  ['写入页面', '使用已验证的本地图片', '原任务继续；已完成页不重复生产'],
]
s.addTable(rows, {
  x: 0.7,
  y: 1.4,
  w: 11.9,
  h: 4.35,
  border: { type: 'solid', color: 'B9D4D8', pt: 1 },
  fill: C.white,
  color: C.navy,
  fontFace: 'Arial',
  fontSize: 15,
  margin: 0.1,
  autoFit: false,
  colW: [2.05, 3.8, 6.05],
  rowH: [0.65, 0.9, 0.9, 0.9, 0.9],
})
text(s, '这是预期故障注入流程；参考稿本身不是实际运行日志。', 0.75, 6.12, 11.6, 0.4, {
  fontSize: 13,
  color: C.gray,
})

s = page('权限与交付检查', 7)
card(
  s,
  '使用范围',
  'NASA 媒体使用说明一般允许教育/信息用途；该 SVS 页面要求署名。正式商用、标识、第三方素材例外仍需人工核查。',
  0.7,
  1.33,
  5.9,
  2.2,
)
card(
  s,
  '本地备份',
  '2K 与 4K 原图在材料包内保留 SHA256；降级时必须展示最终采用的 2K 图及其来源。',
  6.72,
  1.33,
  5.9,
  2.2,
)
card(
  s,
  '交付门禁',
  '逐页截图、图片可见性、来源署名、保存关闭重开和可编辑对象需在真实 PowerPoint 另行验证。',
  0.7,
  3.83,
  11.92,
  1.7,
)

s = page('结论与待核查项', 8)
card(
  s,
  '可陈述',
  '该图显示 2024 年相对 1951–1980 年基线的全球地表温度异常分布；NASA 报告全球平均异常为 +1.28 °C。',
  0.7,
  1.4,
  11.9,
  1.7,
)
card(
  s,
  '故障恢复',
  '首选 URL 超时后可切换候选并命中缓存；是否达到原方案恢复指标，要用真实 PC / Office 回执和日志证明。',
  0.7,
  3.55,
  5.82,
  1.8,
)
card(
  s,
  '未证明',
  '本稿未经气候领域审阅，也未在 PowerPoint 实机执行、保存重开或核验许可例外。',
  6.8,
  3.55,
  5.82,
  1.8,
)

await pptx.writeFile({ fileName: join(root, 'p0-14-reference.pptx') })
