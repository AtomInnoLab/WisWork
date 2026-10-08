import PptxGenJS from 'pptxgenjs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_WIDE'
pptx.author = 'WisWork benchmark team'
pptx.subject = 'PPT-P0-12 self-authored brand template'
pptx.title = 'WisWork Benchmark Research Template'
pptx.lang = 'zh-CN'
pptx.theme = {
  headFontFace: 'WisWork Benchmark Display 2026',
  bodyFontFace: 'Noto Sans CJK SC',
  lang: 'zh-CN',
}

const shape = pptx.ShapeType
const colors = { ink: '102A43', teal: '007F86', pale: 'E8F5F3', white: 'FFFFFF', muted: '52606D' }
function text(slide, value, x, y, w, h, fontSize, color = colors.ink, heading = false) {
  slide.addText(value, {
    x,
    y,
    w,
    h,
    fontSize,
    color,
    margin: 0,
    fontFace: heading ? 'WisWork Benchmark Display 2026' : 'Noto Sans CJK SC',
    bold: heading,
  })
}
function frame(slide, page) {
  slide.background = { color: colors.white }
  slide.addShape(shape.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.13,
    line: { color: colors.teal },
    fill: { color: colors.teal },
  })
  text(slide, 'WW  /  RESEARCH', 0.62, 0.3, 3, 0.3, 10, colors.teal)
  slide.addShape(shape.line, {
    x: 0.62,
    y: 6.95,
    w: 12.05,
    h: 0,
    line: { color: 'D8E2E9', width: 1 },
  })
  text(slide, '来源与日期须在每页可见  •  自制基准模板', 0.62, 7.02, 9.5, 0.2, 8, colors.muted)
  text(slide, String(page).padStart(2, '0'), 12.1, 7.02, 0.55, 0.2, 8, colors.muted)
}

const cover = pptx.addSlide()
frame(cover, 1)
cover.addShape(shape.rect, {
  x: 0.62,
  y: 1.18,
  w: 0.16,
  h: 4.6,
  line: { color: colors.teal },
  fill: { color: colors.teal },
})
text(cover, '研究汇报标题', 1.03, 1.65, 10.8, 1, 36, colors.ink, true)
text(cover, '副标题与研究范围', 1.03, 2.85, 10.8, 0.5, 20, colors.muted)
text(cover, '报告日期 / 汇报人 / 来源版本', 1.03, 4.87, 10.8, 0.4, 13, colors.teal)

const content = pptx.addSlide()
frame(content, 2)
text(content, '单页结论标题', 0.62, 0.94, 11.6, 0.7, 28, colors.ink, true)
content.addShape(shape.rect, {
  x: 0.62,
  y: 1.9,
  w: 7.05,
  h: 4.55,
  line: { color: colors.pale },
  fill: { color: colors.pale },
})
text(content, '左侧正文或原生图表区域\n\n关键发现应有来源与限定语。', 0.9, 2.17, 6.4, 3.8, 21)
content.addShape(shape.rect, {
  x: 8.05,
  y: 1.9,
  w: 4.63,
  h: 4.55,
  line: { color: 'D8E2E9' },
  fill: { color: colors.white },
})
text(content, '右侧证据或图片区域', 8.38, 2.27, 3.95, 0.5, 16, colors.teal)
text(content, '来源：作者、年份、页码\n截至：YYYY-MM-DD', 8.38, 5.45, 3.95, 0.55, 10, colors.muted)

const chart = pptx.addSlide()
frame(chart, 3)
text(chart, '数据图表页', 0.62, 0.94, 11.6, 0.7, 28, colors.ink, true)
chart.addShape(shape.rect, {
  x: 0.62,
  y: 1.9,
  w: 12.05,
  h: 4.55,
  line: { color: 'D8E2E9' },
  fill: { color: colors.white },
})
text(
  chart,
  '在此放置原生可编辑图表；标明分母、单位、期间与来源。',
  1.03,
  2.2,
  10.9,
  0.6,
  18,
  colors.muted,
)
text(chart, '本模板只定义布局，不预填研究数据。', 1.03, 5.57, 10.9, 0.4, 12, colors.teal)

await pptx.writeFile({
  fileName: join(dirname(fileURLToPath(import.meta.url)), 'wiswork-benchmark-brand-template.pptx'),
})
