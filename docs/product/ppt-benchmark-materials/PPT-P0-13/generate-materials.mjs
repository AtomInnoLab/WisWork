import { createHash } from 'node:crypto'
import { readFile, mkdir, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import PptxGenJS from 'pptxgenjs'
import { PNG } from 'pngjs'

const root = dirname(fileURLToPath(import.meta.url))
const imageDir = join(root, 'images')
await mkdir(imageDir, { recursive: true })
const palette = {
  navy: [18, 43, 67],
  teal: [0, 128, 134],
  blue: [39, 105, 155],
  amber: [225, 151, 64],
  pale: [235, 244, 247],
  white: [255, 255, 255],
  gray: [175, 198, 207],
}

// All twelve images are self-authored schematics, not measurements from the paper.
function schematic(index) {
  const png = new PNG({ width: 960, height: 540 })
  const rect = (x, y, w, h, color) => {
    for (let yy = Math.max(0, y); yy < Math.min(540, y + h); yy++)
      for (let xx = Math.max(0, x); xx < Math.min(960, x + w); xx++)
        png.data.set([...color, 255], (yy * 960 + xx) * 4)
  }
  const disk = (cx, cy, radius, color) => {
    for (let y = cy - radius; y <= cy + radius; y++)
      for (let x = cx - radius; x <= cx + radius; x++)
        if ((x - cx) ** 2 + (y - cy) ** 2 <= radius ** 2) rect(x, y, 1, 1, color)
  }
  const line = (x1, y1, x2, y2, color) => {
    const steps = Math.max(Math.abs(x2 - x1), Math.abs(y2 - y1))
    for (let i = 0; i <= steps; i++)
      disk(
        Math.round(x1 + ((x2 - x1) * i) / steps),
        Math.round(y1 + ((y2 - y1) * i) / steps),
        3,
        color,
      )
  }
  rect(0, 0, 960, 540, palette.pale)
  rect(34, 34, 892, 472, palette.white)
  rect(34, 34, 892, 10, index % 3 ? palette.teal : palette.amber)
  switch (index % 6) {
    case 0:
      for (let i = 0; i < 5; i++) {
        disk(170 + i * 155, 270 + (i % 2) * 42, 49, i % 2 ? palette.blue : palette.teal)
        if (i)
          line(
            170 + (i - 1) * 155,
            270 + ((i - 1) % 2) * 42,
            170 + i * 155,
            270 + (i % 2) * 42,
            palette.gray,
          )
      }
      break
    case 1:
      for (let i = 0; i < 6; i++) {
        const h = 65 + ((i * 83 + index * 19) % 245)
        rect(150 + i * 115, 425 - h, 56, h, i % 2 ? palette.blue : palette.teal)
      }
      break
    case 2:
      for (let i = 0; i < 4; i++) {
        rect(120 + i * 178, 125 + i * 44, 235, 65, i % 2 ? palette.blue : palette.teal)
        rect(155 + i * 178, 205 + i * 44, 140, 13, palette.gray)
      }
      break
    case 3:
      for (let i = 0; i < 4; i++) {
        const x = 250 + (i % 2) * 450,
          y = 170 + Math.floor(i / 2) * 230
        disk(x, y, 66, i % 2 ? palette.blue : palette.teal)
        if (i % 2) line(x - 450, y, x, y, palette.gray)
        if (i > 1) line(x, y - 230, x, y, palette.gray)
      }
      break
    case 4:
      for (let row = 0; row < 4; row++)
        for (let col = 0; col < 6; col++) {
          const tone = (row * 7 + col * 3 + index) % 5
          rect(
            130 + col * 119,
            102 + row * 91,
            96,
            68,
            tone < 2 ? palette.teal : tone < 4 ? palette.blue : palette.amber,
          )
        }
      break
    default:
      for (let i = 0; i < 5; i++) {
        const x = 145 + i * 145,
          y = 385 - ((i * 71 + index * 37) % 230)
        line(x, y, x + 145, 160 + ((i * 41) % 160), palette.teal)
        disk(x, y, 16, palette.amber)
      }
  }
  rect(70 + index * 60, 458, 26, 26, palette.amber)
  return PNG.sync.write(png)
}

const hashes = []
const assets = []
for (let number = 1; number <= 12; number++) {
  const name = `schematic-${String(number).padStart(2, '0')}.png`
  const bytes = schematic(number)
  await writeFile(join(imageDir, name), bytes)
  const sha256 = createHash('sha256').update(bytes).digest('hex')
  hashes.push(`${sha256}  images/${name}`)
  assets.push({
    assetId: `P0-13-IMG-${String(number).padStart(2, '0')}`,
    file: `images/${name}`,
    sha256,
    author: 'WisWork benchmark team',
    usePermission: 'Authorized for this benchmark and image replacement testing',
    attribution: `WisWork self-authored schematic ${number}; not research measurement data`,
  })
}
const rights = Buffer.from(`${JSON.stringify({ version: 1, assets }, null, 2)}\n`)
await writeFile(join(root, 'asset-rights.json'), rights)
hashes.push(`${createHash('sha256').update(rights).digest('hex')}  asset-rights.json`)

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_WIDE'
pptx.author = 'WisWork benchmark team'
pptx.subject = 'PPT-P0-13 self-authored image-dense research draft'
pptx.title = 'Computational reproducibility workshop: benchmark draft'
pptx.lang = 'zh-CN'
pptx.theme = { headFontFace: 'Noto Sans CJK SC', bodyFontFace: 'Noto Sans CJK SC', lang: 'zh-CN' }
const pages = [
  [
    '计算可复现性工作坊：研究汇报',
    'Deardorff 等，PLOS ONE 2020；本稿仅为图片编辑验收候选。',
    [1, 2],
  ],
  ['研究问题与汇报范围', '关注工作坊后的计算实践；不把观察变化表述为因果效果。', [3, 4]],
  [
    '研究方法与时间点',
    '工作坊前访谈 14 人；三个月后 12 人完成后访谈。论文 PDF 第 1、3、5 页。',
    [5, 6],
  ],
  ['参与者与证据路径', '左图为验收替换目标；示意图不代表原始研究数据。', [7, 8]],
  [
    '六项清单均值',
    '前测 1.6/6，三个月后 2.2/6；差异未达统计显著，p=0.318。论文 PDF 第 5 页。',
    [9],
  ],
  [
    '开源软件使用：不同分母',
    'Table 1：前测 7/14，后测 10/12；不能当作同一批人逐个转变。论文 PDF 第 5 页。',
    [10],
  ],
  [
    '解释边界与局限',
    '小样本、招募/应答偏差、单人编码、定量分析功效不足。论文 PDF 第 9 页。',
    [11, 12],
  ],
  ['结论与后续审阅', '本稿为候选现稿；科研审阅和真实 PowerPoint 保存重开尚未完成。', [1, 3]],
]
const text = (slide, value, x, y, w, h, fontSize, color = '122B43', bold = false) =>
  slide.addText(value, {
    x,
    y,
    w,
    h,
    fontFace: 'Noto Sans CJK SC',
    fontSize,
    color,
    bold,
    margin: 0,
  })

for (const [index, [title, note, images]] of pages.entries()) {
  const slide = pptx.addSlide()
  slide.background = { color: 'FFFFFF' }
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.12,
    line: { color: '008086' },
    fill: { color: '008086' },
  })
  text(slide, title, 0.62, 0.43, 12, 0.55, 27, '122B43', true)
  text(
    slide,
    `PPT-P0-13 / ${String(index + 1).padStart(2, '0')}`,
    0.62,
    1.08,
    4.5,
    0.25,
    10,
    '008086',
  )
  for (const [slot, number] of images.entries()) {
    const x = 0.62 + slot * 6.14
    slide.addImage({
      path: join(imageDir, `schematic-${String(number).padStart(2, '0')}.png`),
      x,
      y: 1.55,
      w: 5.78,
      h: 3.25,
      altText: `自制示意图 ${number}，非研究测量数据`,
    })
    text(
      slide,
      `自制示意图 ${String(number).padStart(2, '0')}（非研究数据）`,
      x,
      4.88,
      5.78,
      0.3,
      11,
      '58717D',
    )
  }
  if (index === 4 || index === 5) {
    const values = index === 4 ? [1.6, 2.2] : [7, 10]
    slide.addChart(
      pptx.ChartType.bar,
      [{ name: '论文报告值', labels: ['前测', '三个月'], values }],
      {
        x: 6.75,
        y: 1.56,
        w: 5.9,
        h: 3.55,
        chartColors: ['008086'],
        showLegend: false,
        showValue: true,
        showTitle: false,
        dataLabelFormatCode: index === 4 ? '0.0' : '0',
        catAxisLabelFontFace: 'Noto Sans CJK SC',
        valAxisLabelFontFace: 'Noto Sans CJK SC',
      },
    )
  }
  slide.addShape(pptx.ShapeType.line, {
    x: 0.62,
    y: 5.55,
    w: 12.05,
    h: 0,
    line: { color: 'D6E4E9', width: 1 },
  })
  text(slide, note, 0.62, 5.76, 11.85, 0.77, 15)
  text(
    slide,
    '来源：Deardorff et al., PLOS ONE 2020, CC BY 4.0；示意图：WisWork 自制',
    0.62,
    7.06,
    12,
    0.22,
    8,
    '58717D',
  )
}
const deckPath = join(root, 'wiswork-image-dense-research-draft.pptx')
await pptx.writeFile({ fileName: deckPath })
hashes.push(
  `${createHash('sha256')
    .update(await readFile(deckPath))
    .digest('hex')}  wiswork-image-dense-research-draft.pptx`,
)
hashes.push(
  `${createHash('sha256')
    .update(await readFile(join(root, 'deardorff-2020-article.pdf')))
    .digest('hex')}  deardorff-2020-article.pdf`,
)
await writeFile(join(root, 'SHA256SUMS'), `${hashes.join('\n')}\n`)
