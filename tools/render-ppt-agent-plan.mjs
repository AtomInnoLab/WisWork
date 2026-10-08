import fs from 'node:fs/promises'
import path from 'node:path'
import { marked } from 'marked'

const root = process.cwd()
const input = path.join(
  root,
  'docs/product/wiswork-ppt-agent-solution-and-implementation-plan-2026-09-22.md',
)
const output = path.join(
  root,
  'output/pdf/wiswork-ppt-agent-solution-plan-2026-09-22.html',
)
const assetSource = path.join(root, 'docs/product/assets')
const assetOutput = path.join(root, 'output/pdf/assets')

const markdown = await fs.readFile(input, 'utf8')
const rendered = await marked.parse(markdown, { gfm: true })
const body = rendered.replace(/<p>(<img [^>]+>)<\/p>/g, '<div class="figure">$1</div>')

const html = `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>WisWork PPT Agent 完整方案与实施计划</title>
  <style>
    @page { size: A4; margin: 17mm 16mm 18mm; }
    * { box-sizing: border-box; }
    html { color: #18202a; background: #fff; }
    body {
      margin: 0;
      font-family: "Noto Sans CJK SC", "Source Han Sans SC", "Microsoft YaHei", sans-serif;
      font-size: 10pt;
      line-height: 1.62;
    }
    h1, h2, h3, h4 { break-after: avoid; page-break-after: avoid; color: #12243d; }
    h1 {
      margin: 0 0 9mm;
      padding: 27mm 12mm 13mm;
      color: #fff;
      background: #102a43;
      font-size: 26pt;
      line-height: 1.25;
      letter-spacing: .02em;
      border-bottom: 5px solid #ff7043;
    }
    h2 {
      margin: 10mm 0 4mm;
      padding-bottom: 2mm;
      font-size: 17pt;
      border-bottom: 1px solid #aebccc;
    }
    h3 { margin: 7mm 0 2.5mm; font-size: 13pt; color: #1f5f8b; }
    h4 { margin: 5mm 0 2mm; font-size: 11pt; }
    p { margin: 0 0 3.2mm; }
    blockquote {
      margin: 0 0 7mm;
      padding: 4mm 5mm;
      color: #354a5f;
      background: #edf4f8;
      border-left: 4px solid #2f80a8;
    }
    blockquote p { margin: 0 0 1mm; }
    ul, ol { margin: 1.5mm 0 4mm 6mm; padding-left: 5mm; }
    li { margin: 0 0 1.2mm; }
    strong { color: #102a43; }
    a { color: #1d6f9f; text-decoration: none; }
    hr { margin: 9mm 0; border: 0; border-top: 1px solid #c9d4df; }
    table {
      width: 100%;
      margin: 3mm 0 6mm;
      border-collapse: collapse;
      table-layout: fixed;
      font-size: 8.2pt;
      line-height: 1.43;
    }
    thead { display: table-header-group; }
    tr { break-inside: avoid; page-break-inside: avoid; }
    th, td {
      padding: 2.2mm 2.5mm;
      vertical-align: top;
      border: 1px solid #c5d0da;
      overflow-wrap: anywhere;
    }
    th { color: #fff; background: #244b68; text-align: left; }
    tbody tr:nth-child(even) { background: #f5f8fa; }
    pre {
      margin: 3mm 0 6mm;
      padding: 4mm;
      color: #eaf2f8;
      background: #152532;
      border-radius: 3px;
      font-family: "Noto Sans Mono CJK SC", "SFMono-Regular", monospace;
      font-size: 8pt;
      line-height: 1.48;
      white-space: pre-wrap;
      overflow-wrap: anywhere;
      break-inside: avoid;
    }
    code {
      font-family: "Noto Sans Mono CJK SC", "SFMono-Regular", monospace;
      font-size: .92em;
      color: #8a3d22;
      background: #f6eee9;
      padding: .1em .28em;
      border-radius: 2px;
    }
    pre code { color: inherit; background: transparent; padding: 0; }
    .figure {
      margin: 5mm 0 8mm;
      padding: 3mm 4mm 7mm;
      background: #f8fafb;
      border: 1px solid #e3e9ed;
      border-radius: 3px;
      line-height: 0;
      overflow: visible;
      clear: both;
      break-inside: avoid;
      page-break-inside: avoid;
    }
    .figure img {
      display: block;
      width: 100%;
      height: auto;
      max-height: none;
      object-fit: scale-down;
      margin: 0 auto;
      padding: 0 0 2mm;
      overflow: visible;
      break-inside: avoid;
      page-break-inside: avoid;
    }
    h2:nth-of-type(1) { margin-top: 5mm; }
    @media print {
      a { color: inherit; }
      h2 { break-before: auto; }
    }
  </style>
</head>
<body>${body}</body>
</html>`

await fs.mkdir(path.dirname(output), { recursive: true })
await fs.mkdir(assetOutput, { recursive: true })
for (const file of await fs.readdir(assetSource)) {
  if (file.endsWith('.svg') || file.endsWith('.png')) {
    await fs.copyFile(path.join(assetSource, file), path.join(assetOutput, file))
  }
}
await fs.writeFile(output, html)
console.log(output)
