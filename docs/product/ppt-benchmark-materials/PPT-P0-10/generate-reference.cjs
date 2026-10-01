const { readFile, writeFile } = require('node:fs/promises')
const { join } = require('node:path')
const { compilePresentationDeck } = require('@wiswork/pptx-engine/presentation-compiler')
const {
  parsePresentationPlan,
  presentationPlanClaims,
  assertDeckMatchesPresentationPlan,
} = require('@wiswork/pptx-engine/presentation-plan')
async function main() {
  const basis = JSON.parse(await readFile(join(__dirname, 'basis.json'), 'utf8'))
  const plan = parsePresentationPlan(
    JSON.parse(await readFile(join(__dirname, 'reference-plan.json'), 'utf8')),
  )
  const deck = {
    version: 1,
    id: plan.projectId,
    title: plan.title,
    style: plan.style,
    assets: [],
    claims: presentationPlanClaims(plan),
    slides: basis.slides.map((s, i) => ({
      id: s.id,
      title: s.title,
      claimIds: plan.slides[i].claimIds,
      notes: `${s.summary.join('\n')}\n限定：${s.limitations.join('；')}\n${s.locator}\n全文见 ${basis.source.officialUrl}；候选摘要，范围未经人工确认，专业、版权及 PowerPoint 宿主验收待完成。`,
      elements: [
        { kind: 'text', id: 'title', x: 0.65, y: 0.4, w: 12, h: 0.75, text: s.title, fontSize: 25 },
        ...s.summary.map((text, j) => ({
          kind: 'text',
          id: `summary-${j}`,
          x: 0.85,
          y: 1.35 + j * 1.12,
          w: 11.6,
          h: 1.05,
          text,
          fontSize: 18,
        })),
        {
          kind: 'text',
          id: 'limitations',
          x: 0.85,
          y: 4.85,
          w: 11.6,
          h: 1.4,
          text: `限定与下一步：${s.limitations.join('；')}`,
          fontSize: 15,
        },
        {
          kind: 'text',
          id: 'source',
          x: 0.85,
          y: 6.3,
          w: 11.6,
          h: 0.25,
          text: s.locator,
          fontSize: 10,
        },
        {
          kind: 'text',
          id: 'status',
          role: 'decoration',
          x: 0.85,
          y: 6.67,
          w: 11.6,
          h: 0.2,
          text: `${i + 1}/8 · 自制候选摘要 · 专业/版权/人工范围/宿主验收未完成`,
          fontSize: 9,
        },
      ],
    })),
  }
  assertDeckMatchesPresentationPlan(deck, plan)
  const result = await compilePresentationDeck(deck)
  await writeFile(join(__dirname, 'reference-deck.json'), JSON.stringify(deck, null, 2) + '\n')
  await writeFile(join(__dirname, 'p0-10-reference.pptx'), result.bytes)
  console.log(`8 editable slides; ${result.bytes.length} bytes; candidate only`)
}
main().catch((e) => {
  console.error(e)
  process.exitCode = 1
})
