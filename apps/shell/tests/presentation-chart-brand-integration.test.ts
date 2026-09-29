import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { expect, it } from 'vitest'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'

it('propagates a saved brand palette and font into actual native chart XML through production', async () => {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-chart-brand-'))
  try {
    const service = createPresentationService({ userDataPath: root })
    const plan = benchmarkPlan()
    const deck = benchmarkPlannedDeck()
    const style = {
      fontFace: 'DejaVu Sans Mono',
      background: 'E8F0D2',
      textColor: '643A71',
      accentColor: 'B35E24',
    }
    const kit = {
      id: 'chart-brand',
      revision: 1,
      name: '图表品牌',
      allowedColors: Object.values(style).filter((value) => value !== style.fontFace),
    }
    plan.style = style
    plan.brandKit = kit
    deck.style = style
    const chart = deck.slides[6]!.elements.find((element) => element.kind === 'chart')!
    if (chart.kind !== 'chart') throw new Error('missing chart fixture')
    chart.series.push({ name: '第二组（合成）', values: [80, 70] })
    const call = async (operation: string, extra: Record<string, unknown> = {}) =>
      JSON.parse(
        Buffer.from(
          await service(
            {
              operation,
              documentId: 'brand-doc',
              ...(operation.startsWith('brand_kit_') ? {} : { projectId: deck.id }),
              ...extra,
            },
            new AbortController().signal,
          ),
        ).toString('utf8'),
      )
    expect(await call('brand_kit_save', { expectedRevision: 0, brandKit: kit })).toEqual({
      brandKit: kit,
    })
    expect(await call('brand_kit_get', { brandKitId: kit.id, revision: 1 })).toEqual({
      brandKit: kit,
    })
    expect(await call('save_plan', { expectedRevision: 0, plan })).toMatchObject({ revision: 1 })
    expect(
      await call('production_begin', { requestId: 'brand-production', planRevision: 1, deck }),
    ).not.toHaveProperty('error')
    expect(await call('production_run', { requestId: 'brand-production' })).toMatchObject({
      status: 'compiled',
    })
    const page = await call('production_page', {
      requestId: 'brand-production',
      pageId: deck.slides[6]!.id,
    })
    expect(page).not.toHaveProperty('error')
    const zip = await JSZip.loadAsync(Buffer.from(page.pptxBase64, 'base64'))
    const charts = Object.keys(zip.files).filter((name) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(name),
    )
    expect(charts).toHaveLength(1)
    const xml = await zip.file(charts[0]!)!.async('string')
    expect(xml).toContain('typeface="DejaVu Sans Mono"')
    expect(xml).toContain('val="643A71"')
    expect(xml).toContain('val="E8F0D2"')
    expect(xml).toContain('val="B35E24"')
    for (const part of ['dLbls', 'legend', 'catAx', 'valAx']) {
      const body = xml.match(new RegExp(`<c:${part}\\b[^>]*>[\\s\\S]*?</c:${part}>`))?.[0]
      expect(body, part).toBeDefined()
      expect(body, part).toContain('typeface="DejaVu Sans Mono"')
      expect(body, part).toContain('val="643A71"')
    }
    const series = xml.match(/<c:ser>[\s\S]*?<\/c:ser>/g)
    expect(series).toHaveLength(2)
    for (const item of series!) expect(item).toContain('val="B35E24"')
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
