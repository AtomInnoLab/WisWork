import { readFileSync } from 'node:fs'
import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { openPptx } from '@wiswork/pptx-engine'
import { replacePowerPointTextRangePackage } from '../src/skills/powerpoint/presentation-text-revision-package.js'
import { replacePowerPointShapeGeometryPackage } from '../src/skills/powerpoint/presentation-geometry-revision-package.js'

const materials = new URL(
  '../../../docs/product/ppt-benchmark-materials/PPT-P0-15/',
  import.meta.url,
)
const scenario = JSON.parse(readFileSync(new URL('scenario.json', materials), 'utf8'))

for (const file of [
  'wiswork-image-dense-research-draft.pptx',
  'wiswork-generated-candidate.pptx',
] as const) {
  it(`edits only two native P0-15 objects and rejects a stale caption in ${file}`, async () => {
    const archive = await JSZip.loadAsync(readFileSync(new URL(file, materials)))
    for (let page = 1; page <= 8; page++) {
      if (page === 4) continue
      archive.remove(`ppt/slides/slide${page}.xml`)
      archive.remove(`ppt/slides/_rels/slide${page}.xml.rels`)
    }
    const presentation = await archive.file('ppt/presentation.xml')!.async('string')
    archive.file(
      'ppt/presentation.xml',
      presentation.replace(/<p:sldId\b[^>]*\/>/g, (entry) =>
        entry.includes('r:id="rId5"') ? entry : '',
      ),
    )
    const source = await archive.generateAsync({ type: 'base64' })
    const slideBefore = await archive.file('ppt/slides/slide4.xml')!.async('string')
    const objects = (xml: string) =>
      [...xml.matchAll(/<p:(sp|pic|graphicFrame|cxnSp)\b[^>]*>[\s\S]*?<\/p:\1>/g)].map(
        ([item]) => item,
      )
    const id = (xml: string) => /<p:cNvPr\b[^>]*\bid="(\d+)"/.exec(xml)?.[1]
    const titleId = id(
      objects(slideBefore).find((shape) =>
        shape.includes(`<a:t>${scenario.existingDeckEdits[0].before}</a:t>`),
      )!,
    )!
    const caption = objects(slideBefore).find((shape) =>
      shape.includes(
        file.startsWith('wiswork-generated') ? '招募、工作坊、三个月后访谈' : '自制示意图 07',
      ),
    )!
    const captionId = id(caption)!
    const off = /<a:off x="(\d+)" y="(\d+)"\/>/.exec(caption)!
    const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(caption)!
    const before = {
      left: Number(off[1]) / 12700,
      top: Number(off[2]) / 12700,
      width: Number(ext[1]) / 12700,
      height: Number(ext[2]) / 12700,
    }
    const manual = await replacePowerPointShapeGeometryPackage(source, captionId, before, {
      ...before,
      top: before.top + 0.04 * 72,
    })
    await expect(
      replacePowerPointShapeGeometryPackage(manual.base64, captionId, before, {
        ...before,
        top: before.top + 0.08 * 72,
      }),
    ).rejects.toThrow('presentation_baseline_changed')

    const text = await replacePowerPointTextRangePackage(
      source,
      titleId,
      0,
      scenario.existingDeckEdits[0].before,
      scenario.existingDeckEdits[0].after,
    )
    const geometry = await replacePowerPointShapeGeometryPackage(text.base64, captionId, before, {
      ...before,
      top: before.top + 0.08 * 72,
    })
    const slideAfter = await (
      await JSZip.loadAsync(geometry.base64, { base64: true })
    )
      .file('ppt/slides/slide4.xml')!
      .async('string')
    const beforeObjects = objects(slideBefore)
    const afterObjects = objects(slideAfter)
    expect(afterObjects).toHaveLength(beforeObjects.length)
    for (let index = 0; index < beforeObjects.length; index++) {
      expect(id(afterObjects[index]!)).toBe(id(beforeObjects[index]!))
      if (![titleId, captionId].includes(id(beforeObjects[index]!)!))
        expect(afterObjects[index]).toBe(beforeObjects[index])
    }
    expect(slideAfter).toContain(`<a:t>${scenario.existingDeckEdits[0].after}</a:t>`)
    expect(afterObjects.find((shape) => id(shape) === captionId)).toContain(
      `y="${Math.round((before.top + 0.08 * 72) * 12700)}"`,
    )
    expect((await openPptx(Buffer.from(geometry.base64, 'base64'))).deck.slides).toHaveLength(1)
  })
}
