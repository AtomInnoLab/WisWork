import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { replacePowerPointTextRangePackage } from '../src/skills/powerpoint/presentation-text-revision-package.js'
import { createPresentationExistingPageEditingSkill } from '../src/skills/powerpoint/presentation-existing-page-editing.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'

async function fixture() {
  const deck = benchmarkDeck()
  deck.slides = [{ ...deck.slides[0]!, claimIds: [], elements: [deck.slides[0]!.elements[0]!] }]
  const source = (await compilePresentationDeck(deck)).bytes
  const zip = await JSZip.loadAsync(source)
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const shape = [...xml.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)].find(([part]) =>
    part.includes('name="title"'),
  )![0]
  const id = /<p:cNvPr\b[^>]*\bid="(\d+)"/.exec(shape)![1]!
  return { zip, xml, shape, id }
}

it('changes a selected span across styled native runs without flattening their formatting', async () => {
  const { zip, xml, shape, id } = await fixture()
  const oldText = '科研汇报'
  const mixed = shape.replace(
    `<a:t>${oldText}</a:t>`,
    '<a:t>科研</a:t></a:r><a:r><a:rPr lang="zh-CN" b="1"/><a:t>汇报</a:t>',
  )
  expect(mixed).not.toBe(shape)
  zip.file('ppt/slides/slide1.xml', xml.replace(shape, mixed))
  const source = await zip.generateAsync({ type: 'base64' })
  const revised = await replacePowerPointTextRangePackage(source, id, 1, '研汇', '究展')
  expect(revised.changedRuns).toBe(2)
  const output = await JSZip.loadAsync(revised.base64, { base64: true })
  const after = await output.file('ppt/slides/slide1.xml')!.async('string')
  expect(after).toContain('<a:t>科究</a:t>')
  expect(after).toContain('<a:rPr lang="zh-CN" b="1"/><a:t>展报</a:t>')
  expect(
    after
      .replace('<a:t>科究</a:t>', '<a:t>科研</a:t>')
      .replace('<a:t>展报</a:t>', '<a:t>汇报</a:t>'),
  ).toBe(xml.replace(shape, mixed))
  expect(revised.beforeDigest).not.toBe(revised.afterDigest)
})

it('supports a length-changing edit inside one native run and rejects stale text', async () => {
  const { zip, id } = await fixture()
  const source = await zip.generateAsync({ type: 'base64' })
  await expect(replacePowerPointTextRangePackage(source, id, 0, '错误', '替换')).rejects.toThrow(
    'presentation_baseline_changed',
  )
  const revised = await replacePowerPointTextRangePackage(source, id, 0, '科研', '研究组')
  expect(revised.changedRuns).toBe(1)
  const output = await JSZip.loadAsync(revised.base64, { base64: true })
  expect(await output.file('ppt/slides/slide1.xml')!.async('string')).toContain(
    '<a:t>研究组汇报</a:t>',
  )
})

it('rejects a length-changing edit across differently formatted runs', async () => {
  const { zip, xml, shape, id } = await fixture()
  const mixed = shape.replace(
    '<a:t>科研汇报</a:t>',
    '<a:t>科研</a:t></a:r><a:r><a:rPr lang="zh-CN" b="1"/><a:t>汇报</a:t>',
  )
  zip.file('ppt/slides/slide1.xml', xml.replace(shape, mixed))
  await expect(
    replacePowerPointTextRangePackage(
      await zip.generateAsync({ type: 'base64' }),
      id,
      1,
      '研汇',
      '主题页',
    ),
  ).rejects.toThrow('presentation_existing_target_unsupported')
})

it('escapes replacement text and rejects field-backed text', async () => {
  const { zip, xml, shape, id } = await fixture()
  zip.file('ppt/slides/slide1.xml', xml.replace(shape, shape.replace('科研汇报', 'A&amp;B汇报')))
  const source = await zip.generateAsync({ type: 'base64' })
  const revised = await replacePowerPointTextRangePackage(source, id, 1, '&', '<')
  const output = await JSZip.loadAsync(revised.base64, { base64: true })
  expect(await output.file('ppt/slides/slide1.xml')!.async('string')).toContain('A&lt;B汇报')
  const field = shape.replace('<a:t>科研汇报</a:t>', '<a:fld><a:t>科研汇报</a:t></a:fld>')
  zip.file('ppt/slides/slide1.xml', xml.replace(shape, field))
  await expect(
    replacePowerPointTextRangePackage(
      await zip.generateAsync({ type: 'base64' }),
      id,
      0,
      '科研',
      '项目',
    ),
  ).rejects.toThrow('presentation_existing_target_unsupported')
})

it('prepares a VFS revision from a fresh page baseline without writing PowerPoint', async () => {
  const { zip, id } = await fixture()
  const source = await zip.generateAsync({ type: 'base64' })
  const vfs = new InMemoryVfs()
  let writes = 0
  const baseline = {
    baselineId: 'baseline',
    documentId: 'doc',
    scope: { kind: 'current', slideIds: ['host-slide'] },
    context: { slideIds: ['host-slide'], selectedSlideIds: ['host-slide'], selectedShapeIds: [] },
  }
  const skill = createPresentationExistingPageEditingSkill({
    baseline: {
      snapshot: () => baseline,
      executeTool: async () => ({ output: '{"unchanged":true}', mutated: false }),
    },
    adapter: {
      stage: async () => {
        writes++
      },
    },
    inspectPage: async () => {
      throw new Error('unexpected inspection')
    },
    exportAdapter: {
      exportPresentationPagePackage: async () => ({
        slideId: 'host-slide',
        slideIds: ['host-slide'],
        base64: source,
      }),
    },
    vfs,
    request: async () => {
      throw new Error('unexpected request')
    },
    proposals: createStructuredProposalController(),
    documentId: async () => 'doc',
    readExistingPageChange: () => undefined,
    writeExistingPageChange: async () => {
      writes++
    },
    available: () => true,
  } as unknown as Parameters<typeof createPresentationExistingPageEditingSkill>[0])
  const result = await skill.executeTool({
    id: 'prepare',
    name: 'prepare_existing_presentation_text_revision',
    input: {
      baseline_id: 'baseline',
      slide_id: 'host-slide',
      shape_id: id,
      start: 0,
      before: '科研',
      after: '项目',
    },
  })
  expect(result.isError, result.output).not.toBe(true)
  const prepared = JSON.parse(result.output) as {
    path: string
    changedRuns: number
    nextTool: string
  }
  expect(prepared.nextTool).toBe('stage_existing_presentation_page_change')
  expect(prepared.changedRuns).toBe(1)
  const revised = Buffer.from(vfs.readBytes(prepared.path, { maxBytes: 8 * 1024 * 1024 })).toString(
    'base64',
  )
  const output = await JSZip.loadAsync(revised, { base64: true })
  expect(await output.file('ppt/slides/slide1.xml')!.async('string')).toContain(
    '<a:t>项目汇报</a:t>',
  )
  expect(writes).toBe(0)
})
