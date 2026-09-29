import { describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import type { PowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter.js'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill.js'
import type {
  StructuredProposalController,
  StructuredProposalRequest,
} from '../src/agent/proposal-controller.js'
import { presentationPackageDigest } from '../src/skills/powerpoint/powerpoint-package.js'

const cases = [
  [
    'edit_slide_xml',
    'ppt/slides/slide1.xml',
    '<p:sld xmlns:p="urn:p"/>',
    '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>',
  ],
  [
    'edit_slide_chart',
    'ppt/charts/chart1.xml',
    '<c:chart xmlns:c="urn:c"/>',
    '<c:chart xmlns:c="urn:c"><c:title/></c:chart>',
  ],
  [
    'edit_slide_master_xml',
    'ppt/slideMasters/slideMaster1.xml',
    '<p:sldMaster xmlns:p="urn:p"/>',
    '<p:sldMaster xmlns:p="urn:p"><p:cSld/></p:sldMaster>',
  ],
] as const

async function fixture(name: string, path: string, before: string, after: string) {
  const zip = new JSZip()
  zip.file(path, before)
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships/>')
  zip.file('docProps/core.xml', '<core value="original"/>')
  const base64 = await zip.generateAsync({ type: 'base64' })
  const observed = { slideId: 'source', base64, fingerprint: 'volatile' }
  const order = Array.from({ length: 600 }, (_, index) => (index === 599 ? 'source' : `s${index}`))
  const readSlideOrder = vi.fn(async () => order)
  let request!: StructuredProposalRequest
  const replaceSlidePackage = vi.fn(async () => ({ slideId: 'inserted' }))
  const adapter = {
    verifySlides: vi.fn(async () => ({ slides: [] })),
    exportSlidePackage: vi.fn(async () => observed),
    readSlideOrder,
    replaceSlidePackage,
  } as unknown as PowerPointAdapter
  const proposals = {
    propose: vi.fn((value: StructuredProposalRequest) => {
      request = value
      return { ...value, id: 'proposal' }
    }),
  } as unknown as StructuredProposalController
  const slideIndex = name === 'edit_slide_master_xml' ? 0 : 599
  if (slideIndex === 0) [order[0], order[599]] = [order[599]!, order[0]!]
  const skill = createPowerPointSkill({ adapter, proposals })
  await skill.executeTool({
    id: 'call',
    name,
    input: {
      ...(name === 'edit_slide_master_xml' ? {} : { slide_index: slideIndex }),
      program: { version: 1, operations: [{ op: 'replace_xml', path, xml: after }] },
    },
  })
  return { zip, base64, observed, order, adapter, replaceSlidePackage, request, slideIndex }
}

describe.each(cases)('%s exact package proposal', (name, path, before, after) => {
  it('keeps unchanged full order and passes exact preimage without narrowing large slide indices', async () => {
    const f = await fixture(name, path, before, after)
    expect(await f.request.validate()).toBe(true)
    await f.request.execute()
    expect(f.replaceSlidePackage).toHaveBeenCalledWith(
      f.slideIndex,
      expect.any(String),
      name === 'edit_slide_master_xml',
      expect.any(Object),
      undefined,
      {
        slideId: 'source',
        packageDigest: await presentationPackageDigest(f.base64),
        slideIds: [...f.order],
      },
    )
  })

  it.each(['unrelated', 'structural', 'identity', 'order'])(
    'rejects %s drift in validation and execution before mutation',
    async (drift) => {
      const f = await fixture(name, path, before, after)
      if (drift === 'identity') f.observed.slideId = 'other'
      else if (drift === 'order') [f.order[2], f.order[3]] = [f.order[3]!, f.order[2]!]
      else {
        f.zip.file(
          drift === 'unrelated' ? 'docProps/core.xml' : 'ppt/slides/_rels/slide1.xml.rels',
          drift === 'unrelated'
            ? '<core value="changed"/>'
            : '<Relationships><Relationship Id="r1" Target="other.xml"/></Relationships>',
        )
        f.observed.base64 = await f.zip.generateAsync({ type: 'base64' })
      }
      expect(await f.request.validate()).toBe(false)
      await expect(f.request.execute()).rejects.toThrow('proposal_stale')
      expect(f.replaceSlidePackage).not.toHaveBeenCalled()
    },
  )

  it('copies the observed object before later awaited order reads can mutate its aliases', async () => {
    const f = await fixture(name, path, before, after)
    vi.mocked(f.adapter.readSlideOrder!).mockImplementationOnce(async () => {
      f.observed.slideId = 'other'
      return f.order
    })
    expect(await f.request.validate()).toBe(true)
    await expect(f.request.execute()).rejects.toThrow('proposal_stale')
    expect(f.replaceSlidePackage).not.toHaveBeenCalled()
  })
})
