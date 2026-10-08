import { expect, it, vi } from 'vitest'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill.js'
import type { PowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
const input = {
  program: {
    version: 1,
    operations: [
      { op: 'replace_xml', path: 'ppt/slideMasters/slideMaster1.xml', xml: '<p:sldMaster/>' },
    ],
  },
}
function fixture(extra: Partial<Parameters<typeof createPowerPointSkill>[0]> = {}) {
  const adapter = {
    verifySlides: vi.fn(),
    exportSlidePackage: vi.fn(),
    replaceSlidePackage: vi.fn(),
  } as unknown as PowerPointAdapter
  const durableMasterXml = vi.fn(async () => ({ id: 'master-proposal' }))
  const skill = createPowerPointSkill({
    adapter,
    proposals: createStructuredProposalController(),
    durableMasterXml,
    ...extra,
  })
  return { adapter, durableMasterXml, skill }
}
it('routes parsed master XML to the durable transaction without a legacy package write', async () => {
  const f = fixture(),
    signal = new AbortController().signal
  const result = await f.skill.executeTool(
    { id: 'call', name: 'edit_slide_master_xml', input },
    signal,
  )
  expect(result.isError, result.output).not.toBe(true)
  expect(f.durableMasterXml).toHaveBeenCalledWith(
    [{ path: 'ppt/slideMasters/slideMaster1.xml', xml: '<p:sldMaster/>' }],
    undefined,
    signal,
  )
  expect(f.adapter.exportSlidePackage).not.toHaveBeenCalled()
  expect(f.adapter.replaceSlidePackage).not.toHaveBeenCalled()
})
it('refuses missing durable master XML support before reading or writing the host', async () => {
  const f = fixture({ durableMasterXml: undefined })
  const result = await f.skill.executeTool({ id: 'call', name: 'edit_slide_master_xml', input })
  expect(result).toMatchObject({
    isError: true,
    mutated: false,
    output: 'presentation_master_xml_persistence_unavailable',
  })
  expect(f.adapter.verifySlides).not.toHaveBeenCalled()
  expect(f.adapter.exportSlidePackage).not.toHaveBeenCalled()
  expect(f.adapter.replaceSlidePackage).not.toHaveBeenCalled()
})
it('keeps Mac exclusion before invoking the master transaction', async () => {
  const f = fixture({ platform: 'mac' })
  expect(f.skill.tools.some((t) => t.name === 'edit_slide_master_xml')).toBe(false)
  const result = await f.skill.executeTool({ id: 'call', name: 'edit_slide_master_xml', input })
  expect(result).toMatchObject({ isError: true, mutated: false, output: 'office_api_unsupported' })
  expect(f.durableMasterXml).not.toHaveBeenCalled()
})
it.each([
  'PowerPoint.run(...)',
  { version: 1, operations: [{ op: 'delete_slide', slide_index: 0 }] },
])('refuses invalid master programs before the durable transaction', async (program) => {
  const f = fixture()
  const result = await f.skill.executeTool({
    id: 'call',
    name: 'edit_slide_master_xml',
    input: typeof program === 'string' ? { code: program } : { program },
  })
  expect(result).toMatchObject({ isError: true, output: 'invalid_tool_input' })
  expect(f.durableMasterXml).not.toHaveBeenCalled()
})
