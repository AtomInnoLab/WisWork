import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  editPowerPointPackage,
  verifyPowerPointPackage,
} from '../src/skills/powerpoint/powerpoint-package.js'

async function fixture(extra: Record<string, string> = {}): Promise<string> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<Types/>')
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>')
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    '<Relationships><Relationship Id="rId1" Target="../charts/chart1.xml"/></Relationships>',
  )
  zip.file('ppt/charts/chart1.xml', '<c:chart xmlns:c="urn:c"><c:title/></c:chart>')
  zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"/>')
  for (const [path, value] of Object.entries(extra)) zip.file(path, value)
  return zip.generateAsync({ type: 'base64' })
}

describe('bounded PowerPoint package editing', () => {
  it('round-trips a slide XML replacement while preserving relationships and unrelated parts', async () => {
    const input = await fixture()
    const edit = await editPowerPointPackage(input, 'slide', [
      { path: 'ppt/slides/slide1.xml', xml: '<p:sld xmlns:p="urn:p"><p:cSld name="new"/></p:sld>' },
    ])
    expect(await verifyPowerPointPackage(edit.base64, edit)).toBe(true)
    const tampered = await JSZip.loadAsync(edit.base64, { base64: true })
    tampered.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    expect(
      await verifyPowerPointPackage(await tampered.generateAsync({ type: 'base64' }), edit),
    ).toBe(false)
    const relationshipTamper = await JSZip.loadAsync(edit.base64, { base64: true })
    relationshipTamper.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships/>')
    expect(
      await verifyPowerPointPackage(
        await relationshipTamper.generateAsync({ type: 'base64' }),
        edit,
      ),
    ).toBe(false)
    const before = await JSZip.loadAsync(input, { base64: true })
    const after = await JSZip.loadAsync(edit.base64, { base64: true })
    expect(await after.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')).toBe(
      await before.file('ppt/slides/_rels/slide1.xml.rels')!.async('string'),
    )
    expect(await after.file('ppt/charts/chart1.xml')!.async('string')).toBe(
      await before.file('ppt/charts/chart1.xml')!.async('string'),
    )
  })

  it('edits only allowlisted chart or master XML parts', async () => {
    const input = await fixture()
    await expect(
      editPowerPointPackage(input, 'chart', [
        {
          path: 'ppt/charts/chart1.xml',
          xml: '<c:chart xmlns:c="urn:c"><c:title>Q1</c:title></c:chart>',
        },
      ]),
    ).resolves.toMatchObject({ changedPaths: ['ppt/charts/chart1.xml'] })
    await expect(
      editPowerPointPackage(input, 'master', [
        {
          path: 'ppt/slideMasters/slideMaster1.xml',
          xml: '<p:sldMaster xmlns:p="urn:p"><p:cSld/></p:sldMaster>',
        },
      ]),
    ).resolves.toMatchObject({ changedPaths: ['ppt/slideMasters/slideMaster1.xml'] })
    await expect(
      editPowerPointPackage(input, 'chart', [{ path: 'ppt/slides/slide1.xml', xml: '<x/>' }]),
    ).rejects.toThrow('invalid_tool_input')
  })

  it('rejects malformed XML, traversal paths, excessive entries, and declared zip bombs', async () => {
    const input = await fixture()
    await expect(
      editPowerPointPackage(input, 'slide', [{ path: 'ppt/slides/slide1.xml', xml: '<broken>' }]),
    ).rejects.toThrow('invalid_tool_input')
    await expect(
      editPowerPointPackage(input, 'slide', [
        {
          path: 'ppt/slides/slide1.xml',
          xml: '<!DOCTYPE x [<!ENTITY leak SYSTEM "file:///etc/passwd">]><x>&leak;</x>',
        },
      ]),
    ).rejects.toThrow('invalid_tool_input')
    await expect(
      editPowerPointPackage(input, 'slide', [{ path: '../ppt/slides/slide1.xml', xml: '<x/>' }]),
    ).rejects.toThrow('invalid_tool_input')
    const many = Object.fromEntries(
      Array.from({ length: 260 }, (_, index) => [`ppt/media/${index}.txt`, 'x']),
    )
    await expect(
      editPowerPointPackage(await fixture(many), 'slide', [
        { path: 'ppt/slides/slide1.xml', xml: '<x/>' },
      ]),
    ).rejects.toThrow('invalid_tool_input')
    await expect(
      editPowerPointPackage(
        await fixture({ 'ppt/media/bomb.bin': 'x'.repeat(2 * 1024 * 1024 + 1) }),
        'slide',
        [{ path: 'ppt/slides/slide1.xml', xml: '<x/>' }],
      ),
    ).rejects.toThrow('invalid_tool_input')
  })

  it('honors cancellation before producing a replacement package', async () => {
    const controller = new AbortController()
    controller.abort()
    await expect(
      editPowerPointPackage(
        await fixture(),
        'slide',
        [{ path: 'ppt/slides/slide1.xml', xml: '<x/>' }],
        controller.signal,
      ),
    ).rejects.toThrow('cancelled')
  })
})
