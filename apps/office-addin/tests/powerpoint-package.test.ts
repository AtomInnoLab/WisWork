import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  editPowerPointPackage,
  verifyImportedPowerPointPackage,
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
    const extraPart = await JSZip.loadAsync(edit.base64, { base64: true })
    extraPart.file('ppt/media/unexpected.bin', 'unexpected')
    expect(
      await verifyPowerPointPackage(await extraPart.generateAsync({ type: 'base64' }), edit),
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
    await expect(
      editPowerPointPackage(input, 'master', [
        {
          path: 'ppt/slideMasters/slideMaster1.xml',
          xml: '<p:sldMaster xmlns:p="urn:p" xmlns:r="urn:r"><p:sldLayoutIdLst><p:sldLayoutId r:id="rId9"/></p:sldLayoutIdLst></p:sldMaster>',
        },
      ]),
    ).rejects.toThrow('office_api_unsupported')
  })

  it('semantically verifies a host-normalized background-only master import', async () => {
    const input = await fixture({
      'ppt/slideMasters/slideMaster1.xml':
        '<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree/></p:cSld></p:sldMaster>',
      'docProps/core.xml': '<core modified="before"/>',
    })
    const black =
      '<p:sldMaster xmlns:a="urn:a" xmlns:p="urn:p"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></p:bgPr></p:bg><p:spTree/></p:cSld></p:sldMaster>'
    const edit = await editPowerPointPackage(input, 'master', [
      { path: 'ppt/slideMasters/slideMaster1.xml', xml: black },
    ])
    const normalized = await JSZip.loadAsync(edit.base64, { base64: true })
    normalized.file(
      'ppt/slideMasters/slideMaster1.xml',
      '<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a">\n<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></p:bgPr></p:bg><p:spTree/></p:cSld></p:sldMaster>',
    )
    expect(
      await verifyImportedPowerPointPackage(
        await normalized.generateAsync({ type: 'base64' }),
        edit,
      ),
    ).toBe(true)
    normalized.file('docProps/core.xml', '<core modified="after-host-import"/>')
    expect(
      await verifyImportedPowerPointPackage(
        await normalized.generateAsync({ type: 'base64' }),
        edit,
      ),
    ).toBe(false)
    normalized.file('docProps/core.xml', '<core modified="before"/>')
    normalized.file('ppt/slideMasters/slideMaster1.xml', black.replace('000000', 'FFFFFF'))
    expect(
      await verifyImportedPowerPointPackage(
        await normalized.generateAsync({ type: 'base64' }),
        edit,
      ),
    ).toBe(false)
  })

  it('does not ignore non-background changes bundled with a background edit', async () => {
    const input = await fixture({
      'ppt/slideMasters/slideMaster1.xml':
        '<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld name="before"><p:spTree/></p:cSld></p:sldMaster>',
    })
    const edit = await editPowerPointPackage(input, 'master', [
      {
        path: 'ppt/slideMasters/slideMaster1.xml',
        xml: '<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld name="after"><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></p:bgPr></p:bg><p:spTree/></p:cSld></p:sldMaster>',
      },
    ])
    const incomplete = await JSZip.loadAsync(edit.base64, { base64: true })
    incomplete.file(
      'ppt/slideMasters/slideMaster1.xml',
      '<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld name="before"><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></p:bgPr></p:bg><p:spTree/></p:cSld></p:sldMaster>',
    )
    expect(
      await verifyImportedPowerPointPackage(
        await incomplete.generateAsync({ type: 'base64' }),
        edit,
      ),
    ).toBe(false)
  })

  it('rejects non-background mutations in an otherwise valid background import', async () => {
    const input = await fixture({
      'ppt/slideMasters/slideMaster1.xml':
        '<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp/></p:spTree></p:cSld></p:sldMaster>',
    })
    const edit = await editPowerPointPackage(input, 'master', [
      {
        path: 'ppt/slideMasters/slideMaster1.xml',
        xml: '<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:sp/></p:spTree></p:cSld></p:sldMaster>',
      },
    ])
    const mutated = await JSZip.loadAsync(edit.base64, { base64: true })
    mutated.file(
      'ppt/slideMasters/slideMaster1.xml',
      '<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></p:bgPr></p:bg><p:spTree/></p:cSld></p:sldMaster>',
    )
    expect(
      await verifyImportedPowerPointPackage(await mutated.generateAsync({ type: 'base64' }), edit),
    ).toBe(false)
  })

  it('preserves whitespace-only text when OOXML marks it as meaningful', async () => {
    const input = await fixture({
      'ppt/slideMasters/slideMaster1.xml':
        '<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><a:t xml:space="preserve"> </a:t></p:spTree></p:cSld></p:sldMaster>',
    })
    const edit = await editPowerPointPackage(input, 'master', [
      {
        path: 'ppt/slideMasters/slideMaster1.xml',
        xml: '<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></p:bgPr></p:bg><p:spTree><a:t xml:space="preserve"> </a:t></p:spTree></p:cSld></p:sldMaster>',
      },
    ])
    const mutated = await JSZip.loadAsync(edit.base64, { base64: true })
    mutated.file(
      'ppt/slideMasters/slideMaster1.xml',
      '<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></p:bgPr></p:bg><p:spTree><a:t xml:space="preserve"></a:t></p:spTree></p:cSld></p:sldMaster>',
    )
    expect(
      await verifyImportedPowerPointPackage(await mutated.generateAsync({ type: 'base64' }), edit),
    ).toBe(false)
  })

  it('keeps generic imported XML edits byte-exact', async () => {
    const input = await fixture()
    const edit = await editPowerPointPackage(input, 'slide', [
      {
        path: 'ppt/slides/slide1.xml',
        xml: '<p:sld xmlns:p="urn:p"><p:cSld name="new"/></p:sld>',
      },
    ])
    const reformatted = await JSZip.loadAsync(edit.base64, { base64: true })
    reformatted.file(
      'ppt/slides/slide1.xml',
      '<p:sld xmlns:p="urn:p">\n<p:cSld name="new"/></p:sld>',
    )
    expect(
      await verifyImportedPowerPointPackage(
        await reformatted.generateAsync({ type: 'base64' }),
        edit,
      ),
    ).toBe(false)
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
