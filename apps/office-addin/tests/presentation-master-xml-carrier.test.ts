import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { deriveMasterXmlCarrier } from '../src/skills/powerpoint/presentation-master-xml-package.js'
const ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
async function fixture() {
  const zip = new JSZip(),
    rel = (items: [string, string, string][]) =>
      `<Relationships>${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${ns + type}" Target="${target}"/>`).join('')}</Relationships>`
  zip.file(
    '[Content_Types].xml',
    '<Types><Default Extension="xml" ContentType="application/xml"/></Types>',
  )
  zip.file(
    'ppt/presentation.xml',
    '<p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="m1"/><p:sldMasterId id="2147483649" r:id="m2"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="512" r:id="s"/></p:sldIdLst></p:presentation>',
  )
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    rel([
      ['m1', 'slideMaster', 'slideMasters/first.xml'],
      ['m2', 'slideMaster', 'slideMasters/second.xml'],
      ['s', 'slide', 'slides/slide7.xml'],
    ]),
  )
  zip.file(
    'ppt/slides/slide7.xml',
    '<p:sld xmlns:p="urn:p"><p:cSld name="untouched payload"/></p:sld>',
  )
  zip.file(
    'ppt/slides/_rels/slide7.xml.rels',
    rel([
      ['layout', 'slideLayout', '../slideLayouts/first.xml'],
      ['media', 'image', '../media/linked.bin'],
      ['notes', 'notesSlide', '../notesSlides/notes.xml'],
    ]),
  )
  zip.file('ppt/media/linked.bin', Uint8Array.of(1, 2, 3))
  zip.file('ppt/media/opaque-orphan.bin', Uint8Array.of(9, 8, 7))
  zip.file(
    'ppt/notesSlides/notes.xml',
    '<p:notes xmlns:p="urn:p"><p:cSld>protected notes</p:cSld></p:notes>',
  )
  zip.file('docProps/core.xml', '<core> untouched </core>')
  zip.file('ppt/theme/theme1.xml', '<a:theme xmlns:a="urn:a" name="shared"/>')
  for (const [master, layouts] of [
    ['first', ['first']],
    ['second', ['second-a', 'second-b']],
  ] as const) {
    zip.file(
      `ppt/slideMasters/${master}.xml`,
      `<p:sldMaster xmlns:p="urn:p" xmlns:r="urn:r"><p:cSld name="duplicate"/><p:sldLayoutIdLst>${layouts.map((_, i) => `<p:sldLayoutId id="${801 + i}" r:id="l${i}"/>`).join('')}</p:sldLayoutIdLst></p:sldMaster>`,
    )
    zip.file(
      `ppt/slideMasters/_rels/${master}.xml.rels`,
      rel([
        ...layouts.map(
          (name, i) =>
            [`l${i}`, 'slideLayout', `../slideLayouts/${name}.xml`] as [string, string, string],
        ),
        ['theme', 'theme', '../theme/theme1.xml'],
      ]),
    )
    for (const name of layouts) {
      zip.file(
        `ppt/slideLayouts/${name}.xml`,
        '<p:sldLayout xmlns:p="urn:p"><p:cSld name="duplicate"/></p:sldLayout>',
      )
      zip.file(
        `ppt/slideLayouts/_rels/${name}.xml.rels`,
        rel([['owner', 'slideMaster', `../slideMasters/${master}.xml`]]),
      )
    }
  }
  return { zip, base64: await zip.generateAsync({ type: 'base64' }) }
}
const target = 'ppt/slideMasters/second.xml',
  chosen = 'ppt/slideLayouts/second-b.xml'
it('derives a secondary-master carrier by ordered OOXML membership preserving source identity and all other raw entries', async () => {
  const f = await fixture(),
    result = await deriveMasterXmlCarrier(f.base64, target, chosen)
  expect(result).toMatchObject({
    sourceSlideId: '512#',
    sourceMasterPath: target,
    sourceLayoutPath: chosen,
  })
  expect(result.inventory.sourceMasterPath).toBe(target)
  const actual = await JSZip.loadAsync(result.base64, { base64: true })
  const paths = Object.keys(f.zip.files)
    .filter((p) => !f.zip.files[p]!.dir)
    .sort()
  expect(
    Object.keys(actual.files)
      .filter((p) => !actual.files[p]!.dir)
      .sort(),
  ).toEqual(paths)
  for (const path of paths)
    if (path !== 'ppt/slides/_rels/slide7.xml.rels')
      expect(await actual.file(path)!.async('uint8array')).toEqual(
        await f.zip.file(path)!.async('uint8array'),
      )
  expect(await actual.file('ppt/slides/_rels/slide7.xml.rels')!.async('string')).toBe(
    (await f.zip.file('ppt/slides/_rels/slide7.xml.rels')!.async('string')).replace(
      'Target="../slideLayouts/first.xml"',
      'Target="../slideLayouts/second-b.xml"',
    ),
  )
  expect(result.inventory.masters.map((m) => m.contentDigest)).toEqual(
    (await deriveMasterXmlCarrier(f.base64, 'ppt/slideMasters/first.xml')).inventory.masters.map(
      (m) => m.contentDigest,
    ),
  )
})
it('uses the selected master first relationship-listed layout as default and preserves an existing carrier unchanged', async () => {
  const f = await fixture()
  expect((await deriveMasterXmlCarrier(f.base64, target)).sourceLayoutPath).toBe(
    'ppt/slideLayouts/second-a.xml',
  )
  expect((await deriveMasterXmlCarrier(f.base64, 'ppt/slideMasters/first.xml')).base64).toBe(
    f.base64,
  )
})
it.each([
  [target, 'ppt/slideLayouts/first.xml'],
  ['ppt/slideMasters/missing.xml', undefined],
  [target, '../escape.xml'],
])('rejects unknown or foreign membership %s %s', async (master, layout) => {
  const f = await fixture()
  await expect(deriveMasterXmlCarrier(f.base64, master!, layout)).rejects.toThrow(
    'presentation_master_xml_package_unproven',
  )
})
it.each(['external', 'conflict'])('rejects %s source layout relationships', async (kind) => {
  const f = await fixture(),
    path = 'ppt/slides/_rels/slide7.xml.rels',
    xml = await f.zip.file(path)!.async('string')
  f.zip.file(
    path,
    kind === 'external'
      ? xml.replace('Id="layout"', 'Id="layout" TargetMode="External"')
      : xml.replace(
          '</Relationships>',
          `<Relationship Id="duplicate" Type="${ns}slideLayout" Target="../slideLayouts/first.xml"/></Relationships>`,
        ),
  )
  await expect(
    deriveMasterXmlCarrier(await f.zip.generateAsync({ type: 'base64' }), target),
  ).rejects.toThrow('presentation_master_xml_package_unproven')
})
it('rejects runtime object aliases, cancellation and ZIP entry bounds', async () => {
  const f = await fixture(),
    alias = { toString: () => target }
  await expect(deriveMasterXmlCarrier(f.base64, alias as unknown as string)).rejects.toThrow(
    'presentation_master_xml_package_unproven',
  )
  const abort = new AbortController()
  abort.abort()
  await expect(deriveMasterXmlCarrier(f.base64, target, undefined, abort.signal)).rejects.toThrow(
    'cancelled',
  )
  f.zip.file('ppt/media/oversize.bin', new Uint8Array(2 * 1024 * 1024 + 1))
  await expect(
    deriveMasterXmlCarrier(await f.zip.generateAsync({ type: 'base64' }), target),
  ).rejects.toThrow('presentation_master_xml_package_unproven')
})
it('changes exactly the real layout Target while preserving quoted greater-than URLs and other relationship spelling', async () => {
  const f = await fixture(),
    path = 'ppt/slides/_rels/slide7.xml.rels'
  const raw = (await f.zip.file(path)!.async('string'))
    .replace('Target="../slideLayouts/first.xml"', "Target = '../slideLayouts/first.xml'")
    .replace(
      '</Relationships>',
      `<Relationship Target="https://example.org/a>z" Type="${ns}hyperlink" Id="external" TargetMode="External"/></Relationships>`,
    )
  f.zip.file(path, raw)
  const result = await deriveMasterXmlCarrier(
    await f.zip.generateAsync({ type: 'base64' }),
    target,
    chosen,
  )
  const actual = await JSZip.loadAsync(result.base64, { base64: true })
  expect(await actual.file(path)!.async('string')).toBe(
    raw.replace("Target = '../slideLayouts/first.xml'", "Target = '../slideLayouts/second-b.xml'"),
  )
})
