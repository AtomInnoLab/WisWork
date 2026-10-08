import JSZip from 'jszip'
import { expect, it } from 'vitest'
import {
  inspectMasterXmlPackage,
  assertMasterXmlPagePreserved,
  assertMasterXmlPreparation,
  proveMasterXmlLayoutMapping,
} from '../src/skills/powerpoint/presentation-master-xml-package.js'
const ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
async function fixture(selected = 0, renumber = false) {
  const zip = new JSZip(),
    num = (n: number) => n + (renumber ? 10 : 0),
    rel = (n: number) => `r${n + (renumber ? 20 : 0)}`
  const master = (m: number) => `ppt/slideMasters/slideMaster${num(m)}.xml`,
    layout = (n: number) => `ppt/slideLayouts/slideLayout${num(n)}.xml`
  const relationships = (entries: { id: string; type: string; target: string }[]) =>
    `<Relationships>${entries.map((e) => `<Relationship Id="${e.id}" Type="${ns + e.type}" Target="${e.target}"/>`).join('')}</Relationships>`
  const putRels = (path: string, entries: { id: string; type: string; target: string }[]) =>
    zip.file(path.replace(/([^/]+)$/, '_rels/$1.rels'), relationships(entries))
  zip.file('[Content_Types].xml', '<Types/>')
  zip.file(
    'ppt/presentation.xml',
    `<p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="${rel(1)}"/><p:sldMasterId id="2147483649" r:id="${rel(2)}"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="512" r:id="${rel(3)}"/></p:sldIdLst></p:presentation>`,
  )
  putRels('ppt/presentation.xml', [
    { id: rel(1), type: 'slideMaster', target: `slideMasters/slideMaster${num(1)}.xml` },
    { id: rel(2), type: 'slideMaster', target: `slideMasters/slideMaster${num(2)}.xml` },
    { id: rel(3), type: 'slide', target: 'slides/slide1.xml' },
  ])
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>')
  putRels('ppt/slides/slide1.xml', [
    {
      id: rel(4),
      type: 'slideLayout',
      target: `../slideLayouts/slideLayout${num(selected + 1)}.xml`,
    },
  ])
  for (let m = 1; m <= 2; m++) {
    const slots = m === 1 ? [1, 2, 3] : [4]
    zip.file(
      master(m),
      `<p:sldMaster xmlns:p="urn:p" xmlns:r="urn:r"><p:cSld name="master${m}"/><p:sldLayoutIdLst>${slots.map((n) => `<p:sldLayoutId id="${800 + n + (renumber ? 100 : 0)}" r:id="${rel(n + 5)}"/>`).join('')}</p:sldLayoutIdLst></p:sldMaster>`,
    )
    putRels(master(m), [
      ...slots.map((n) => ({
        id: rel(n + 5),
        type: 'slideLayout',
        target: `../slideLayouts/slideLayout${num(n)}.xml`,
      })),
      { id: rel(10), type: 'theme', target: `../theme/theme${num(m)}.xml` },
    ])
    zip.file(`ppt/theme/theme${num(m)}.xml`, `<a:theme xmlns:a="urn:a" name="theme${m}"/>`)
    for (const n of slots) {
      zip.file(
        layout(n),
        `<p:sldLayout xmlns:p="urn:p" xmlns:r="urn:r"><p:cSld name="duplicate"><p:spTree value="${n}"/></p:cSld></p:sldLayout>`,
      )
      putRels(layout(n), [
        { id: rel(11), type: 'slideMaster', target: `../slideMasters/slideMaster${num(m)}.xml` },
      ])
    }
  }
  return { zip, base64: await zip.generateAsync({ type: 'base64' }), master, layout }
}
async function inventory(renumber = false) {
  const reps = await Promise.all([0, 1, 2].map((i) => fixture(i, renumber)))
  return {
    masters: [
      {
        masterId: 'actual-master',
        layouts: [2, 0, 1].map((i) => ({
          layoutId: `actual-layout-${i}`,
          representativeBase64: reps[i]!.base64,
        })),
      },
    ],
  }
}
it('reads complete ordered master/layout identities and actual package source ID', async () => {
  const f = await fixture(),
    v = await inspectMasterXmlPackage(f.base64)
  expect(v.sourceSlideId).toBe('512#')
  expect(v.sourceLayoutPath).toBe(f.layout(1))
  expect(v.sourceMasterPath).toBe(f.master(1))
  expect(v.masters).toHaveLength(2)
  expect(v.masters[0]!.orderedLayouts.map((l) => l.packageLayoutId)).toEqual(['801', '802', '803'])
  expect(v.masters[0]!.orderedLayouts.every((l) => /^[a-f0-9]{64}$/.test(l.contentDigest))).toBe(
    true,
  )
})
it('proves complete native mappings despite duplicate names, reordered SDK inventory and renamed package paths/IDs', async () => {
  const f = await fixture(),
    mapping = await proveMasterXmlLayoutMapping(f.base64, await inventory(true))
  expect(mapping).toEqual({
    masterId: 'actual-master',
    sourceLayoutId: 'actual-layout-0',
    layouts: [0, 1, 2].map((i) => ({
      packageLayoutPath: f.layout(i + 1),
      nativeLayoutId: `actual-layout-${i}`,
    })),
  })
})
it('proves recovered original layouts against new native IDs after the original master disappeared', async () => {
  const f = await fixture(),
    native = await inventory()
  native.masters[0]!.masterId = 'restored-master'
  expect((await proveMasterXmlLayoutMapping(f.base64, native)).masterId).toBe('restored-master')
})
it('rejects missing/ambiguous representatives and foreign-master ownership', async () => {
  const f = await fixture()
  for (const mode of ['missing', 'duplicate', 'foreign', 'ambiguous']) {
    const native = await inventory()
    if (mode === 'missing') native.masters[0]!.layouts.pop()
    if (mode === 'duplicate')
      native.masters[0]!.layouts[0]!.representativeBase64 =
        native.masters[0]!.layouts[1]!.representativeBase64
    if (mode === 'foreign')
      native.masters[0]!.layouts[0]!.representativeBase64 = (await fixture(3)).base64
    if (mode === 'ambiguous')
      native.masters.push({
        ...structuredClone(native.masters[0]!),
        masterId: 'other',
        layouts: native.masters[0]!.layouts.map((l) => ({ ...l, layoutId: l.layoutId + 'other' })),
      })
    await expect(proveMasterXmlLayoutMapping(f.base64, native)).rejects.toThrow(
      'presentation_master_xml_package_unproven',
    )
  }
})
it('requires exact immutable dependencies and ordered identities around approved master XML edits', async () => {
  const f = await fixture()
  f.zip.file(
    f.master(1),
    (await f.zip.file(f.master(1))!.async('string')).replace('name="master1"', 'name="edited"'),
  )
  const edited = await f.zip.generateAsync({ type: 'base64' })
  expect(
    (await assertMasterXmlPreparation(f.base64, edited, [f.master(1)])).prepared.sourceSlideId,
  ).toBe('512#')
  f.zip.file('ppt/theme/theme1.xml', '<changed/>')
  await expect(
    assertMasterXmlPreparation(f.base64, await f.zip.generateAsync({ type: 'base64' }), [
      f.master(1),
    ]),
  ).rejects.toThrow('presentation_master_xml_package_unproven')
  const redirected = await fixture()
  redirected.zip.file(
    redirected.master(1),
    (await redirected.zip.file(redirected.master(1))!.async('string')).replace(
      'id="801"',
      'id="804"',
    ),
  )
  await expect(
    assertMasterXmlPreparation(
      redirected.base64,
      await redirected.zip.generateAsync({ type: 'base64' }),
      [redirected.master(1)],
    ),
  ).rejects.toThrow('presentation_master_xml_package_unproven')
})
it.each(['missing', 'duplicate', 'external', 'traversal', 'redirect'])(
  'rejects %s relationship inventories',
  async (mode) => {
    const f = await fixture(),
      path = 'ppt/slideLayouts/_rels/slideLayout1.xml.rels'
    let xml = await f.zip.file(path)!.async('string')
    if (mode === 'missing') f.zip.remove(f.layout(2))
    if (mode === 'duplicate')
      xml = xml.replace(
        '</Relationships>',
        xml.match(/<Relationship .*?\/>/)![0] + '</Relationships>',
      )
    if (mode === 'external') xml = xml.replace('/>', ' TargetMode="External"/>')
    if (mode === 'traversal')
      xml = xml.replace('../slideMasters/slideMaster1.xml', '../../../../foreign.xml')
    if (mode === 'redirect')
      xml = xml.replace('../slideMasters/slideMaster1.xml', '../slideMasters/slideMaster2.xml')
    f.zip.file(path, xml)
    await expect(
      inspectMasterXmlPackage(await f.zip.generateAsync({ type: 'base64' })),
    ).rejects.toThrow('presentation_master_xml_package_unproven')
  },
)
it('freezes supplied representative inventory before awaits and honors cancellation', async () => {
  const f = await fixture(),
    native = await inventory(),
    pending = proveMasterXmlLayoutMapping(f.base64, native)
  native.masters[0]!.layouts[0]!.representativeBase64 = 'aliased'
  expect((await pending).masterId).toBe('actual-master')
  const controller = new AbortController()
  controller.abort()
  await expect(inspectMasterXmlPackage(f.base64, controller.signal)).rejects.toThrow('cancelled')
})

it('accepts XML attribute/relationship order normalization without accepting resource changes', async () => {
  const f = await fixture(),
    native = await inventory()
  for (const layout of native.masters[0]!.layouts) {
    const zip = await JSZip.loadAsync(layout.representativeBase64, { base64: true })
    for (const path of Object.keys(zip.files)) {
      if (path.endsWith('.rels')) {
        const xml = await zip.file(path)!.async('string')
        zip.file(
          path,
          xml.replace(
            /<Relationship Id="([^"]+)" Type="([^"]+)" Target="([^"]+)"\/>/g,
            '<Relationship Target="$3" Type="$2" Id="$1"/>',
          ),
        )
      }
    }
    layout.representativeBase64 = await zip.generateAsync({ type: 'base64' })
  }
  expect((await proveMasterXmlLayoutMapping(f.base64, native)).masterId).toBe('actual-master')
})
it('proves every linked theme and binary resource using SHA content, including FNV collisions', async () => {
  const old = Uint8Array.from([0xb9, 0x31, 0x56, 0xc4, 0x6b, 0xd9, 0x4d, 0xe1]),
    collision = Uint8Array.from([0x5b, 0x9e, 0x9c, 0xd7, 0xfd, 0x87, 0x20, 0x42])
  const enrich = async (base64: string, bytes: Uint8Array) => {
    const zip = await JSZip.loadAsync(base64, { base64: true })
    zip.file('ppt/media/image1.bin', bytes)
    zip.file(
      'ppt/theme/_rels/theme1.xml.rels',
      `<Relationships><Relationship Id="img" Type="${ns}image" Target="../media/image1.bin"/></Relationships>`,
    )
    return zip.generateAsync({ type: 'base64' })
  }
  const expected = await enrich((await fixture()).base64, old),
    native = await inventory()
  for (const l of native.masters[0]!.layouts)
    l.representativeBase64 = await enrich(l.representativeBase64, old)
  expect((await proveMasterXmlLayoutMapping(expected, native)).masterId).toBe('actual-master')
  native.masters[0]!.layouts[0]!.representativeBase64 = await enrich(
    native.masters[0]!.layouts[0]!.representativeBase64,
    collision,
  )
  await expect(proveMasterXmlLayoutMapping(expected, native)).rejects.toThrow(
    'presentation_master_xml_package_unproven',
  )
  const themes = await inventory()
  const zip = await JSZip.loadAsync(themes.masters[0]!.layouts[0]!.representativeBase64, {
    base64: true,
  })
  zip.file('ppt/theme/theme1.xml', '<a:theme xmlns:a="urn:a" name="changed"/>')
  themes.masters[0]!.layouts[0]!.representativeBase64 = await zip.generateAsync({ type: 'base64' })
  await expect(proveMasterXmlLayoutMapping((await fixture()).base64, themes)).rejects.toThrow(
    'presentation_master_xml_package_unproven',
  )
})
it('requires numeric unique ordered layout identities and bounded XML programs', async () => {
  const f = await fixture()
  f.zip.file(
    f.master(1),
    (await f.zip.file(f.master(1))!.async('string')).replace('id="802"', 'id="801"'),
  )
  await expect(
    inspectMasterXmlPackage(await f.zip.generateAsync({ type: 'base64' })),
  ).rejects.toThrow('presentation_master_xml_package_unproven')
  await expect(
    assertMasterXmlPreparation(
      f.base64,
      f.base64,
      Array.from({ length: 33 }, (_, i) => `ppt/slideMasters/slideMaster${i + 1}.xml`),
    ),
  ).rejects.toThrow('presentation_master_xml_package_unproven')
  await expect(
    assertMasterXmlPreparation(f.base64, f.base64, ['ppt/slides/slide1.xml']),
  ).rejects.toThrow('presentation_master_xml_package_unproven')
})

it('maps a requested secondary master rather than silently mapping the carrier master', async () => {
  const f = await fixture(),
    representative = await fixture(3, true)
  const mapping = await proveMasterXmlLayoutMapping(
    f.base64,
    {
      masters: [
        {
          masterId: 'secondary-native',
          layouts: [{ layoutId: 'secondary-layout', representativeBase64: representative.base64 }],
        },
      ],
    },
    undefined,
    f.master(2),
  )
  expect(mapping).toEqual({
    masterId: 'secondary-native',
    sourceLayoutId: 'secondary-layout',
    layouts: [{ packageLayoutPath: f.layout(4), nativeLayoutId: 'secondary-layout' }],
  })
})
it('derives all affected master graphs for approved layout/theme edits including shared themes', async () => {
  const f = await fixture()
  const relPath = 'ppt/slideMasters/_rels/slideMaster2.xml.rels'
  f.zip.file(
    relPath,
    (await f.zip.file(relPath)!.async('string')).replace('theme2.xml', 'theme1.xml'),
  )
  const original = await f.zip.generateAsync({ type: 'base64' })
  f.zip.file('ppt/theme/theme1.xml', '<a:theme xmlns:a="urn:a" name="shared changed"/>')
  const changed = await assertMasterXmlPreparation(
    original,
    await f.zip.generateAsync({ type: 'base64' }),
    ['ppt/theme/theme1.xml'],
  )
  expect(changed.affectedMasterPaths).toEqual([f.master(1), f.master(2)])
  const l = await fixture()
  l.zip.file(l.layout(4), '<p:sldLayout xmlns:p="urn:p"><p:cSld name="changed"/></p:sldLayout>')
  expect(
    (
      await assertMasterXmlPreparation(l.base64, await l.zip.generateAsync({ type: 'base64' }), [
        l.layout(4),
      ])
    ).affectedMasterPaths,
  ).toEqual([l.master(2)])
})

async function payload(base64: string) {
  const zip = await JSZip.loadAsync(base64, { base64: true })
  zip.file(
    'ppt/slides/slide1.xml',
    '<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r"><p:cSld><p:spTree><a:blip r:embed="image"/><a:t>protected payload</a:t></p:spTree></p:cSld></p:sld>',
  )
  const rels = 'ppt/slides/_rels/slide1.xml.rels'
  zip.file(
    rels,
    (await zip.file(rels)!.async('string')).replace(
      '</Relationships>',
      `<Relationship Id="image" Type="${ns}image" Target="../media/image1.bin"/><Relationship Id="notes" Type="${ns}notesSlide" Target="../notesSlides/notesSlide1.xml"/><Relationship Id="chart" Type="${ns}chart" Target="../charts/chart1.xml"/><Relationship Id="link" Type="${ns}hyperlink" Target="https://example.com/source?a=1&amp;b=2" TargetMode="External"/></Relationships>`,
    ),
  )
  zip.file(
    'ppt/media/image1.bin',
    Uint8Array.from([0xb9, 0x31, 0x56, 0xc4, 0x6b, 0xd9, 0x4d, 0xe1]),
  )
  zip.file(
    'ppt/notesSlides/notesSlide1.xml',
    '<p:notes xmlns:p="urn:p"><p:cSld name="protected notes"/></p:notes>',
  )
  zip.file(
    'ppt/charts/chart1.xml',
    '<c:chart xmlns:c="urn:c"><c:title>protected chart</c:title></c:chart>',
  )
  zip.file('ppt/media/unreferenced.bin', 'protected unreferenced')
  zip.file(
    '[Content_Types].xml',
    '<Types><Override PartName="/ppt/slides/slide1.xml" ContentType="slide"/><Override PartName="/ppt/media/image1.bin" ContentType="image"/></Types>',
  )
  return { zip, base64: await zip.generateAsync({ type: 'base64' }) }
}
async function pageRecipe() {
  const f = await fixture(),
    original = await payload(f.base64),
    actual = await payload((await fixture(0, true)).base64)
  const expected = await fixture()
  expected.zip.file(
    expected.master(1),
    (await expected.zip.file(expected.master(1))!.async('string')).replace(
      'name="master1"',
      'name="edited master"',
    ),
  )
  actual.zip.file(
    'ppt/slideMasters/slideMaster11.xml',
    (await actual.zip.file('ppt/slideMasters/slideMaster11.xml')!.async('string')).replace(
      'name="master1"',
      'name="edited master"',
    ),
  )
  return {
    original,
    actual,
    options: {
      expectedMasterBase64: await expected.zip.generateAsync({ type: 'base64' }),
      targetMasterPath: f.master(1),
      packageLayoutPath: f.layout(1),
    },
  }
}
it('preserves entire projected page archives with native-renumbered master paths and unchanged external leaves', async () => {
  const f = await pageRecipe()
  await expect(
    assertMasterXmlPagePreserved(
      f.original.base64,
      await f.actual.zip.generateAsync({ type: 'base64' }),
      f.options,
    ),
  ).resolves.toBeUndefined()
})
it.each(['payload', 'notes', 'chart', 'media', 'unreferenced', 'url', 'mode', 'layout'])(
  'refuses %s drift after layout application instead of blessing only its master',
  async (kind) => {
    const f = await pageRecipe()
    if (kind === 'payload')
      f.actual.zip.file(
        'ppt/slides/slide1.xml',
        '<p:sld xmlns:p="urn:p"><p:cSld name="changed placeholder"/></p:sld>',
      )
    if (kind === 'notes') f.actual.zip.file('ppt/notesSlides/notesSlide1.xml', '<changed/>')
    if (kind === 'chart') f.actual.zip.file('ppt/charts/chart1.xml', '<changed/>')
    if (kind === 'media')
      f.actual.zip.file(
        'ppt/media/image1.bin',
        Uint8Array.from([0x5b, 0x9e, 0x9c, 0xd7, 0xfd, 0x87, 0x20, 0x42]),
      )
    if (kind === 'unreferenced')
      f.actual.zip.file('ppt/media/unreferenced.bin', 'modified unreferenced')
    if (['url', 'mode', 'layout'].includes(kind)) {
      const path = 'ppt/slides/_rels/slide1.xml.rels',
        xml = await f.actual.zip.file(path)!.async('string')
      f.actual.zip.file(
        path,
        kind === 'url'
          ? xml.replace('https://example.com/source', 'https://other.example/source')
          : kind === 'mode'
            ? xml.replace('TargetMode="External"', 'TargetMode="Internal"')
            : xml.replace('slideLayout11.xml', 'slideLayout12.xml'),
      )
    }
    await expect(
      assertMasterXmlPagePreserved(
        f.original.base64,
        await f.actual.zip.generateAsync({ type: 'base64' }),
        f.options,
      ),
    ).rejects.toThrow('presentation_master_xml_package_unproven')
  },
)
it('refuses content-type redirection and preserves remote hyperlink metadata without fetching', async () => {
  const f = await pageRecipe()
  const path = '[Content_Types].xml'
  f.actual.zip.file(
    path,
    (await f.actual.zip.file(path)!.async('string')).replace(
      'ContentType="image"',
      'ContentType="other"',
    ),
  )
  await expect(
    assertMasterXmlPagePreserved(
      f.original.base64,
      await f.actual.zip.generateAsync({ type: 'base64' }),
      f.options,
    ),
  ).rejects.toThrow('presentation_master_xml_package_unproven')
  const original = await payload((await fixture()).base64)
  expect((await inspectMasterXmlPackage(original.base64)).sourceSlideId).toBe('512#')
})

it('preserves opaque metadata text bytes rather than treating meaningful whitespace as formatting', async () => {
  const f = await pageRecipe()
  f.original.zip.file(
    'docProps/core.xml',
    '<cp:coreProperties xmlns:cp="urn:cp" xmlns:dc="urn:dc"><dc:title> </dc:title></cp:coreProperties>',
  )
  f.actual.zip.file(
    'docProps/core.xml',
    '<cp:coreProperties xmlns:cp="urn:cp" xmlns:dc="urn:dc"><dc:title></dc:title></cp:coreProperties>',
  )
  await expect(
    assertMasterXmlPagePreserved(
      await f.original.zip.generateAsync({ type: 'base64' }),
      await f.actual.zip.generateAsync({ type: 'base64' }),
      f.options,
    ),
  ).rejects.toThrow('presentation_master_xml_package_unproven')
})
it('protects unreferenced binary entries against the same-length FNV collision', async () => {
  const f = await pageRecipe()
  f.original.zip.file(
    'ppt/media/unreferenced.bin',
    Uint8Array.from([0xb9, 0x31, 0x56, 0xc4, 0x6b, 0xd9, 0x4d, 0xe1]),
  )
  f.actual.zip.file(
    'ppt/media/unreferenced.bin',
    Uint8Array.from([0x5b, 0x9e, 0x9c, 0xd7, 0xfd, 0x87, 0x20, 0x42]),
  )
  await expect(
    assertMasterXmlPagePreserved(
      await f.original.zip.generateAsync({ type: 'base64' }),
      await f.actual.zip.generateAsync({ type: 'base64' }),
      f.options,
    ),
  ).rejects.toThrow('presentation_master_xml_package_unproven')
})
it('copies page-preservation options before the first await and preserves explicit cancellation', async () => {
  const f = await pageRecipe(),
    actual = await f.actual.zip.generateAsync({ type: 'base64' })
  const pending = assertMasterXmlPagePreserved(f.original.base64, actual, f.options)
  f.options.expectedMasterBase64 = 'aliased'
  f.options.targetMasterPath = 'aliased'
  await expect(pending).resolves.toBeUndefined()
  const controller = new AbortController()
  controller.abort()
  await expect(
    assertMasterXmlPagePreserved(f.original.base64, actual, f.options, controller.signal),
  ).rejects.toThrow('cancelled')
})
it('preserves exact paths for opaque parts without relationship identity evidence', async () => {
  const f = await pageRecipe()
  const bytes = await f.actual.zip.file('ppt/media/unreferenced.bin')!.async('uint8array')
  f.actual.zip.remove('ppt/media/unreferenced.bin')
  f.actual.zip.file('ppt/media/foreign-renamed.bin', bytes)
  await expect(
    assertMasterXmlPagePreserved(
      f.original.base64,
      await f.actual.zip.generateAsync({ type: 'base64' }),
      f.options,
    ),
  ).rejects.toThrow('presentation_master_xml_package_unproven')
})
