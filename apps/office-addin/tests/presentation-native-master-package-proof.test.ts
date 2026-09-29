import { afterEach, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  nativeMasterFixture,
  cleanupNativeMasterFixtures,
} from './helpers/native-master-fixture.js'
import {
  prepareMasterPackageProtection,
  verifyMasterPackageProtection,
} from '../src/skills/powerpoint/presentation-native-master-package-proof.js'
import type { StoredMasterOperation } from '../src/skills/powerpoint/presentation-native-master-change.js'
afterEach(cleanupNativeMasterFixtures)

it('normalizes only the authorized theme slot while preserving shared theme fonts and master placeholders', async () => {
  const f = await nativeMasterFixture(),
    operations: StoredMasterOperation[] = [f.op]
  const original = await f.adapter.exportSlidePackage(0)
  const dependency = f.dependencies.slides[0]!
  const protection = await prepareMasterPackageProtection(original.base64, dependency, operations)
  await f.adapter.executeMasterOperations([f.op])
  const changed = await f.adapter.exportSlidePackage(0)
  expect(
    await verifyMasterPackageProtection(
      changed.base64,
      protection,
      dependency,
      operations,
      operations,
    ),
  ).toBe(true)
  f.drift.set('theme_font', 'Foreign font')
  expect(
    await verifyMasterPackageProtection(
      (await f.adapter.exportSlidePackage(0)).base64,
      protection,
      dependency,
      operations,
      operations,
    ),
  ).toBe(false)
})
it.each(['slide', 'notes', 'placeholder', 'unselected_theme_slot', 'relationship', 'media'])(
  'refuses foreign %s changes alongside an authorized theme update',
  async (part) => {
    const f = await nativeMasterFixture(),
      operations: StoredMasterOperation[] = [f.op],
      dependency = f.dependencies.slides[0]!
    const protection = await prepareMasterPackageProtection(
      (await f.adapter.exportSlidePackage(0)).base64,
      dependency,
      operations,
    )
    await f.adapter.executeMasterOperations([f.op])
    const zip = await JSZip.loadAsync((await f.adapter.exportSlidePackage(0)).base64, {
      base64: true,
    })
    if (part === 'slide')
      zip.file(
        'ppt/slides/slide1.xml',
        (await zip.file('ppt/slides/slide1.xml')!.async('string')).replace('0-', 'Foreign'),
      )
    if (part === 'notes')
      zip.file(
        'ppt/notesSlides/notesSlide1.xml',
        (await zip.file('ppt/notesSlides/notesSlide1.xml')!.async('string')).replace(
          'Original note',
          'Foreign note',
        ),
      )
    if (part === 'placeholder')
      zip.file(
        'ppt/slideMasters/slideMaster1.xml',
        (await zip.file('ppt/slideMasters/slideMaster1.xml')!.async('string')).replace(
          'Original placeholder',
          'Foreign placeholder',
        ),
      )
    if (part === 'unselected_theme_slot')
      zip.file(
        'ppt/theme/theme1.xml',
        (await zip.file('ppt/theme/theme1.xml')!.async('string')).replace(
          '<a:accent2><a:srgbClr val="FFFFFF"/>',
          '<a:accent2><a:srgbClr val="123456"/>',
        ),
      )
    if (part === 'relationship')
      zip.file(
        'ppt/slideMasters/_rels/slideMaster1.xml.rels',
        (await zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels')!.async('string')).replace(
          'Id="theme"',
          'Id="Foreign"',
        ),
      )
    if (part === 'media') zip.file('ppt/media/foreign.png', new Uint8Array([1, 2, 3]))
    const base64 = await zip.generateAsync({ type: 'base64' })
    if (part === 'media')
      await expect(
        verifyMasterPackageProtection(base64, protection, dependency, operations, operations),
      ).rejects.toThrow('presentation_native_master_package_unproven')
    else
      expect(
        await verifyMasterPackageProtection(base64, protection, dependency, operations, operations),
      ).toBe(false)
  },
)
it('preserves formatting run order and XML comments while ignoring harmless indentation', async () => {
  const f = await nativeMasterFixture(),
    dependency = f.dependencies.slides[0]!,
    operations: StoredMasterOperation[] = [f.op]
  const zip = await JSZip.loadAsync((await f.adapter.exportSlidePackage(0)).base64, {
    base64: true,
  })
  const source =
    '<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><!--keep--><a:p><a:r><a:t>First</a:t></a:r><a:br/><a:r><a:t>Second</a:t></a:r></a:p></p:cSld></p:sld>'
  zip.file('ppt/slides/slide1.xml', source)
  const protection = await prepareMasterPackageProtection(
    await zip.generateAsync({ type: 'base64' }),
    dependency,
    operations,
  )
  zip.file('ppt/slides/slide1.xml', source.replace('<p:cSld>', '<p:cSld>\n  '))
  expect(
    await verifyMasterPackageProtection(
      await zip.generateAsync({ type: 'base64' }),
      protection,
      dependency,
      operations,
      [],
    ),
  ).toBe(true)
  zip.file('ppt/slides/slide1.xml', source.replace('<!--keep-->', '<!--changed-->'))
  expect(
    await verifyMasterPackageProtection(
      await zip.generateAsync({ type: 'base64' }),
      protection,
      dependency,
      operations,
      [],
    ),
  ).toBe(false)
  zip.file(
    'ppt/slides/slide1.xml',
    source.replace('<a:r><a:t>First</a:t></a:r><a:br/>', '<a:br/><a:r><a:t>First</a:t></a:r>'),
  )
  expect(
    await verifyMasterPackageProtection(
      await zip.generateAsync({ type: 'base64' }),
      protection,
      dependency,
      operations,
      [],
    ),
  ).toBe(false)
})

it('rejects persisted protection missing required original target keys', async () => {
  const f = await nativeMasterFixture()
  const dependency = f.dependencies.slides[0]!
  const operations: StoredMasterOperation[] = [f.op]
  const base64 = (await f.adapter.exportSlidePackage(0)).base64
  const protection = await prepareMasterPackageProtection(base64, dependency, operations)
  protection.targetDigests = {}
  await expect(
    verifyMasterPackageProtection(base64, protection, dependency, operations, []),
  ).rejects.toThrow('presentation_master_backup_invalid')
})
