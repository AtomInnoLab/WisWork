import JSZip from 'jszip'
import { vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { createPresentationPackageBackupService } from '../../../shell/src/main/presentation-package-backups.js'
import { createStructuredProposalController } from '../../src/agent/proposal-controller.js'
import { createPresentationMasterXmlSkill } from '../../src/skills/powerpoint/presentation-master-xml.js'
import {
  validMasterXmlTransition,
  type PresentationMasterXmlChange,
} from '../../src/skills/powerpoint/presentation-master-xml-change.js'
import { presentationPackageDigest } from '../../src/skills/powerpoint/powerpoint-package.js'
import { inspectMasterXmlPackage } from '../../src/skills/powerpoint/presentation-master-xml-package.js'
import type { MasterXmlHostSnapshot } from '../../src/skills/powerpoint/browser-presentation-master-xml-adapter.js'
const roots: string[] = []
export function cleanupMasterXmlFixtures() {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
}
const ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
export async function masterXmlFixture(
  count = 3,
  settings: { multipleMasters?: boolean; disappearOriginalOnDelete?: boolean } = {},
) {
  const zip = new JSZip()
  const rel = (entries: { id: string; type: string; target: string }[]) =>
    `<Relationships>${entries.map((e) => `<Relationship Id="${e.id}" Type="${ns + e.type}" Target="${e.target}"/>`).join('')}</Relationships>`
  zip.file(
    '[Content_Types].xml',
    '<Types><Default Extension="xml" ContentType="application/xml"/></Types>',
  )
  zip.file(
    'ppt/presentation.xml',
    '<p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="m1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="s1"/></p:sldIdLst></p:presentation>',
  )
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    rel([
      { id: 'm1', type: 'slideMaster', target: 'slideMasters/slideMaster1.xml' },
      { id: 's1', type: 'slide', target: 'slides/slide1.xml' },
    ]),
  )
  zip.file(
    'ppt/slides/slide1.xml',
    '<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><a:t>original</a:t></p:spTree></p:cSld></p:sld>',
  )
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    rel([{ id: 'l1', type: 'slideLayout', target: '../slideLayouts/slideLayout1.xml' }]),
  )
  const originalMaster =
    '<p:sldMaster xmlns:p="urn:p" xmlns:r="urn:r"><p:cSld name="original"/><p:sldLayoutIdLst><p:sldLayoutId id="801" r:id="l1"/><p:sldLayoutId id="802" r:id="l2"/></p:sldLayoutIdLst></p:sldMaster>'
  zip.file('ppt/slideMasters/slideMaster1.xml', originalMaster)
  zip.file(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    rel([
      { id: 'l1', type: 'slideLayout', target: '../slideLayouts/slideLayout1.xml' },
      { id: 'l2', type: 'slideLayout', target: '../slideLayouts/slideLayout2.xml' },
      { id: 't1', type: 'theme', target: '../theme/theme1.xml' },
    ]),
  )
  zip.file('ppt/theme/theme1.xml', '<a:theme xmlns:a="urn:a" name="original-theme"/>')
  for (let i = 1; i <= 2; i++) {
    zip.file(
      `ppt/slideLayouts/slideLayout${i}.xml`,
      `<p:sldLayout xmlns:p="urn:p"><p:cSld name="duplicate"><p:spTree value="${i}"/></p:cSld></p:sldLayout>`,
    )
    zip.file(
      `ppt/slideLayouts/_rels/slideLayout${i}.xml.rels`,
      rel([{ id: 'm1', type: 'slideMaster', target: '../slideMasters/slideMaster1.xml' }]),
    )
  }
  if (settings.multipleMasters) {
    zip.file(
      'ppt/presentation.xml',
      '<p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="m1"/><p:sldMasterId id="2147483649" r:id="m2"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="s1"/></p:sldIdLst></p:presentation>',
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      rel([
        { id: 'm1', type: 'slideMaster', target: 'slideMasters/slideMaster1.xml' },
        { id: 'm2', type: 'slideMaster', target: 'slideMasters/slideMaster2.xml' },
        { id: 's1', type: 'slide', target: 'slides/slide1.xml' },
      ]),
    )
    zip.file(
      'ppt/slideMasters/slideMaster2.xml',
      '<p:sldMaster xmlns:p="urn:p" xmlns:r="urn:r"><p:cSld name="secondary"/><p:sldLayoutIdLst><p:sldLayoutId id="803" r:id="l3"/></p:sldLayoutIdLst></p:sldMaster>',
    )
    zip.file(
      'ppt/slideMasters/_rels/slideMaster2.xml.rels',
      rel([
        { id: 'l3', type: 'slideLayout', target: '../slideLayouts/slideLayout3.xml' },
        { id: 't2', type: 'theme', target: '../theme/theme2.xml' },
      ]),
    )
    zip.file(
      'ppt/slideLayouts/slideLayout3.xml',
      '<p:sldLayout xmlns:p="urn:p"><p:cSld name="duplicate"><p:spTree value="3"/></p:cSld></p:sldLayout>',
    )
    zip.file(
      'ppt/slideLayouts/_rels/slideLayout3.xml.rels',
      rel([{ id: 'm2', type: 'slideMaster', target: '../slideMasters/slideMaster2.xml' }]),
    )
    zip.file('ppt/theme/theme2.xml', '<a:theme xmlns:a="urn:a" name="secondary-theme"/>')
  }
  const original = await zip.generateAsync({ type: 'base64' }),
    root = mkdtempSync(join(tmpdir(), 'master-xml-'))
  roots.push(root)
  const service = createPresentationPackageBackupService({ userDataPath: root }),
    data = new Map<string, PresentationMasterXmlChange>(),
    order = Array.from({ length: count }, (_, i) => `s${i}`),
    packages = new Map(order.map((id) => [id, original]))
  type Master = {
    masterId: string
    name: string
    layouts: { layoutId: string; name: string; path: string }[]
    base64: string
  }
  const masters: Master[] = [
    {
      masterId: 'm1',
      name: 'same',
      layouts: [1, 2].map((i) => ({
        layoutId: `l${i}`,
        name: 'duplicate',
        path: `ppt/slideLayouts/slideLayout${i}.xml`,
      })),
      base64: original,
    },
  ]
  if (settings.multipleMasters) {
    masters.push({
      masterId: 'm2',
      name: 'same',
      layouts: [{ layoutId: 'l3', name: 'duplicate', path: 'ppt/slideLayouts/slideLayout3.xml' }],
      base64: original,
    })
    for (let i = 1; i < order.length; i += 2) {
      const next = await JSZip.loadAsync(original, { base64: true })
      next.file(
        'ppt/slides/_rels/slide1.xml.rels',
        rel([{ id: 'l3', type: 'slideLayout', target: '../slideLayouts/slideLayout3.xml' }]),
      )
      packages.set(order[i]!, await next.generateAsync({ type: 'base64' }))
    }
  }
  const deps = new Map(
    order.map((id) => [
      id,
      {
        slideId: id,
        masterId: settings.multipleMasters && order.indexOf(id) % 2 ? 'm2' : 'm1',
        layoutId: settings.multipleMasters && order.indexOf(id) % 2 ? 'l3' : 'l1',
      },
    ]),
  )
  let index = 0,
    mode = '',
    doc = 'doc',
    available = true
  let readHook: () => void = () => {}
  const request = vi.fn(
    async (body: unknown, signal?: AbortSignal) =>
      new Response(
        JSON.stringify(
          await service(body as Record<string, unknown>, signal ?? new AbortController().signal),
        ),
      ),
  )
  const digestMemo = new Map<string, Promise<string>>()
  const digest = (base64: string) => {
    let value = digestMemo.get(base64)
    if (!value) {
      value = presentationPackageDigest(base64)
      digestMemo.set(base64, value)
    }
    return value
  }
  const snapshot = async (exports: string[] = []): Promise<MasterXmlHostSnapshot> => ({
    slideIds: [...order],
    pages: await Promise.all(
      order.map(async (slideId) => ({
        slideId,
        digest: await digest(packages.get(slideId)!),
        ...(exports.includes(slideId) ? { base64: packages.get(slideId)! } : {}),
      })),
    ),
    masters: masters.map((m) => ({
      masterId: m.masterId,
      name: m.name,
      layouts: m.layouts.map(({ layoutId, name }) => ({ layoutId, name })),
    })),
    dependencies: order.map((id) => structuredClone(deps.get(id)!)),
  })
  const assertPreimage = async (
    preimage: MasterXmlHostSnapshot,
    beforeWrite: () => Promise<void>,
    writeGuard: () => void,
  ) => {
    await beforeWrite()
    if (JSON.stringify(await snapshot()) !== JSON.stringify(preimage))
      throw Error('office_replacement_preimage_drift')
    writeGuard()
  }
  const adapter = {
    screenshotSlide: vi.fn(async (index: number) => ({
      slideId: order[index]!,
      mime: 'image/png' as const,
      base64:
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII=',
    })),
    inspect: vi.fn(snapshot),
    readPage: vi.fn(async (slideId: string) => ({
      ...deps.get(slideId)!,
      digest: await digest(packages.get(slideId)!),
      base64: packages.get(slideId)!,
    })),
    stage: vi.fn(
      async (
        input: any,
        callback: (v: any) => Promise<void>,
        beforeWrite: () => Promise<void>,
        writeGuard: () => void,
      ) => {
        await assertPreimage(input.preimage, beforeWrite, writeGuard)
        if (mode === 'stage_before') throw Error('lost_ack')
        const inventory = await inspectMasterXmlPackage(input.base64),
          inserted = ++index,
          masterId = `m-import-${inserted}`,
          slideId = `import-${inserted}`
        let actual: { slideId: string; masterId: string; layoutId: string } | undefined
        for (const [masterIndex, graph] of inventory.masters
          .filter((m) => m.path === inventory.sourceMasterPath)
          .entries()) {
          const nativeId = masterIndex === 0 ? masterId : `${masterId}-${masterIndex}`,
            layouts = graph.orderedLayouts.map((l, i) => ({
              layoutId: `${nativeId}-l${i}`,
              name: 'duplicate',
              path: l.path,
            }))
          masters.push({ masterId: nativeId, name: 'same', layouts, base64: input.base64 })
          if (graph.path === inventory.sourceMasterPath)
            actual = {
              slideId,
              masterId: nativeId,
              layoutId: layouts.find((l) => l.path === inventory.sourceLayoutPath)!.layoutId,
            }
        }
        if (!actual) throw Error('missing_primary')
        order.splice(order.indexOf(input.sourceSlideId) + 1, 0, slideId)
        packages.set(slideId, input.base64)
        deps.set(slideId, structuredClone(actual))
        if (mode === 'stage_after') throw Error('lost_ack')
        await callback(actual)
        if (mode === 'callback_after') throw Error('lost_ack')
        return actual
      },
    ),
    applyLayout: vi.fn(
      async (input: any, beforeWrite: () => Promise<void>, writeGuard: () => void) => {
        await assertPreimage(input.preimage, beforeWrite, writeGuard)
        if (mode === 'layout_before') throw Error('lost_ack')
        const master = masters.find((m) => m.masterId === input.masterId)!,
          layout = master.layouts.find((l) => l.layoutId === input.layoutId)!
        const target = await JSZip.loadAsync(master.base64, { base64: true }),
          old = await JSZip.loadAsync(packages.get(input.slideId)!, { base64: true })
        target.file(
          'ppt/slides/slide1.xml',
          await old.file('ppt/slides/slide1.xml')!.async('string'),
        )
        target.file(
          'ppt/slides/_rels/slide1.xml.rels',
          rel([{ id: 'l1', type: 'slideLayout', target: '../' + layout.path.slice(4) }]),
        )
        if (mode === 'payload_loss')
          target.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>')
        packages.set(input.slideId, await target.generateAsync({ type: 'base64' }))
        deps.set(input.slideId, {
          slideId: input.slideId,
          masterId: input.masterId,
          layoutId: input.layoutId,
        })
        if (mode === 'layout_after') throw Error('lost_ack')
      },
    ),
    remove: vi.fn(async (input: any, beforeWrite: () => Promise<void>, writeGuard: () => void) => {
      await assertPreimage(input.preimage, beforeWrite, writeGuard)
      if (mode === 'remove_before') throw Error('lost_ack')
      order.splice(order.indexOf(input.slideId), 1)
      packages.delete(input.slideId)
      deps.delete(input.slideId)
      if (settings.disappearOriginalOnDelete)
        for (let i = masters.length - 1; i >= 0; i--)
          if (
            ['m1', 'm2'].includes(masters[i]!.masterId) &&
            !Array.from(deps.values()).some((d) => d.masterId === masters[i]!.masterId)
          )
            masters.splice(i, 1)
      if (mode === 'remove_after') throw Error('lost_ack')
    }),
  }
  const proposals = createStructuredProposalController(),
    options = {
      documentId: async () => {
        readHook()
        return doc
      },
      assertDocumentId: (expected: string) => {
        if (doc !== expected) throw Error('presentation_document_changed')
      },
      available: () => available,
      adapter,
      request,
      proposals,
      readMasterXmlChange: (id: string) => data.get(id),
      writeMasterXmlChange: async (
        next: PresentationMasterXmlChange,
        expected: PresentationMasterXmlChange | undefined,
      ) => {
        if (
          JSON.stringify(data.get(next.changeId)) !== JSON.stringify(expected) ||
          !validMasterXmlTransition(expected, next)
        )
          throw Error('settings_conflict')
        data.set(next.changeId, structuredClone(next))
      },
    }
  let skill = createPresentationMasterXmlSkill(options)
  return {
    data,
    deps,
    options,
    useBoundedTelemetry() {
      // Stress runs retain only IDs; Vitest otherwise retains every full snapshot and Response.
      const layoutSlideIds: string[] = [],
        apply = adapter.applyLayout.getMockImplementation()!
      options.request = request.getMockImplementation()! as typeof request
      options.adapter = {
        screenshotSlide: adapter.screenshotSlide.getMockImplementation()!,
        inspect: adapter.inspect.getMockImplementation()!,
        readPage: adapter.readPage.getMockImplementation()!,
        stage: adapter.stage.getMockImplementation()!,
        applyLayout: (...args: Parameters<typeof apply>) => {
          layoutSlideIds.push(args[0].slideId)
          return apply(...args)
        },
        remove: adapter.remove.getMockImplementation()!,
      } as typeof adapter
      return { layoutSlideIds }
    },
    readHook(fn?: () => void) {
      readHook = fn ?? (() => {})
    },
    adapter,
    proposals,
    request,
    order,
    packages,
    masters,
    originalMaster,
    original,
    propose: (
      replacements = [
        {
          path: 'ppt/slideMasters/slideMaster1.xml',
          xml: originalMaster.replace('name="original"', 'name="edited"'),
        },
      ],
      signal?: AbortSignal,
    ) => skill.propose(replacements, undefined, signal),
    confirm: () => proposals.confirm(proposals.pending()!.id),
    tool: (action: string, input: any, signal?: AbortSignal) =>
      skill.executeTool({ name: `${action}_master_xml_change`, input }, signal),
    clear() {
      skill.clear()
    },
    reopen() {
      skill.clear()
      skill = createPresentationMasterXmlSkill(options)
    },
    mode(value: string) {
      mode = value
    },
    doc(value: string) {
      doc = value
    },
    available(value: boolean) {
      available = value
    },
  }
}
