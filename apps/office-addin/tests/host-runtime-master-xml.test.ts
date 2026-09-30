import { afterEach, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { createOfficeHostRuntime } from '../src/agent/host-runtime'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { inspectMasterXmlPackage } from '../src/skills/powerpoint/presentation-master-xml-package'
import { masterXmlFixture, cleanupMasterXmlFixtures } from './helpers/master-xml-fixture'
const lockReads = vi.hoisted(() => vi.fn())
vi.mock('../src/skills/powerpoint/presentation-native-locks', async (importOriginal) => {
  const actual =
    await importOriginal<typeof import('../src/skills/powerpoint/presentation-native-locks')>()
  return {
    ...actual,
    readPresentationNativeLocks: (
      ...args: Parameters<typeof actual.readPresentationNativeLocks>
    ) => {
      if (args[1].impact.host === 'powerpoint') lockReads()
      return actual.readPresentationNativeLocks(...args)
    },
  }
})
afterEach(() => {
  vi.unstubAllGlobals()
  cleanupMasterXmlFixtures()
})
async function fixture(platform = 'PC', bindingAvailable = true) {
  const f = await masterXmlFixture(2),
    values = new Map<string, string>()
  const binding = createPresentationDocumentBinding(
    {
      get: (k) => values.get(k),
      set: (k, v) => {
        values.set(k, v)
      },
      save: async () => {},
      location: () => 'deck',
    },
    () => 'doc',
  )
  await binding.documentId()
  const invalidateQa = vi.fn(binding.invalidateQa)
  let capability = true,
    paired = true,
    api = true,
    failPending = false,
    index = 0
  type Layout = { id: string; name: string; masterId: string; path: string }
  const masters: { id: string; name: string; bytes: string; layouts: any }[] = []
  const collection = (items: any[]) => ({
    items,
    load: vi.fn(),
    getCount: () => ({ value: items.length }),
    getItem: (id: string) => items.find((x) => x.id === id)!,
  })
  function master(id: string, bytes: string, paths: string[]) {
    const result = {
      id,
      name: 'duplicate',
      bytes,
      layouts: collection(
        paths.map((path, i) => ({ id: `${id}-l${i}`, name: 'duplicate', masterId: id, path })),
      ),
    }
    masters.push(result)
    return result
  }
  const original = master('m1', f.original, [
    'ppt/slideLayouts/slideLayout1.xml',
    'ppt/slideLayouts/slideLayout2.xml',
  ])
  let queued: (() => Promise<void>) | undefined
  const insert = vi.fn((bytes: string, options: { targetSlideId: string }) => {
    queued = async () => {
      const inventory = await inspectMasterXmlPackage(bytes),
        graph = inventory.masters.find((m) => m.path === inventory.sourceMasterPath)!
      const imported = master(
        `import-master-${++index}`,
        bytes,
        graph.orderedLayouts.map((l) => l.path),
      )
      slides.items.splice(
        slides.items.findIndex((p) => p.id === options.targetSlideId) + 1,
        0,
        page(
          `import-page-${index}`,
          bytes,
          imported.id,
          imported.layouts.items.find((l: Layout) => l.path === inventory.sourceLayoutPath).id,
        ),
      )
    }
  })
  const remove = vi.fn((id: string) => {
    queued = async () => {
      slides.items = slides.items.filter((p) => p.id !== id)
    }
  })
  const apply = vi.fn((id: string, layout: Layout) => {
    queued = async () => {
      const target = masters.find((m) => m.id === layout.masterId)!,
        p = slides.getItem(id)
      const zip = await JSZip.loadAsync(target.bytes, { base64: true }),
        old = await JSZip.loadAsync(p.bytes, { base64: true })
      zip.file('ppt/slides/slide1.xml', await old.file('ppt/slides/slide1.xml')!.async('string'))
      zip.file(
        'ppt/slides/_rels/slide1.xml.rels',
        `<Relationships><Relationship Id="l1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../${layout.path.slice(4)}"/></Relationships>`,
      )
      p.bytes = await zip.generateAsync({ type: 'base64' })
      p.slideMaster.id = layout.masterId
      p.layout.id = layout.id
    }
  })
  function page(id: string, bytes: string, masterId = 'm1', layoutId = 'm1-l0') {
    return {
      id,
      bytes,
      slideMaster: { id: masterId },
      layout: { id: layoutId },
      load: vi.fn(),
      exportAsBase64() {
        return { value: this.bytes }
      },
      delete() {
        remove(id)
      },
      applyLayout(layout: Layout) {
        apply(id, layout)
      },
    }
  }
  const slides = {
    items: [page('s0', f.original), page('s1', f.original)],
    load: vi.fn(),
    getCount: () => ({ value: slides.items.length }),
    getItem: (id: string) => slides.items.find((p) => p.id === id)!,
  }
  const slideMasters = collection(masters)
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', platform, requirements: { isSetSupported: () => api } },
  })
  vi.stubGlobal('PowerPoint', {
    run: (callback: (context: any) => unknown) =>
      callback({
        presentation: { slides, slideMasters, insertSlidesFromBase64: insert },
        sync: async () => {
          if (queued) {
            const action = queued
            queued = undefined
            await action()
          }
        },
      }),
  })
  const runtime = createOfficeHostRuntime('powerpoint', {
    platform,
    presentation: {
      ...binding,
      invalidateQa,
      available: () => paired,
      request: f.request,
      packageBackupAvailable: () => capability,
      packageBackupRequest: f.request,
      readMasterXmlChange: bindingAvailable ? binding.readMasterXmlChange : undefined,
      writeMasterXmlChange: bindingAvailable
        ? async (next: any, expected: any) => {
            if (failPending && next.pending) throw Error('storage_failed')
            await binding.writeMasterXmlChange(next, expected)
          }
        : undefined,
    } as any,
  })
  const input = {
    program: {
      version: 1,
      operations: [
        {
          op: 'replace_xml',
          path: 'ppt/slideMasters/slideMaster1.xml',
          xml: f.originalMaster.replace('name="original"', 'name="edited"'),
        },
      ],
    },
  }
  return {
    runtime,
    binding,
    input,
    insert,
    remove,
    apply,
    invalidateQa,
    request: f.request,
    slides,
    original,
    capability: (v: boolean) => {
      capability = v
    },
    paired: (v: boolean) => {
      paired = v
    },
    api: (v: boolean) => {
      api = v
    },
    failPending: () => {
      failPending = true
    },
  }
}
async function propose(f: Awaited<ReturnType<typeof fixture>>) {
  return f.runtime.skill.executeTool({
    id: 'master',
    name: 'edit_slide_master_xml',
    input: f.input,
  })
}
async function confirm(f: Awaited<ReturnType<typeof fixture>>) {
  const proposal = f.runtime.proposals.pending()!
  await vi.waitFor(() => {
    const current = f.runtime.proposals.pending()
    expect(current && 'lockReview' in current ? current.lockReview?.state : undefined).not.toBe(
      'checking',
    )
  })
  return f.runtime.proposals.confirm(proposal.id)
}
it('prepares without PC uploads, requires confirmation, and cancels without host writes', async () => {
  const f = await fixture()
  const result = await propose(f)
  expect(result.isError, JSON.stringify(result)).not.toBe(true)
  expect(f.runtime.proposals.pending()).toMatchObject({
    operation: 'edit_slide_master_xml',
    preview: { qaScope: { basis: 'master_xml_savepoint', hostSlideIds: ['s0', 's1'] } },
  })
  expect(f.request).not.toHaveBeenCalled()
  expect(f.insert).not.toHaveBeenCalled()
  expect(f.apply).not.toHaveBeenCalled()
  expect(f.remove).not.toHaveBeenCalled()
  f.runtime.clearSession()
  expect(f.runtime.proposals.pending()).toBeUndefined()
  expect(f.binding.listChangeHistory()).toHaveLength(0)
  expect(f.request).not.toHaveBeenCalled()
})
it.each(['capability', 'paired', 'api', 'Mac', 'binding'])(
  'refuses %s without backup or host writes',
  async (mode) => {
    const f = await fixture(mode === 'Mac' ? 'Mac' : 'PC', mode !== 'binding')
    if (mode === 'capability') f.capability(false)
    if (mode === 'paired') f.paired(false)
    if (mode === 'api') f.api(false)
    const result = await propose(f)
    expect(result).toMatchObject(
      mode === 'Mac' ? { isError: true } : { isError: true, mutated: false },
    )
    expect(f.request).not.toHaveBeenCalled()
    expect(f.insert).not.toHaveBeenCalled()
    expect(f.apply).not.toHaveBeenCalled()
    expect(f.runtime.proposals.pending()).toBeUndefined()
  },
)
it('confirms actual SDK stages and per-page layout writes into the durable settings journal', async () => {
  const f = await fixture()
  expect((await propose(f)).isError).not.toBe(true)
  await confirm(f)
  const entry = f.binding.listChangeHistory().find((e) => e.kind === 'master_xml')!
  expect(entry.record).toMatchObject({ state: 'applied', inventoryCleanupVerified: false })
  expect(f.insert).toHaveBeenCalled()
  expect(f.apply).toHaveBeenCalled()
  expect(f.remove).toHaveBeenCalled()
  expect(f.slides.items).toHaveLength(2)
  const result = await f.runtime.skill.executeTool({
    id: 'inspect',
    name: 'inspect_master_xml_change',
    input: { change_id: entry.record.changeId },
  })
  expect(result.isError, JSON.stringify(result)).not.toBe(true)
}, 30000)
it('confirms prepared discard as a local checkpoint with zero host QA or lock hooks', async () => {
  const f = await fixture()
  f.failPending()
  await propose(f)
  await expect(confirm(f)).rejects.toThrow()
  const entry = f.binding.listChangeHistory().find((e) => e.kind === 'master_xml')!
  expect(entry.record.state).toBe('prepared')
  f.invalidateQa.mockClear()
  lockReads.mockClear()
  f.insert.mockClear()
  f.apply.mockClear()
  f.remove.mockClear()
  const result = await f.runtime.skill.executeTool({
    id: 'discard',
    name: 'discard_master_xml_change',
    input: { change_id: entry.record.changeId },
  })
  expect(result.isError, JSON.stringify(result)).not.toBe(true)
  expect(f.runtime.proposals.pending()).toMatchObject({
    operation: 'discard_master_xml_change',
    toolName: 'discard_master_xml_change',
    impact: { host: 'local_checkpoint' },
  })
  await confirm(f)
  expect(f.binding.readMasterXmlChange(entry.record.changeId)?.state).toBe('discarded')
  expect(f.invalidateQa).not.toHaveBeenCalled()
  expect(lockReads).not.toHaveBeenCalled()
  expect(f.insert).not.toHaveBeenCalled()
  expect(f.apply).not.toHaveBeenCalled()
  expect(f.remove).not.toHaveBeenCalled()
})
