import { afterEach, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createOfficeHostRuntime } from '../src/agent/host-runtime'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { createPresentationService } from '../../shell/src/main/presentation-service'
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
const roots: string[] = []
afterEach(() => {
  vi.unstubAllGlobals()
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
async function fixture() {
  const zip = new JSZip()
  zip.file(
    'ppt/presentation.xml',
    '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>',
  )
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>',
  )
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="p"><p:cSld/></p:sld>')
  const base64 = await zip.generateAsync({ type: 'base64' }),
    insert = vi.fn(),
    remove = vi.fn(),
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
  let failPending = false,
    failInsert = false,
    currentProject: string | undefined
  const invalidateQa = vi.fn(binding.invalidateQa)
  let queued: (() => void) | undefined
  const page = (id: string, packageBytes: string) => ({
    id,
    exportAsBase64: () => ({ value: packageBytes }),
    delete: () => {
      remove(id)
      queued = () => {
        slides.items = slides.items.filter((item) => item.id !== id)
      }
    },
  })
  const slides = {
    items: [page('host-source', base64)],
    load: vi.fn(),
    getCount: () => ({ value: slides.items.length }),
    getItem: (id: string) => slides.items.find((s) => s.id === id)!,
  }
  insert.mockImplementation((packageBytes: string, options: { targetSlideId: string }) => {
    queued = () => {
      slides.items.splice(
        slides.items.findIndex((item) => item.id === options.targetSlideId) + 1,
        0,
        page('host-replacement', packageBytes),
      )
    }
  })
  let capability = true,
    paired = true,
    api = true
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: () => api } },
  })
  vi.stubGlobal('PowerPoint', {
    run: (callback: (ctx: unknown) => unknown) =>
      callback({
        presentation: { slides, insertSlidesFromBase64: insert },
        sync: async () => {
          if (queued) {
            const write = queued
            queued = undefined
            if (failInsert) throw Error('office_write_failed')
            write()
          }
        },
      }),
  })
  const root = mkdtempSync(join(tmpdir(), 'runtime-package-'))
  roots.push(root)
  const service = createPresentationService({ userDataPath: root })
  const request = vi.fn(
    async (body: unknown, signal?: AbortSignal) =>
      new Response(
        new Uint8Array(await service(body, signal ?? new AbortController().signal)).buffer,
      ),
  )
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      ...binding,
      invalidateQa,
      lastProject: () => currentProject,
      writePackageChange: async (
        record: Parameters<typeof binding.writePackageChange>[0],
        expected: Parameters<typeof binding.writePackageChange>[1],
      ) => {
        if (failPending && record.pending) throw Error('storage_failed')
        await binding.writePackageChange(record, expected)
      },
      available: () => paired,
      request,
      packageBackupAvailable: () => capability,
      packageBackupRequest: request,
    } as any,
  })
  return {
    runtime,
    request,
    binding,
    insert,
    remove,
    slides,
    invalidateQa,
    setFailPending: () => {
      failPending = true
    },
    setFailInsert: () => {
      failInsert = true
    },
    setProject: (value: string) => {
      currentProject = value
    },
    setCapability: (v: boolean) => {
      capability = v
    },
    setPaired: (v: boolean) => {
      paired = v
    },
    setApi: (v: boolean) => {
      api = v
    },
  }
}
const input = {
  slide_index: 0,
  program: {
    version: 1,
    operations: [
      {
        op: 'replace_xml',
        path: 'ppt/slides/slide1.xml',
        xml: '<p:sld xmlns:p="p"><p:cSld><p:spTree/></p:cSld></p:sld>',
      },
    ],
  },
}
it('prepares the real XML proposal without PC or host writes before confirmation', async () => {
  const f = await fixture()
  const result = await f.runtime.skill.executeTool({ id: 'xml', name: 'edit_slide_xml', input })
  expect(result.isError, result.output).not.toBe(true)
  const pending = f.runtime.proposals.pending()
  expect(pending?.operation).toBe('edit_slide_xml')
  expect(pending && 'preview' in pending ? pending.preview : undefined).toMatchObject({
    qaScope: { basis: 'package_xml_savepoint', hostSlideIds: ['host-source'] },
  })
  expect(f.request).not.toHaveBeenCalled()
  expect(f.insert).not.toHaveBeenCalled()
  expect(f.remove).not.toHaveBeenCalled()
  f.runtime.clearSession()
  expect(f.runtime.proposals.pending()).toBeUndefined()
  expect(f.request).not.toHaveBeenCalled()
})
it.each(['capability', 'paired', 'api'])(
  'refuses missing %s before PC preparation or host writes',
  async (mode) => {
    const f = await fixture()
    if (mode === 'capability') f.setCapability(false)
    if (mode === 'paired') f.setPaired(false)
    if (mode === 'api') f.setApi(false)
    const result = await f.runtime.skill.executeTool({ id: 'xml', name: 'edit_slide_xml', input })
    expect(result).toMatchObject({
      isError: true,
      mutated: false,
      output: 'presentation_package_persistence_unavailable',
    })
    expect(f.request).not.toHaveBeenCalled()
    expect(f.insert).not.toHaveBeenCalled()
    expect(f.runtime.proposals.pending()).toBeUndefined()
  },
)
it('dispatches recovery tools to the actual engine with safe missing-record errors', async () => {
  const f = await fixture()
  expect(f.runtime.skill.tools.some((tool) => tool.name === 'inspect_package_xml_change')).toBe(
    true,
  )
  const result = await f.runtime.skill.executeTool({
    id: 'inspect',
    name: 'inspect_package_xml_change',
    input: { change_id: 'missing' },
  })
  expect(result).toMatchObject({
    isError: true,
    mutated: false,
    summary: 'presentation_package_missing',
  })
  expect(f.insert).not.toHaveBeenCalled()
})

it('confirms the actual durable XML engine through Runtime mutation hooks and records only actual SDK IDs', async () => {
  const f = await fixture()
  await f.runtime.skill.executeTool({ id: 'xml', name: 'edit_slide_xml', input })
  const pending = f.runtime.proposals.pending()!
  expect(f.insert).not.toHaveBeenCalled()
  await vi.waitFor(() => {
    const current = f.runtime.proposals.pending()
    expect(current && 'lockReview' in current ? current.lockReview?.state : undefined).not.toBe(
      'checking',
    )
  })
  await f.runtime.proposals.confirm(pending.id)
  expect(f.insert).toHaveBeenCalledTimes(1)
  expect(f.remove).toHaveBeenCalledExactlyOnceWith('host-source')
  expect(f.slides.items.map((page) => page.id)).toEqual(['host-replacement'])
  const entry = f.binding.listChangeHistory().find((entry) => entry.kind === 'package_xml')!
  expect(entry.record).toMatchObject({
    state: 'applied',
    sourceSlideId: 'host-source',
    replacementSlideId: 'host-replacement',
  })
  const inspected = await f.runtime.skill.executeTool({
    id: 'inspect',
    name: 'inspect_package_xml_change',
    input: { change_id: entry.record.changeId },
  })
  expect(inspected.isError, inspected.output).not.toBe(true)
})
it('audits package backups through the PC inventory against current Office history', async () => {
  const f = await fixture()
  await f.runtime.skill.executeTool({ id: 'xml', name: 'edit_slide_xml', input })
  const pending = f.runtime.proposals.pending()!
  await vi.waitFor(() => {
    const current = f.runtime.proposals.pending()
    expect(current && 'lockReview' in current ? current.lockReview?.state : undefined).not.toBe(
      'checking',
    )
  })
  await f.runtime.proposals.confirm(pending.id)
  await f.runtime.changes!.refresh()
  expect(f.runtime.changes!.snapshot().packageBackupAudit).toMatchObject({ unmatched: 0 })
  const active = f.runtime.changes!.snapshot().packageBackupAudit!.active
  expect(active).toBeGreaterThan(0)
  await f.request({
    operation: 'package_backup_begin',
    documentId: await f.binding.documentId(),
    changeId: 'orphan-upload',
    key: 'page-0',
    sha256: 'a'.repeat(64),
    sizeBytes: 1,
  })
  await f.runtime.changes!.refresh()
  expect(f.runtime.changes!.snapshot().packageBackupAudit).toEqual({
    active: active + 1,
    unmatched: 1,
  })
})

it.each(['discard', 'reconcile'] as const)(
  'confirms metadata-only %s without host writes, QA invalidation or native mutation lock reads',
  async (action) => {
    const f = await fixture()
    if (action === 'discard') f.setFailPending()
    else f.setFailInsert()
    await f.runtime.skill.executeTool({ id: 'xml', name: 'edit_slide_xml', input })
    const initial = f.runtime.proposals.pending()!
    await vi.waitFor(() => {
      const current = f.runtime.proposals.pending()
      expect(current && 'lockReview' in current ? current.lockReview?.state : undefined).not.toBe(
        'checking',
      )
    })
    await expect(f.runtime.proposals.confirm(initial.id)).rejects.toThrow()
    const entry = f.binding.listChangeHistory().find((entry) => entry.kind === 'package_xml')!
    expect(entry.record.state).toBe('prepared')
    expect(f.slides.items.map((page) => page.id)).toEqual(['host-source'])
    f.setProject('locks-project')
    f.invalidateQa.mockClear()
    lockReads.mockClear()
    f.request.mockClear()
    f.insert.mockClear()
    f.remove.mockClear()
    const toolName = `${action}_package_xml_change`
    const result = await f.runtime.skill.executeTool({
      id: 'metadata',
      name: toolName,
      input: { change_id: entry.record.changeId },
    })
    expect(result.isError, result.output).not.toBe(true)
    const proposal = f.runtime.proposals.pending()!
    expect(proposal).toMatchObject({
      operation: toolName,
      toolName,
      impact: { host: 'local_checkpoint' },
    })
    await vi.waitFor(() => {
      const current = f.runtime.proposals.pending()
      expect(current && 'lockReview' in current ? current.lockReview?.state : undefined).not.toBe(
        'checking',
      )
    })
    await f.runtime.proposals.confirm(proposal.id)
    expect(f.binding.readPackageChange(entry.record.changeId)?.state).toBe('discarded')
    expect(f.invalidateQa).not.toHaveBeenCalled()
    expect(f.insert).not.toHaveBeenCalled()
    expect(f.remove).not.toHaveBeenCalled()
    expect(lockReads).not.toHaveBeenCalled()
  },
)
