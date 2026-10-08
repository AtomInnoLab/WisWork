import { afterEach, describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createPresentationService } from '../../shell/src/main/presentation-service.js'
import { createPresentationPackageEditingSkill } from '../src/skills/powerpoint/presentation-package-editing.js'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import type { PackageHostSnapshot } from '../src/skills/powerpoint/browser-presentation-package-edit-adapter.js'
import type { PowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter.js'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill.js'
import type { StructuredProposalRequest } from '../src/agent/proposal-controller.js'
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
] as const

const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})

async function durableFixture(name: string, path: string, before: string, after: string) {
  const zip = new JSZip()
  zip.file(
    'ppt/presentation.xml',
    '<p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldIdLst><p:sldId id="256" r:id="r1"/></p:sldIdLst></p:presentation>',
  )
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    '<Relationships><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>',
  )
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
  zip.file(path, before)
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships/>')
  zip.file('docProps/core.xml', '<core value="original"/>')
  const base64 = await zip.generateAsync({ type: 'base64' })
  const observed = { slideId: 'source', base64, fingerprint: 'volatile' }
  const order = Array.from({ length: 600 }, (_, index) => (index === 599 ? 'source' : `s${index}`))
  const packages = new Map(order.map((id) => [id, base64]))
  // Test-only memoization of immutable package strings; every one of the 600 entries remains in each proof.
  const digests = new Map<string, Promise<string>>()
  const digest = (bytes: string) => {
    let value = digests.get(bytes)
    if (!value) {
      value = presentationPackageDigest(bytes)
      digests.set(bytes, value)
    }
    return value
  }
  let live: PackageHostSnapshot | undefined,
    aliasArmed = false,
    mutateAfterInspect = false
  const snapshot = async (exports: string[] = []) => {
    const result: PackageHostSnapshot = {
      slideIds: order.map((id) => (id === 'source' ? observed.slideId : id)),
      pages: await Promise.all(
        order.map(async (originalId) => {
          const slideId = originalId === 'source' ? observed.slideId : originalId,
            bytes = originalId === 'source' ? observed.base64 : packages.get(originalId)!
          return {
            slideId,
            digest: await digest(bytes),
            ...(exports.includes(slideId) ? { base64: bytes } : {}),
          }
        }),
      ),
    }
    live = result
    return result
  }
  const packageAdapter = {
    inspect: vi.fn(async (exports: string[] = []) => {
      const value = await snapshot(exports)
      if (aliasArmed) {
        aliasArmed = false
        mutateAfterInspect = true
      }
      return value
    }),
    stage: vi.fn(
      async (
        input: {
          base64: string
          sourceSlideId: string
          packageSourceSlideId: string
          preimage: PackageHostSnapshot
        },
        onInserted: (id: string) => Promise<void>,
        beforeWrite: () => Promise<void>,
        guard: () => void,
      ) => {
        const owned = structuredClone(input)
        await beforeWrite()
        expect(await snapshot()).toEqual(owned.preimage)
        guard()
        expect(owned.sourceSlideId).toBe('source')
        expect(owned.packageSourceSlideId).toBe('256#')
        order.splice(order.indexOf(owned.sourceSlideId) + 1, 0, 'inserted')
        packages.set('inserted', owned.base64)
        await onInserted('inserted')
        return { slideId: 'inserted' }
      },
    ),
    remove: vi.fn(
      async (
        input: { slideId: string; preimage: PackageHostSnapshot },
        beforeWrite: () => Promise<void>,
        guard: () => void,
      ) => {
        await beforeWrite()
        expect(await snapshot()).toEqual(input.preimage)
        guard()
        order.splice(order.indexOf(input.slideId), 1)
        packages.delete(input.slideId)
      },
    ),
  }
  const values = new Map<string, string>()
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
  const root = mkdtempSync(join(tmpdir(), 'proposal-preimage-'))
  roots.push(root)
  const service = createPresentationService({ userDataPath: root })
  const actualProposals = createStructuredProposalController()
  let request!: StructuredProposalRequest
  const proposals = {
    ...actualProposals,
    propose: (value: StructuredProposalRequest) => {
      request = value
      return actualProposals.propose(value)
    },
  }
  const durable = createPresentationPackageEditingSkill({
    documentId: async () => {
      if (mutateAfterInspect) {
        mutateAfterInspect = false
        observed.slideId = 'other'
        live!.slideIds[599] = 'other'
        live!.pages[599]!.slideId = 'other'
      }
      return binding.documentId()
    },
    assertDocumentId: binding.assertDocumentId,
    available: () => true,
    adapter: packageAdapter,
    proposals,
    request: async (body, signal) =>
      new Response(
        new Uint8Array(await service(body, signal ?? new AbortController().signal)).buffer,
      ),
    readPackageChange: binding.readPackageChange,
    writePackageChange: binding.writePackageChange,
  })
  const replaceSlidePackage = vi.fn(async () => ({ slideId: 'inserted' }))
  const adapter = {
    verifySlides: vi.fn(),
    exportSlidePackage: vi.fn(),
    readSlideOrder: vi.fn(async () => order),
    replaceSlidePackage,
  } as unknown as PowerPointAdapter
  const skill = createPowerPointSkill({ adapter, proposals, durablePackage: durable.propose })
  const result = await skill.executeTool({
    id: 'call',
    name,
    input: {
      slide_index: 599,
      program: { version: 1, operations: [{ op: 'replace_xml', path, xml: after }] },
    },
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(request).toBeDefined()
  return {
    zip,
    base64,
    observed,
    order,
    adapter,
    replaceSlidePackage,
    request,
    slideIndex: 599,
    packageAdapter,
    armAlias: () => {
      aliasArmed = true
    },
    binding,
  }
}

const fixture = durableFixture

describe.each(cases)('%s exact package proposal', (name, path, before, after) => {
  it('keeps unchanged full order and passes exact preimage without narrowing large slide indices', async () => {
    const f = await fixture(name, path, before, after)
    expect(await f.request.validate()).toBe(true)
    await f.request.execute()
    if ('packageAdapter' in f) {
      expect(f.packageAdapter.stage).toHaveBeenCalledOnce()
      expect(f.packageAdapter.stage.mock.calls[0]![0]).toMatchObject({
        sourceSlideId: 'source',
        packageSourceSlideId: '256#',
        preimage: {
          slideIds: Array.from({ length: 600 }, (_, i) => (i === 599 ? 'source' : `s${i}`)),
        },
      })
      expect(f.packageAdapter.remove).toHaveBeenCalledOnce()
      expect(f.order).toHaveLength(600)
      expect(f.order.slice(0, 599)).toEqual(Array.from({ length: 599 }, (_, i) => `s${i}`))
      expect(f.order[599]).toBe('inserted')
      const record = f.binding.listChangeHistory().find((entry) => entry.kind === 'package_xml')!
      expect(record.record).toMatchObject({
        state: 'applied',
        sourceSlideId: 'source',
        replacementSlideId: 'inserted',
      })
      expect(f.replaceSlidePackage).not.toHaveBeenCalled()
      return
    }
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
      if ('packageAdapter' in f) {
        expect(f.packageAdapter.stage).not.toHaveBeenCalled()
        expect(f.packageAdapter.remove).not.toHaveBeenCalled()
      }
    },
  )

  it('copies the observed object before later awaited order reads can mutate its aliases', async () => {
    const f = await fixture(name, path, before, after)
    if ('packageAdapter' in f) f.armAlias()

    expect(await f.request.validate()).toBe(true)
    await expect(f.request.execute()).rejects.toThrow('proposal_stale')
    expect(f.replaceSlidePackage).not.toHaveBeenCalled()
    if ('packageAdapter' in f) {
      expect(f.packageAdapter.stage).not.toHaveBeenCalled()
      expect(f.packageAdapter.remove).not.toHaveBeenCalled()
    }
  })
})
