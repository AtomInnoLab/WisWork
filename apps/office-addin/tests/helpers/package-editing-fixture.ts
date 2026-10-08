import { expect, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { createPresentationPackageBackupService } from '../../../shell/src/main/presentation-package-backups.js'
import { createPresentationPackageEditingSkill } from '../../src/skills/powerpoint/presentation-package-editing.js'
import {
  validatePresentationPackageChange,
  validPackageTransition,
  type PresentationPackageChange,
} from '../../src/skills/powerpoint/presentation-package-change.js'
import {
  presentationPackageDigest,
  type XmlReplacement,
} from '../../src/skills/powerpoint/powerpoint-package.js'
import { createStructuredProposalController } from '../../src/agent/proposal-controller.js'
const roots: string[] = []
export function cleanupPackageFixtures() {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
}
export async function packageEditingFixture(count = 2) {
  const root = mkdtempSync(join(tmpdir(), 'package-editing-'))
  roots.push(root)
  const service = createPresentationPackageBackupService({ userDataPath: root })
  const zip = new JSZip()
  zip.file(
    'ppt/presentation.xml',
    '<p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldIdLst><p:sldId id="256" r:id="r1"/></p:sldIdLst></p:presentation>',
  )
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    '<Relationships><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>',
  )
  const xml =
    '<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><a:t>Original</a:t></p:spTree></p:cSld></p:sld>'
  zip.file('ppt/slides/slide1.xml', xml)
  for (let i = 1; i <= 32; i++)
    zip.file(`ppt/charts/style${i}.xml`, '<c:style xmlns:c="urn:c" val="1"/>')
  zip.file('ppt/media/image1.bin', 'protected-image')
  const base64 = await zip.generateAsync({ type: 'base64' })
  let doc = 'doc',
    available = true,
    mode = '',
    writeMode = '',
    index = 0,
    beforeSdk: () => void = () => {},
    readHook: () => void = () => {}
  const order = Array.from({ length: count }, (_, i) => `s${i}`),
    packages = new Map(order.map((id) => [id, base64])),
    data = new Map<string, PresentationPackageChange>()
  const request = vi.fn(
    async (body: unknown, signal?: AbortSignal) =>
      new Response(
        JSON.stringify(
          await service(body as Record<string, unknown>, signal ?? new AbortController().signal),
        ),
      ),
  )
  const snapshot = async (exports: string[] = []) => ({
    slideIds: [...order],
    pages: await Promise.all(
      order.map(async (slideId) => ({
        slideId,
        digest: await presentationPackageDigest(packages.get(slideId)!),
        ...(exports.includes(slideId) ? { base64: packages.get(slideId)! } : {}),
      })),
    ),
  })
  const adapter = {
    inspect: vi.fn(async (exports: string[] = []) => {
      readHook()
      return snapshot(exports)
    }),
    stage: vi.fn(
      async (
        input: any,
        onInserted: (id: string) => Promise<void>,
        beforeWrite: () => Promise<void>,
        guard: () => void,
      ) => {
        const owned = structuredClone(input)
        await beforeWrite()
        expect(await snapshot()).toEqual(owned.preimage)
        beforeSdk()
        guard()
        if (mode === 'stage_before') throw Error('sync_failed')
        const id = `import-${++index}`
        order.splice(order.indexOf(owned.sourceSlideId) + 1, 0, id)
        packages.set(id, owned.base64)
        if (mode === 'stage_after' || mode === 'restore_after') throw Error('ack_lost')
        await onInserted(id)
        if (mode === 'callback_after') throw Error('ack_lost')
        return { slideId: id }
      },
    ),
    remove: vi.fn(async (input: any, beforeWrite: () => Promise<void>, guard: () => void) => {
      await beforeWrite()
      expect(await snapshot()).toEqual(input.preimage)
      beforeSdk()
      guard()
      if (mode === 'remove_before') throw Error('sync_failed')
      order.splice(order.indexOf(input.slideId), 1)
      packages.delete(input.slideId)
      if (mode === 'remove_after') throw Error('ack_lost')
    }),
    screenshotSlide: vi.fn(async (i: number) => ({
      slideId: order[i]!,
      mime: 'image/png' as const,
      base64:
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII=',
    })),
  }
  const write = vi.fn(
    async (next: PresentationPackageChange, expected: PresentationPackageChange | undefined) => {
      expect(validatePresentationPackageChange(next)).toBe(true)
      expect(validPackageTransition(expected, next)).toBe(true)
      expect(data.get(next.changeId)).toEqual(expected)
      if (writeMode === 'before_pending' && next.pending && !expected?.pending)
        throw Error('settings_failed')
      data.set(next.changeId, structuredClone(next))
      if (writeMode === 'observed_ack' && next.pending?.afterProofRef)
        throw Error('settings_ack_lost')
    },
  )
  const proposals = createStructuredProposalController()
  const make = () =>
    createPresentationPackageEditingSkill({
      documentId: async () => doc,
      assertDocumentId: (expected) => {
        if (expected !== doc) throw Error('presentation_document_changed')
      },
      available: () => available,
      adapter,
      request,
      proposals,
      readPackageChange: (id) => data.get(id),
      writePackageChange: write,
    })
  let skill = make()
  const replacements: XmlReplacement[] = [
    { path: 'ppt/slides/slide1.xml', xml: xml.replace('Original', 'Edited') },
  ]
  return {
    adapter,
    data,
    request,
    write,
    order,
    packages,
    replacements,
    proposals,
    propose: (
      kind: 'slide' | 'chart' = 'slide',
      slideIndex = 0,
      r = replacements,
      signal?: AbortSignal,
    ) => skill.propose(kind, slideIndex, r, undefined, signal),
    confirm: () => proposals.confirm(proposals.pending()!.id),
    tool: (action: string, input: any) =>
      skill.executeTool({ name: `${action}_package_xml_change`, input }),
    reopen: () => {
      skill = make()
    },
    mode: (v: string) => {
      mode = v
    },
    writeMode: (v: string) => {
      writeMode = v
    },
    setBeforeSdk: (fn: () => void) => {
      beforeSdk = fn
    },
    setRead: (fn: () => void) => {
      readHook = fn
    },
    setDocument: (v: string) => {
      doc = v
    },
    disconnect: () => {
      available = false
    },
    clear: () => skill.clear(),
    base64,
  }
}
