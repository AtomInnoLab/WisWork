import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import JSZip from 'jszip'
import { afterEach, expect, it, vi } from 'vitest'
import { PresentationStore } from '@wiswork/project-store'
import { presentationPlanClaims } from '@wiswork/pptx-engine/presentation-plan'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
import { createPresentationHostBundleSkill } from '../../office-addin/src/skills/powerpoint/presentation-host-bundle'
import {
  exportPowerPointDocument,
  supportsPowerPointDocumentExport,
} from '../../office-addin/src/skills/powerpoint/presentation-document-export'
import { InMemoryVfs } from '../../office-addin/src/skills/shared/vfs'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
  vi.unstubAllGlobals()
})
const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex')
async function setup() {
  const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-host-bundle-cross-'))
  roots.push(userDataPath)
  const plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  deck.claims = presentationPlanClaims(plan)
  const base = { documentId: 'doc', projectId: deck.id, requestId: 'run' }
  const store = new PresentationStore(userDataPath)
  store.savePlan(base.projectId, base.documentId, 0, plan)
  store.beginProduction(base.projectId, base.documentId, base.requestId, deck, {
    revision: 1,
    plan,
  })
  const native = await new JSZip()
    .file('ppt/slides/slide1.xml', '<title>用户手工修改，当前原生宿主内容</title>')
    .generateAsync({ type: 'uint8array' })
  const close = vi.fn((callback: (value: unknown) => void) => callback({ status: 'succeeded' }))
  const getFile = vi.fn((_format: string, _options: unknown, callback: (value: unknown) => void) =>
    callback({
      status: 'succeeded',
      value: {
        size: native.length,
        sliceCount: 1,
        getSliceAsync: (index: number, cb: (value: unknown) => void) =>
          cb({
            status: 'succeeded',
            value: { index, size: native.length, data: Array.from(native) },
          }),
        closeAsync: close,
      },
    }),
  )
  vi.stubGlobal('Office', {
    context: {
      host: 'PowerPoint',
      document: { getFileAsync: getFile },
      requirements: { isSetSupported: () => true },
    },
    FileType: { Compressed: 'compressed', Pdf: 'pdf' },
  })
  const compile = vi.fn()
  let pc = createPresentationService({ userDataPath, compile })
  const request = vi.fn(
    async (body: unknown, signal?: AbortSignal) =>
      new Response(Buffer.from(await pc(body, signal ?? new AbortController().signal)).toString()),
  )
  const client = () => {
    const vfs = new InMemoryVfs()
    const skill = createPresentationHostBundleSkill({
      available: () => true,
      nativeAvailable: supportsPowerPointDocumentExport,
      exportDocument: exportPowerPointDocument,
      documentId: async () => base.documentId,
      request,
      vfs,
    })
    return { skill, vfs }
  }
  const input = { project_id: base.projectId, request_id: base.requestId }
  const call = (
    skill: ReturnType<typeof client>['skill'],
    name = 'export_current_presentation_bundle',
    extra = {},
  ) => skill.executeTool({ id: 'x', name, input: { ...input, ...extra } })
  return {
    base,
    native,
    close,
    getFile,
    compile,
    request,
    client,
    call,
    restart: () => {
      pc = createPresentationService({ userDataPath, compile })
    },
    raw: async (operation: string, extra = {}) =>
      JSON.parse(await (await request({ operation, ...base, ...extra })).text()),
  }
}
it('exports edited native bytes through actual SDK callback and PC persistence, restores across both restarts', async () => {
  const f = await setup(),
    first = f.client()
  const result = await f.call(first.skill)
  expect(result.isError, result.output).toBeFalsy()
  const value = JSON.parse(result.output),
    zipBytes = first.vfs.readBytes(value.paths[0])
  expect(value.bundleId).toBe(hash(zipBytes))
  expect(f.close).toHaveBeenCalledTimes(1)
  const zip = await JSZip.loadAsync(zipBytes)
  expect(await zip.file('presentation.pptx')!.async('uint8array')).toEqual(f.native)
  expect(JSON.parse(await zip.file('quality.json')!.async('string')).checks.completion).toBe(
    'not_verified',
  )
  f.restart()
  const list = await f.raw('delivery_bundle_list')
  expect(list.bundles).toHaveLength(1)
  expect(list.bundles[0]).toEqual(value.receipt)
  const reopened = f.client()
  const restored = await f.call(reopened.skill, 'restore_presentation_delivery_bundle', {
    bundle_id: value.bundleId,
  })
  expect(restored.isError, restored.output).toBeFalsy()
  expect(reopened.vfs.readBytes(JSON.parse(restored.output).paths[0])).toEqual(zipBytes)
  expect(f.getFile).toHaveBeenCalledTimes(1)
  expect(f.compile).not.toHaveBeenCalled()
  expect(
    await f.raw('delivery_bundle_metadata', {
      bundleId: value.bundleId,
      documentId: 'another-document',
    }),
  ).toHaveProperty('error')
})
it('recovers a completed PC package after a lost finish response without native export replay', async () => {
  const f = await setup(),
    client = f.client(),
    actual = f.request.getMockImplementation()!
  f.request.mockImplementation(async (body, signal) => {
    const response = await actual(body, signal)
    if ((body as { operation: string }).operation === 'delivery_bundle_finish')
      throw Error('lost-response')
    return response
  })
  expect((await f.call(client.skill)).isError).toBe(true)
  expect(client.vfs.list('/home/user')).toEqual([])
  f.restart()
  const list = await f.raw('delivery_bundle_list')
  expect(list.bundles[0].state).toBe('ready')
  const reopened = f.client()
  const restored = await f.call(reopened.skill, 'restore_presentation_delivery_bundle', {
    bundle_id: list.bundles[0].bundleId,
  })
  expect(restored.isError, restored.output).toBeFalsy()
  expect(f.getFile).toHaveBeenCalledTimes(1)
})
