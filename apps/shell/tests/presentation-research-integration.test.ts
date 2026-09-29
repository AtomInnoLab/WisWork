import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import JSZip from 'jszip'
import { afterEach, expect, it, vi } from 'vitest'
import { PresentationStore } from '@wiswork/project-store'
import { createPresentationService } from '../src/main/presentation-service'
import { createPresentationResearchSkill } from '../../office-addin/src/skills/powerpoint/presentation-research'
import { createPresentationHostBundleSkill } from '../../office-addin/src/skills/powerpoint/presentation-host-bundle'
import { researchDraftFixture } from '../../office-addin/tests/presentation-research-root-fixture'
import { InMemoryVfs } from '../../office-addin/src/skills/shared/vfs'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const hash = (bytes: string | Uint8Array) => createHash('sha256').update(bytes).digest('hex')
async function setup() {
  const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-research-cross-'))
  roots.push(userDataPath)
  const compile = vi.fn(),
    fetchPage = vi.fn(
      async () =>
        new Response('<html><p>原始研究存在限制</p></html>', {
          headers: { 'content-type': 'text/html' },
        }),
    )
  let service = createPresentationService({ userDataPath, compile, fetchPage })
  let documentId = 'doc'
  const request = vi.fn(
    async (body: unknown, signal?: AbortSignal) =>
      new Response(
        Buffer.from(await service(body, signal ?? new AbortController().signal)).toString(),
      ),
  )
  const raw = async (operation: string, fields: Record<string, unknown> = {}) =>
    JSON.parse(await (await request({ operation, documentId, ...fields })).text())
  const original = Buffer.from('原始研究存在限制'),
    attachmentId = hash(original)
  await raw('attachment_begin', {
    attachmentId,
    sha256: attachmentId,
    name: 'original.txt',
    sizeBytes: original.length,
  })
  await raw('attachment_chunk', { attachmentId, offset: 0, base64: original.toString('base64') })
  await raw('attachment_finish', { attachmentId })
  const draft = researchDraftFixture()
  draft.sources[0]!.uri = `attachment:${attachmentId}`
  const client = () => {
    const vfs = new InMemoryVfs(),
      skill = createPresentationResearchSkill({
        available: () => true,
        request,
        documentId: async () => documentId,
        vfs,
      })
    return { vfs, skill }
  }
  const input = { project_id: 'project', ledger_id: 'research1', expected_revision: 0, draft }
  const call = (skill: ReturnType<typeof client>['skill']) =>
    skill.executeTool({ id: 'build', name: 'build_research_ledger', input })
  return {
    userDataPath,
    compile,
    fetchPage,
    request,
    raw,
    draft,
    input,
    attachmentId,
    client,
    call,
    restart: () => {
      service = createPresentationService({ userDataPath, compile, fetchPage })
    },
    switchDocument: () => {
      documentId = 'other'
    },
  }
}
it('creates research before any plan, keeps conflicts and actual originals, and restores the same result after both restarts', async () => {
  const f = await setup(),
    first = f.client(),
    result = await f.call(first.skill)
  expect(result.isError, result.output).toBeFalsy()
  const value = JSON.parse(result.output)
  expect(value.record).toMatchObject({
    state: 'completed',
    sources: [
      {
        attachmentId: f.attachmentId,
        status: 'found',
        offset: 0,
        provenance: 'user_supplied',
        sha256: f.attachmentId,
      },
    ],
  })
  expect(value.record.sources[0]).not.toHaveProperty('retrievedAt')
  expect(value.record.draft.facts).toEqual(f.draft.facts)
  expect(value.history.records[0].conflictCount).toBe(1)
  expect(await f.raw('status', { projectId: 'project' })).toHaveProperty('error', 'not_found')
  f.restart()
  const reopened = f.client()
  expect(await reopened.skill.latest('project')).toEqual(value.record)
  const repeated = await f.call(reopened.skill)
  expect(JSON.parse(repeated.output)).toEqual(value)
  expect(f.compile).not.toHaveBeenCalled()
  f.switchDocument()
  expect(await f.client().skill.latest('project')).toBeUndefined()
})
it('recovers a saved research conclusion after a lost build response, without repeating the build', async () => {
  const f = await setup(),
    first = f.client(),
    actual = f.request.getMockImplementation()!
  f.request.mockImplementation(async (body, signal) => {
    const response = await actual(body, signal)
    if ((body as { operation: string }).operation === 'research_build')
      throw Error('private lost response')
    return response
  })
  expect((await f.call(first.skill)).isError).toBe(true)
  f.restart()
  const latest = await f.client().skill.latest('project')
  expect(latest?.state).toBe('completed')
  expect(
    f.request.mock.calls.filter(
      ([body]) => (body as { operation: string }).operation === 'research_build',
    ),
  ).toHaveLength(1)
})
it('preserves actual webpage retrieval time and source URI binding separately from organizing time', async () => {
  const f = await setup(),
    uri = 'https://8.8.8.8/study?version=1'
  const imported = await f.raw('attachment_import_webpage', { url: uri })
  expect(imported).toHaveProperty('attachmentId')
  const metadata = await f.raw('attachment_metadata', { attachmentId: imported.attachmentId })
  const draft = {
    ...f.draft,
    sources: [{ ...f.draft.sources[0]!, uri, snapshotAttachmentId: imported.attachmentId }],
  }
  const result = await f
    .client()
    .skill.executeTool({ id: 'web', name: 'build_research_ledger', input: { ...f.input, draft } })
  expect(result.isError, result.output).toBeFalsy()
  const record = JSON.parse(result.output).record
  expect(record.sources[0]).toMatchObject({
    provenance: 'fetched_url_matched',
    retrievedAt: new Date(metadata.retrievedAt).toISOString(),
  })
  const changed = {
    ...draft,
    sources: [{ ...draft.sources[0]!, uri: 'https://8.8.8.8/another?version=1' }],
  }
  const mismatch = await f.client().skill.executeTool({
    id: 'other-url',
    name: 'build_research_ledger',
    input: { ...f.input, ledger_id: 'research2', expected_revision: 2, draft: changed },
  })
  expect(mismatch.isError, mismatch.output).toBeFalsy()
  expect(JSON.parse(mismatch.output).record.sources[0]).toMatchObject({
    status: 'source_mismatch',
    provenance: 'unavailable',
  })
  expect(f.fetchPage).toHaveBeenCalledTimes(1)
})
it('carries the actual pre-plan historical research into a native host package and validates it against PC archive', async () => {
  const f = await setup(),
    research = f.client()
  expect((await f.call(research.skill)).isError).toBeFalsy()
  const plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  plan.projectId = 'project'
  deck.id = 'project'
  const store = new PresentationStore(f.userDataPath)
  store.savePlan('project', 'doc', 0, plan)
  store.beginProduction('project', 'doc', 'run', deck, { revision: 1, plan })
  const native = await new JSZip()
    .file('ppt/slides/slide1.xml', '<title>current edited host</title>')
    .generateAsync({ type: 'uint8array' })
  const vfs = new InMemoryVfs(),
    skill = createPresentationHostBundleSkill({
      available: () => true,
      nativeAvailable: () => true,
      exportDocument: async () => native,
      documentId: async () => 'doc',
      request: f.request,
      vfs,
      readResearch: research.skill.readLatestCompleted,
    })
  const result = await skill.executeTool({
    id: 'package',
    name: 'export_current_presentation_bundle',
    input: { project_id: 'project', request_id: 'run' },
  })
  expect(result.isError, result.output).toBeFalsy()
  const value = JSON.parse(result.output),
    bytes = vfs.readBytes(value.paths[0]),
    zip = await JSZip.loadAsync(bytes)
  expect(JSON.parse(await zip.file('research.json')!.async('string')).draft.facts).toEqual(
    f.draft.facts,
  )
  expect(await zip.file('research.md')!.async('string')).toContain(f.draft.facts[1]!.statement)
  expect(value.receipt.manifest.checks.completion).toBe('not_verified')
  f.restart()
  const restored = createPresentationHostBundleSkill({
    available: () => true,
    nativeAvailable: () => false,
    exportDocument: async () => {
      throw Error('must not export')
    },
    documentId: async () => 'doc',
    request: f.request,
    vfs: new InMemoryVfs(),
  })
  expect(
    (
      await restored.executeTool({
        id: 'restore',
        name: 'restore_presentation_delivery_bundle',
        input: { project_id: 'project', request_id: 'run', bundle_id: value.bundleId },
      })
    ).isError,
  ).toBeFalsy()
  expect(f.compile).not.toHaveBeenCalled()
})
it('retries an archived original ledger after 33 real builds without appending or changing it', async () => {
  const f = await setup(),
    first = f.client(),
    original = JSON.parse((await f.call(first.skill)).output).record
  for (let i = 2; i <= 33; i++) {
    const result = await first.skill.executeTool({
      id: `b${i}`,
      name: 'build_research_ledger',
      input: { ...f.input, ledger_id: `research${i}`, expected_revision: 2 * (i - 1) },
    })
    expect(result.isError, result.output).toBeFalsy()
  }
  f.restart()
  const retried = await f.call(f.client().skill)
  expect(retried.isError, retried.output).toBeFalsy()
  const value = JSON.parse(retried.output)
  expect(value.record).toEqual(original)
  expect(value.history).toMatchObject({ totalRecords: 33, revision: 66 })
  expect(value.history.records.some((r: { id: string }) => r.id === 'research1')).toBe(false)
})
