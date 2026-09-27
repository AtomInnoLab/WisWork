import { expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { createPresentationExistingPageEditingSkill } from '../src/skills/powerpoint/presentation-existing-page-editing'
import { inspectPowerPointPicturePackage } from '../src/skills/powerpoint/powerpoint-package'
import { InMemoryVfs } from '../src/skills/shared/vfs'
import { readBoundedImage } from '../src/skills/shared/import-media'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import type { PresentationExistingPageChange } from '../src/skills/powerpoint/presentation-existing-page'

it('prepares a native picture revision for the confirmed page transaction without writing the host', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const source = Buffer.from((await compilePresentationDeck(deck)).bytes).toString('base64')
  const zip = await JSZip.loadAsync(source, { base64: true })
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const shapeId = xml.match(/<p:pic\b[^]*?<p:cNvPr\b[^>]*\bid="(\d+)"/)![1]!
  const original = await inspectPowerPointPicturePackage(source, shapeId)
  const vfs = new InMemoryVfs()
  const png = new PNG({ width: 2, height: 1 })
  png.data[0] = 255
  vfs.writeFile('/home/user/new.png', PNG.sync.write(png))
  vi.stubGlobal('createImageBitmap', async () => ({ width: 2, height: 1, close() {} }))
  expect((await readBoundedImage(vfs, '/home/user/new.png')).mime).toBe('image/png')
  let hostWrites = 0
  let slideIds = ['old']
  let corruptNew = false
  const records = new Map<string, PresentationExistingPageChange>()
  const backup = { backupId: '', documentId: 'doc', hostSlideId: 'old', slideIds: ['old'], sha256: '', sizeBytes: 0, receivedBytes: 0, status: 'uploading' }
  let backupBytes = new Uint8Array()
  const request = async (body: unknown) => {
    const input = body as Record<string, unknown>
    if (input.operation === 'existing_page_backup_begin') {
      Object.assign(backup, input, { receivedBytes: 0, status: 'uploading' })
      backupBytes = new Uint8Array(backup.sizeBytes)
    } else if (input.operation === 'existing_page_backup_chunk') {
      const chunk = Buffer.from(input.base64 as string, 'base64')
      backupBytes.set(chunk, input.offset as number)
      backup.receivedBytes += chunk.length
    } else if (input.operation === 'existing_page_backup_finish') backup.status = 'ready'
    else if (input.operation === 'existing_page_backup_read') return new Response(JSON.stringify({
      backupId: backup.backupId, offset: input.offset, sizeBytes: backup.sizeBytes, sha256: backup.sha256,
      base64: Buffer.from(backupBytes.subarray(input.offset as number, (input.offset as number) + (input.length as number))).toString('base64'),
    }))
    return new Response(JSON.stringify(backup))
  }
  const baseline = {
    baselineId: 'baseline', documentId: 'doc', contentDigest: 'a'.repeat(64),
    scope: { kind: 'current', slideIds: ['old'] },
    context: { slideIds: ['old'], selectedSlideIds: ['old'], selectedShapeIds: [] },
    pages: [{ slideId: 'old', shapes: [] }],
  }
  const options = {
    baseline: { snapshot: () => structuredClone(baseline), executeTool: async () => ({ output: '{"unchanged":true}', mutated: false }) },
    adapter: {
      inspect: async () => ({ status: slideIds.length === 2 ? 'staged' : slideIds[0] === 'old' ? 'baseline' : slideIds[0] === 'new' ? 'applied' : 'undone', slideIds: [...slideIds] }),
      stage: async (_record: unknown, revision: string, onInserted: (id: string) => Promise<void>) => {
        expect(revision).toBe(changed)
        expect(backup.status).toBe('ready')
        hostWrites++
        slideIds = ['old', 'new']
        await onInserted('new')
      },
      commit: async () => { hostWrites++; slideIds = ['new'] },
      undo: async (_record: unknown, original: string, onRestored: (id: string) => Promise<void>) => {
        expect(original).toBe(source)
        hostWrites++
        slideIds = ['new', 'restored']
        await onRestored('restored')
        slideIds = ['restored']
      },
    },
    inspectPage: async (slideId: string) => ({ slideId, shapesTruncated: false, screenshot: {
      mime: 'image/png', base64: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII=',
    } }),
    exportAdapter: { exportPresentationPagePackage: async (slideId: string) => ({ slideId, slideIds: [...slideIds], base64: slideId === 'new' && !corruptNew ? changed : source }) },
    vfs, request, proposals: createStructuredProposalController(),
    documentId: async () => 'doc', readExistingPageChange: (id: string) => records.get(id),
    writeExistingPageChange: async (record: PresentationExistingPageChange) => { records.set(record.changeId, structuredClone(record)) }, available: () => true,
  } as unknown as Parameters<typeof createPresentationExistingPageEditingSkill>[0]
  const skill = createPresentationExistingPageEditingSkill(options)
  const missing = await skill.executeTool({ id: 'missing', name: 'prepare_existing_presentation_image_revision', input: {
    baseline_id: 'baseline', slide_id: 'old', shape_id: '999999', path: '/home/user/new.png',
  } })
  expect(missing.isError).toBe(true)
  expect(hostWrites).toBe(0)
  const result = await skill.executeTool({ id: 'prepare', name: 'prepare_existing_presentation_image_revision', input: {
    baseline_id: 'baseline', slide_id: 'old', shape_id: shapeId, path: '/home/user/new.png',
  } })
  expect(result.isError, result.output).not.toBe(true)
  const prepared = JSON.parse(result.output) as { path: string; beforeDigest: string; afterDigest: string; nextTool: string }
  expect(prepared.path).toMatch(/^\/home\/user\/presentation-image-revision-.*\.pptx$/)
  expect(prepared.beforeDigest).not.toBe(prepared.afterDigest)
  expect(prepared.nextTool).toBe('stage_existing_presentation_page_change')
  const changed = Buffer.from(vfs.readBytes(prepared.path, { maxBytes: 8 * 1024 * 1024 })).toString('base64')
  expect((await inspectPowerPointPicturePackage(changed, shapeId)).mediaDigest).not.toBe(original.mediaDigest)
  const staged = await skill.executeTool({ id: 'stage', name: 'stage_existing_presentation_page_change', input: {
    baseline_id: 'baseline', slide_id: 'old', path: prepared.path, picture_shape_id: shapeId,
  } })
  expect(staged.isError, staged.output).not.toBe(true)
  expect(JSON.parse(staged.output)).toMatchObject({ status: 'awaiting_confirmation' })
  expect(hostWrites).toBe(0)
  const proposals = options.proposals
  const confirm = async () => {
    const proposalId = proposals.pending()!.id
    const decision = proposals.waitForDecision(proposalId)
    await proposals.confirm(proposalId)
    return decision
  }
  expect((await confirm()).status).toBe('confirmed')
  expect(hostWrites).toBe(1)
  expect(Buffer.from(backupBytes).toString('base64')).toBe(source)
  const changeId = [...records.keys()][0]!
  expect(records.get(changeId)?.state).toBe('staged')
  expect(records.get(changeId)?.pictureTarget?.afterDigest).toBe((await inspectPowerPointPicturePackage(changed, shapeId)).mediaDigest)
  const commit = await skill.executeTool({ id: 'commit', name: 'commit_existing_presentation_page_change', input: { change_id: changeId } })
  expect(commit.isError, commit.output).not.toBe(true)
  corruptNew = true
  expect(await confirm()).toMatchObject({ status: 'confirmed', postWrite: { status: 'unavailable', reason: 'picture_readback_failed' } })
  corruptNew = false
  expect(records.get(changeId)?.state).toBe('applied')
  const undo = await skill.executeTool({ id: 'undo', name: 'undo_existing_presentation_page_change', input: { change_id: changeId } })
  expect(undo.isError, undo.output).not.toBe(true)
  expect((await confirm()).status).toBe('confirmed')
  expect(records.get(changeId)?.state).toBe('undone')
  expect(hostWrites).toBe(3)
  vi.unstubAllGlobals()
})
