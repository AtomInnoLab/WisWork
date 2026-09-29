import { afterEach, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { PNG } from 'pngjs'
import { createPresentationChangesController } from '../src/agent/presentation-changes.js'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill'
import type { PowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
import { validExistingBatchTransition } from '../src/skills/powerpoint/presentation-existing-batch'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { createPresentationExistingPageEditingSkill } from '../src/skills/powerpoint/presentation-existing-page-editing'
import {
  validExistingPageTransition,
  type PresentationExistingPageChange,
} from '../src/skills/powerpoint/presentation-existing-page'
import type { PresentationBaselineSkill } from '../src/skills/powerpoint/presentation-baseline'
import type { PresentationPageReplacementInspection } from '../src/skills/powerpoint/browser-presentation-page-replacement-adapter'
import type { InMemoryVfs } from '../src/skills/shared/vfs'
import type { PresentationExistingChange } from '../src/skills/powerpoint/presentation-existing-change'
import type { PresentationExistingBatch } from '../src/skills/powerpoint/presentation-existing-batch'

vi.mock('../src/skills/powerpoint/powerpoint-package', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/skills/powerpoint/powerpoint-package')>()),
  MAX_PPTX_PACKAGE_BYTES: 8 * 1024 * 1024,
  presentationPackageDigest: async (base64: string) => {
    const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
    return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (x) =>
      x.toString(16).padStart(2, '0'),
    ).join('')
  },
}))
const binary = (base64: string) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
const base64 = (value: Uint8Array) =>
  btoa(Array.from(value, (x) => String.fromCharCode(x)).join(''))
afterEach(() => vi.unstubAllGlobals())

it('keeps a frozen P0-19 three-object edit in one saved page transaction through undo', async () => {
  const material = new URL(
    '../../../docs/product/ppt-benchmark-materials/PPT-P0-19/',
    import.meta.url,
  )
  const zip = await JSZip.loadAsync(
    readFileSync(new URL('wiswork-image-dense-research-draft.pptx', material)),
  )
  for (let page = 1; page <= 8; page++)
    if (page !== 4) {
      zip.remove(`ppt/slides/slide${page}.xml`)
      zip.remove(`ppt/slides/_rels/slide${page}.xml.rels`)
    }
  zip.file(
    'ppt/presentation.xml',
    (await zip.file('ppt/presentation.xml')!.async('string')).replace(
      /<p:sldId\b[^>]*\/>/g,
      (item) => (item.includes('r:id="rId5"') ? item : ''),
    ),
  )
  const original = await zip.generateAsync({ type: 'base64' })
  const sourceXml = await zip.file('ppt/slides/slide4.xml')!.async('string')
  const caption = [...sourceXml.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)].find(([xml]) =>
    xml.includes('id="6"'),
  )![0]
  const off = /<a:off x="(\d+)" y="(\d+)"\/>/.exec(caption)!
  const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(caption)!
  const before = {
    left: Number(off[1]) / 12700,
    top: Number(off[2]) / 12700,
    width: Number(ext[1]) / 12700,
    height: Number(ext[2]) / 12700,
  }
  const f = await fixture()
  f.changeBackup(binary(original))
  f.setCurrentPage(binary(original))
  f.preparedFiles.set(
    '/home/user/schematic-12.png',
    readFileSync(new URL('images/schematic-12.png', material)),
  )
  vi.stubGlobal(
    'createImageBitmap',
    vi.fn(async (blob: Blob) => {
      const decoded = PNG.sync.read(Buffer.from(await blob.arrayBuffer()))
      return { width: decoded.width, height: decoded.height, close: vi.fn() }
    }),
  )
  const prepared = await f.skill.executeTool({
    id: 'prepare-p0-19',
    name: 'prepare_existing_presentation_composite_revision',
    input: {
      baseline_id: 'baseline',
      slide_id: 'old',
      text: { shape_id: '3', start: 0, before: '参与者与证据路径', after: '参与者与证据链' },
      geometry: { shape_id: '6', before, after: { ...before, top: before.top + 5.76 } },
      picture: { shape_id: '7', path: '/home/user/schematic-12.png' },
    },
  })
  expect(prepared.isError, prepared.output).not.toBe(true)
  const revision = JSON.parse(prepared.output)
  const stage = await f.call('stage', revision.nextInput)
  expect(stage.isError, stage.output).not.toBe(true)
  expect((await f.confirm()).status).toBe('confirmed')
  const changeId = [...f.records.keys()][0]!
  expect(f.records.get(changeId)).toMatchObject({
    state: 'staged',
    pictureTarget: { shapeId: '7' },
    originalPackageDigest: revision.beforeDigest,
    replacementPackageDigest: revision.afterDigest,
  })
  const commit = await f.call('commit', { change_id: changeId })
  expect(commit.isError, commit.output).not.toBe(true)
  expect((await f.confirm()).status).toBe('confirmed')
  expect(f.records.get(changeId)?.state).toBe('applied')
  const undo = await f.call('undo', { change_id: changeId })
  expect(undo.isError, undo.output).not.toBe(true)
  expect((await f.confirm()).status).toBe('confirmed')
  expect(f.records.get(changeId)?.state).toBe('undone')
  expect(f.data()).toEqual(binary(original))
})
async function fixture() {
  const make = async (text: string) => {
    const zip = new JSZip()
    zip.file(
      'ppt/presentation.xml',
      '<p:presentation><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>',
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>',
    )
    zip.file('ppt/slides/slide1.xml', text)
    return zip.generateAsync({ type: 'uint8array' })
  }
  let source = await make('new'),
    backupBytes = await make('old')
  let currentPageBytes = backupBytes
  const preparedFiles = new Map<string, Uint8Array>()
  const baselineSnapshot = {
    baselineId: 'baseline',
    documentId: 'doc',
    contentDigest: 'a'.repeat(64),
    scope: { kind: 'current', slideIds: ['old'] },
    context: { slideIds: ['old'], selectedSlideIds: ['old'], selectedShapeIds: [] },
    pages: [{ slideId: 'old', shapes: [{ id: 'title', type: 'TextBox' }] }],
  }
  const baseline = {
    snapshot: () => structuredClone(baselineSnapshot),
    executeTool: async () => ({ output: JSON.stringify({ unchanged: true }), mutated: false }),
  } as unknown as PresentationBaselineSkill
  const records = new Map<string, PresentationExistingPageChange>()
  let failWrite = false
  let failAfterWrite = false
  let slideIds = ['old']
  const digest = async (value: Uint8Array) =>
    Array.from(
      new Uint8Array(await crypto.subtle.digest('SHA-256', Uint8Array.from(value).buffer)),
      (x) => x.toString(16).padStart(2, '0'),
    ).join('')
  const sourceDigest = await digest(backupBytes)
  const originalSourceRecord: PresentationExistingChange = {
    version: 1,
    changeId: 'single',
    documentId: 'doc',
    baselineId: 'baseline',
    baselineDigest: 'a'.repeat(64),
    scope: { slideIds: ['old'], shapeIds: ['title'] },
    hostSlideId: 'old',
    shapeId: 'title',
    shapeType: 'TextBox',
    kind: 'text',
    before: 'original',
    after: 'modified',
    state: 'applied',
    beforeSlideIds: ['old'],
    backup: {
      hostSlideId: 'old',
      backupId: 'source-backup',
      sha256: sourceDigest,
      packageDigest: sourceDigest,
      sizeBytes: backupBytes.length,
    },
  }
  let batchSourceRecord: PresentationExistingBatch = {
    version: 1,
    changeId: 'batch',
    documentId: 'doc',
    baselineId: 'baseline',
    baselineDigest: 'a'.repeat(64),
    scope: { slideIds: ['old'], shapeIds: ['title', 'subtitle'] },
    intent: 'Edit two objects',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      {
        kind: 'text',
        hostSlideId: 'old',
        shapeId: 'title',
        shapeType: 'TextBox',
        before: 'a',
        after: 'b',
      },
      {
        kind: 'text',
        hostSlideId: 'old',
        shapeId: 'subtitle',
        shapeType: 'TextBox',
        before: 'c',
        after: 'd',
      },
    ],
    state: 'applied',
    cursor: 2,
    beforeSlideIds: ['old'],
    backups: [
      {
        hostSlideId: 'old',
        backupId: 'source-backup',
        sha256: sourceDigest,
        packageDigest: sourceDigest,
        sizeBytes: backupBytes.length,
      },
    ],
  }
  const defaultMeta = {
    backupId: '',
    documentId: 'doc',
    hostSlideId: 'old',
    slideIds: ['old'],
    sha256: '',
    sizeBytes: 0,
    receivedBytes: 0,
    status: 'uploading',
  }
  const backupStore = new Map<string, { meta: typeof defaultMeta; data: Uint8Array }>()
  let failCurrentBackup = false
  let onSourceRead: (() => void) | undefined
  const request = vi.fn(async (body: unknown) => {
    const input = body as Record<string, unknown>
    const op = input.operation
    const entry = backupStore.get(String(input.backupId)) ?? {
      meta: { ...defaultMeta },
      data: new Uint8Array(),
    }
    backupStore.set(String(input.backupId), entry)
    const meta = entry.meta
    const data = entry.data
    if (input.backupId === 'source-backup') {
      if (op === 'existing_page_backup_status') {
        onSourceRead?.()
        return new Response(
          JSON.stringify({
            backupId: 'source-backup',
            documentId: 'doc',
            hostSlideId: 'old',
            slideIds: ['old'],
            sha256: sourceDigest,
            sizeBytes: backupBytes.length,
            receivedBytes: backupBytes.length,
            status: 'ready',
          }),
        )
      }
      if (op === 'existing_page_backup_read')
        return new Response(
          JSON.stringify({
            backupId: 'source-backup',
            offset: input.offset,
            sizeBytes: backupBytes.length,
            sha256: sourceDigest,
            base64: base64(
              backupBytes.subarray(
                input.offset as number,
                (input.offset as number) + (input.length as number),
              ),
            ),
          }),
        )
    }
    if (op === 'existing_page_backup_begin' && failCurrentBackup) throw new Error('quota_exceeded')
    if (op === 'existing_page_backup_begin') {
      Object.assign(meta, input, { receivedBytes: 0, status: 'uploading' })
      entry.data = new Uint8Array(meta.sizeBytes)
    } else if (op === 'existing_page_backup_chunk') {
      const part = binary(input.base64 as string)
      data.set(part, input.offset as number)
      meta.receivedBytes += part.length
    } else if (op === 'existing_page_backup_finish') meta.status = 'ready'
    else if (op === 'existing_page_backup_release') {
      entry.data = new Uint8Array()
      return new Response(JSON.stringify({ ...input, status: 'released' }))
    } else if (op === 'existing_page_backup_read')
      return new Response(
        JSON.stringify({
          backupId: meta.backupId,
          offset: input.offset,
          sizeBytes: meta.sizeBytes,
          sha256: meta.sha256,
          base64: base64(
            data.subarray(
              input.offset as number,
              (input.offset as number) + (input.length as number),
            ),
          ),
        }),
      )
    return new Response(JSON.stringify(meta))
  })
  const adapter = {
    captureUnchangedPageDigests: vi.fn(async () => []),
    reconcilePending: vi.fn(async () =>
      slideIds.length === 2
        ? ({ status: 'inserted', newSlideId: 'new' } as const)
        : ({ status: 'baseline' } as const),
    ),
    inspect: vi.fn(
      async (record: {
        oldSlideId: string
        newSlideId?: string
        restoredSlideId?: string
      }): Promise<PresentationPageReplacementInspection> => ({
        status:
          slideIds.length === 2
            ? 'staged'
            : slideIds[0] === record.oldSlideId
              ? 'baseline'
              : slideIds[0] === record.newSlideId
                ? 'applied'
                : 'undone',
        slideIds: [...slideIds],
      }),
    ),
    stage: vi.fn(
      async (
        record: { oldSlideId: string; beforeSlideIds: string[] },
        _base64: string,
        onInserted: (id: string) => Promise<void>,
      ) => {
        const inserted = record.oldSlideId === 'old' ? 'new' : `${record.oldSlideId}-new`
        slideIds = record.beforeSlideIds.flatMap((id) =>
          id === record.oldSlideId ? [id, inserted] : [id],
        )
        await onInserted(inserted)
      },
    ),
    commit: vi.fn(async (record: { oldSlideId: string; newSlideId?: string }) => {
      slideIds = slideIds.filter((id) => id !== record.oldSlideId)
    }),
    discard: vi.fn(async () => {
      slideIds = ['old']
    }),
    undo: vi.fn(
      async (
        record: { oldSlideId: string; newSlideId?: string },
        _base64: string,
        onRestored: (id: string) => Promise<void>,
      ) => {
        const restored = record.oldSlideId === 'old' ? 'restored' : `${record.oldSlideId}-restored`
        slideIds = [record.newSlideId!, restored]
        await onRestored(restored)
        slideIds = [restored]
      },
    ),
  }
  const proposals = createStructuredProposalController()
  const inspectPage = vi.fn(async (slideId: string) => ({
    slideId,
    shapesTruncated: false,
    screenshot: {
      mime: 'image/png' as const,
      base64:
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII=',
    },
  }))
  const create = () =>
    createPresentationExistingPageEditingSkill({
      baseline,
      adapter,
      inspectPage,
      exportAdapter: {
        exportPresentationPagePackage: async (slideId: string) => ({
          slideId,
          slideIds: [...slideIds],
          base64: base64(currentPageBytes),
        }),
      },
      vfs: {
        readBytes: (path: string) => preparedFiles.get(path) ?? source,
        writeFile: (path: string, value: Uint8Array) => {
          preparedFiles.set(path, value)
        },
      } as unknown as InMemoryVfs,
      request,
      proposals,
      documentId: async () => 'doc',
      available: () => true,
      readExistingChange: () => originalSourceRecord,
      readExistingBatch: () => batchSourceRecord,
      readExistingPageChange: (id) => records.get(id),
      writeExistingPageChange: async (record, expected) => {
        if (failWrite) throw new Error('settings_save_failed')
        expect(records.get(record.changeId)).toEqual(expected)
        expect(validExistingPageTransition(expected, record)).toBe(true)
        records.set(record.changeId, structuredClone(record))
        if (failAfterWrite) throw new Error('journal_ack_lost')
      },
    })
  let skill = create()
  const call = (action: string, input: Record<string, unknown>) =>
    skill.executeTool({ id: 'tool', name: `${action}_existing_presentation_page_change`, input })
  const confirm = async () => {
    const proposalId = proposals.pending()!.id
    const decision = proposals.waitForDecision(proposalId)
    await proposals.confirm(proposalId)
    return decision
  }
  return {
    skill,
    baselineSnapshot: () => structuredClone(baselineSnapshot),
    reopen: () => {
      skill.clear()
      skill = create()
    },
    setOrder: (ids: string[]) => {
      slideIds = ids
    },
    call,
    confirm,
    records,
    adapter,
    inspectPage,
    request,
    setCurrentBackupUnavailable: (value: boolean) => {
      failCurrentBackup = value
    },
    onSourceBackupRead: (callback: () => void) => {
      onSourceRead = callback
    },
    setWriteFailure: (value: boolean) => {
      failWrite = value
    },
    setAckFailure: () => {
      failAfterWrite = true
    },
    removeStaged: () => {
      slideIds = ['old']
    },
    markUnknownInserted: () => {
      slideIds = ['old', 'new']
    },
    source: () => source,
    backup: () => backupBytes,
    changeSource: (value: Uint8Array) => {
      source = value
    },
    changeBackup: (value: Uint8Array) => {
      backupBytes = value
    },
    setCurrentPage: (value: Uint8Array) => {
      currentPageBytes = value
    },
    originalSourceRecord,
    batchSourceRecord,
    readBatchSourceRecord: () => structuredClone(batchSourceRecord),
    exportPage: async (slideId: string) => ({
      slideId,
      slideIds: [...slideIds],
      base64: base64(currentPageBytes),
    }),
    setBatchSourceRecord: (record: PresentationExistingBatch) => {
      batchSourceRecord = structuredClone(record)
    },
    preparedFiles,
    digest,
    backupStore,
    data: () =>
      [...backupStore.values()].find((entry) => entry.meta.backupId)?.data ?? new Uint8Array(),
  }
}

it('captures exact durable page IDs after stage, commit and undo', async () => {
  const f = await fixture()
  await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  const staged = await f.confirm()
  if (staged.status !== 'confirmed') throw new Error('not confirmed')
  expect(staged.postWrite).toMatchObject({
    status: 'captured',
    pages: [{ slideId: 'old' }, { slideId: 'new' }],
  })
  const changeId = [...f.records.keys()][0]!
  expect(f.records.get(changeId)?.captures?.map((capture) => capture.hostSlideId)).toEqual([
    'old',
    'new',
  ])
  await f.call('commit', { change_id: changeId })
  const committed = await f.confirm()
  if (committed.status !== 'confirmed') throw new Error('not confirmed')
  expect(committed.postWrite).toMatchObject({ status: 'captured', pages: [{ slideId: 'new' }] })
  expect(f.records.get(changeId)?.captures?.map((capture) => capture.hostSlideId)).toEqual(['new'])
  await f.call('undo', { change_id: changeId })
  const undone = await f.confirm()
  if (undone.status !== 'confirmed') throw new Error('not confirmed')
  expect(undone.postWrite).toMatchObject({ status: 'captured', pages: [{ slideId: 'restored' }] })
  expect(f.records.get(changeId)?.captures?.map((capture) => capture.hostSlideId)).toEqual([
    'restored',
  ])
  expect(f.inspectPage.mock.calls.map(([id]) => id)).toEqual(['old', 'new', 'new', 'restored'])
})

it('prepares an exact original-page backup as a verified staged restore source', async () => {
  const f = await fixture()
  f.setCurrentPage(f.source())
  const prepared = await f.skill.executeTool({
    id: 'restore-source',
    name: 'prepare_existing_presentation_original_page_restore',
    input: { source_kind: 'single', change_id: 'single', slide_id: 'old' },
  })
  expect(prepared.isError, prepared.output).not.toBe(true)
  expect(JSON.parse(prepared.output)).toMatchObject({
    sourceKind: 'single',
    sourceChangeId: 'single',
    slideId: 'old',
    nextTool: 'stage_existing_presentation_page_change',
  })
  const next = JSON.parse(prepared.output).nextInput as Record<string, unknown>
  expect(f.preparedFiles.get(next.path as string)).toEqual(f.backup())
  expect(f.adapter.stage).not.toHaveBeenCalled()
  const staged = await f.call('stage', { baseline_id: 'baseline', ...next })
  expect(staged.isError, staged.output).not.toBe(true)
  await f.confirm()
  const record = [...f.records.values()][0]!
  expect(record.restores).toMatchObject({
    sourceKind: 'single',
    sourceChangeId: 'single',
    sourceHostSlideId: 'old',
    originalBackupId: 'source-backup',
  })
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
  expect(f.data()).toEqual(f.source())
  const committed = await f.call('commit', { change_id: record.changeId })
  expect(committed.isError, committed.output).not.toBe(true)
  await f.confirm()
  expect(f.records.get(record.changeId)?.state).toBe('applied')
})

it('refuses a released or changed original backup before staging a restore', async () => {
  const f = await fixture()
  const prepared = await f.skill.executeTool({
    id: 'restore-source',
    name: 'prepare_existing_presentation_original_page_restore',
    input: { source_kind: 'single', change_id: 'single', slide_id: 'old' },
  })
  expect(prepared.isError).not.toBe(true)
  const next = JSON.parse(prepared.output).nextInput as Record<string, unknown>
  f.originalSourceRecord.backupReleasedAt = '2026-09-28T00:00:00.000Z'
  const blocked = await f.call('stage', { baseline_id: 'baseline', ...next })
  expect(blocked.isError).toBe(true)
  expect(f.adapter.stage).not.toHaveBeenCalled()
  expect(f.records.size).toBe(0)
})
it('freezes source metadata before PC read so an in-place release cannot pass as unchanged', async () => {
  const f = await fixture()
  f.onSourceBackupRead(() => {
    f.originalSourceRecord.backupReleasedAt = '2026-09-28T00:00:00.000Z'
  })
  const prepared = await f.skill.executeTool({
    id: 'restore-source',
    name: 'prepare_existing_presentation_original_page_restore',
    input: { source_kind: 'single', change_id: 'single', slide_id: 'old' },
  })
  expect(prepared.isError).toBe(true)
  expect(f.preparedFiles.size).toBe(0)
})
it('blocks restore staging if its source changes after proposal or its current-page backup fails', async () => {
  const sourceChanged = await fixture()
  const prepared = await sourceChanged.skill.executeTool({
    id: 'restore-source',
    name: 'prepare_existing_presentation_original_page_restore',
    input: { source_kind: 'single', change_id: 'single', slide_id: 'old' },
  })
  const next = JSON.parse(prepared.output).nextInput as Record<string, unknown>
  const proposed = await sourceChanged.call('stage', { baseline_id: 'baseline', ...next })
  expect(proposed.isError).not.toBe(true)
  sourceChanged.originalSourceRecord.backupReleasedAt = '2026-09-28T00:00:00.000Z'
  await expect(sourceChanged.confirm()).rejects.toThrow()
  expect(sourceChanged.adapter.stage).not.toHaveBeenCalled()
  const quota = await fixture()
  const other = await quota.skill.executeTool({
    id: 'restore-source',
    name: 'prepare_existing_presentation_original_page_restore',
    input: { source_kind: 'single', change_id: 'single', slide_id: 'old' },
  })
  const nextInput = JSON.parse(other.output).nextInput as Record<string, unknown>
  const restore = await quota.call('stage', { baseline_id: 'baseline', ...nextInput })
  expect(restore.isError).not.toBe(true)
  quota.setCurrentBackupUnavailable(true)
  await expect(quota.confirm()).rejects.toThrow()
  expect(quota.adapter.stage).not.toHaveBeenCalled()
  expect(quota.records.size).toBe(0)
})
it('prepares one affected page of a completed batch from its exact PC backup', async () => {
  const f = await fixture()
  const prepared = await f.skill.executeTool({
    id: 'restore-batch',
    name: 'prepare_existing_presentation_original_page_restore',
    input: { source_kind: 'batch', change_id: 'batch', slide_id: 'old' },
  })
  expect(prepared.isError, prepared.output).not.toBe(true)
  expect(JSON.parse(prepared.output)).toMatchObject({
    sourceKind: 'batch',
    sourceChangeId: 'batch',
    slideId: 'old',
  })
  f.batchSourceRecord.state = 'undoing'
  const blocked = await f.skill.executeTool({
    id: 'restore-batch',
    name: 'prepare_existing_presentation_original_page_restore',
    input: { source_kind: 'batch', change_id: 'batch', slide_id: 'old' },
  })
  expect(blocked.isError).toBe(true)
})

it('records staged visual judgments for both pages and clears them before commit', async () => {
  const f = await fixture()
  const proposed = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  await f.confirm()
  const changeId = JSON.parse(proposed.output).changeId as string
  for (const slideId of ['old', 'new']) {
    const captured = await f.call('capture', { change_id: changeId, slide_id: slideId })
    expect(captured.isError).toBeUndefined()
    const screenshotDigest = JSON.parse(captured.output).screenshotDigest as string
    const reviewed = await f.call('record', {
      change_id: changeId,
      slide_id: slideId,
      screenshot_digest: screenshotDigest,
      status: 'pass',
      notes: 'checked',
    })
    expect(reviewed.isError).toBeUndefined()
  }
  expect(f.records.get(changeId)?.reviews?.map((review) => review.hostSlideId)).toEqual([
    'old',
    'new',
  ])
  await f.call('commit', { change_id: changeId })
  await f.confirm()
  expect(f.records.get(changeId)?.reviews).toBeUndefined()
})

it('compares both staged pages with persisted screenshot receipts', async () => {
  const f = await fixture()
  const proposed = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  await f.confirm()
  const changeId = JSON.parse(proposed.output).changeId as string
  const matched = await f.call('inspect', { change_id: changeId })
  expect(JSON.parse(matched.output).visualReceipts).toEqual([
    { hostSlideId: 'old', status: 'matched' },
    { hostSlideId: 'new', status: 'matched' },
  ])
  f.inspectPage.mockImplementation(async (slideId) => ({
    slideId,
    shapesTruncated: false,
    screenshot: {
      mime: 'image/png',
      base64:
        slideId === 'new'
          ? 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='
          : 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII=',
    },
  }))
  const different = await f.call('inspect', { change_id: changeId })
  expect(JSON.parse(different.output).visualReceipts).toEqual([
    { hostSlideId: 'old', status: 'matched' },
    { hostSlideId: 'new', status: 'different' },
  ])
})

it('does not claim page evidence if the staged page disappears during capture', async () => {
  const f = await fixture()
  const proposed = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  f.inspectPage.mockImplementationOnce(async (slideId) => {
    f.removeStaged()
    return {
      slideId,
      shapesTruncated: false,
      screenshot: {
        mime: 'image/png',
        base64:
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII=',
      },
    }
  })
  expect(await f.confirm()).toMatchObject({
    status: 'confirmed',
    postWrite: { status: 'unavailable' },
  })
  expect(f.records.get(JSON.parse(proposed.output).changeId)?.state).toBe('staged')
  expect(f.records.get(JSON.parse(proposed.output).changeId)?.captures).toBeUndefined()
})

it('rejects an image-only replacement when the original page has native content', async () => {
  const f = await fixture()
  const zip = await JSZip.loadAsync(f.source())
  zip.file('ppt/slides/slide1.xml', '<p:sld><p:cSld><p:spTree><p:pic/></p:spTree></p:cSld></p:sld>')
  f.changeSource(await zip.generateAsync({ type: 'uint8array' }))
  const result = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/flattened.pptx',
  })
  expect(result.isError).toBe(true)
  expect(result.output).toContain('presentation_page_source_rasterized')
  expect(f.adapter.stage).not.toHaveBeenCalled()
  expect(f.request).not.toHaveBeenCalled()
})

it('stages with durable backup, then separately commits and restores after reopen', async () => {
  const f = await fixture()
  const proposed = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  expect(proposed.isError).toBeUndefined()
  expect(f.request).not.toHaveBeenCalled()
  await f.confirm()
  const id = JSON.parse(proposed.output).changeId as string
  expect(f.records.get(id)?.state).toBe('staged')
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
  await f.call('commit', { change_id: id })
  await f.confirm()
  expect(f.records.get(id)?.state).toBe('applied')
  expect((await f.call('release', { change_id: id })).isError).toBe(true)
  await f.call('undo', { change_id: id })
  await f.confirm()
  expect(f.records.get(id)?.state).toBe('undone')
  await f.call('release', { change_id: id })
  await f.confirm()
  expect(f.records.get(id)?.backupReleasedAt).toMatch(/^\d{4}-/)
})

it('stops before host insertion if source drifts during backup', async () => {
  const f = await fixture()
  await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  f.request.mockImplementationOnce(async (body: unknown) => {
    f.changeSource(Uint8Array.from([1, 2, 3]))
    const input = body as Record<string, unknown>
    return new Response(JSON.stringify({ ...input, receivedBytes: 0, status: 'uploading' }))
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.adapter.stage).not.toHaveBeenCalled()
})

it('marks unknown pending insertion for manual review and never replays it', async () => {
  const f = await fixture()
  const proposed = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  f.adapter.stage.mockImplementationOnce(async () => {
    throw new Error('office_state_uncertain')
  })
  await expect(f.confirm()).rejects.toThrow()
  const id = JSON.parse(proposed.output).changeId as string
  expect(f.records.get(id)?.state).toBe('pending')
  const inspected = await f.call('inspect', { change_id: id })
  expect(JSON.parse(inspected.output).manualReview).toBe(true)
  expect(JSON.parse(inspected.output).nextTool).toBe(
    'reconcile_pending_existing_presentation_page_change',
  )
  expect((await f.call('resume', { change_id: id })).isError).toBe(true)
  const reconciled = await f.call('reconcile', { change_id: id })
  expect(JSON.parse(reconciled.output)).toMatchObject({
    status: 'pending_no_insert_observed',
    hostWrite: false,
  })
  expect(f.records.get(id)?.state).toBe('pending')
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
})

it('reconciles a pending insertion into durable staged state without replaying the host write', async () => {
  const f = await fixture()
  const proposed = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  f.adapter.stage.mockImplementationOnce(async () => {
    f.markUnknownInserted()
    throw new Error('office_state_uncertain')
  })
  await expect(f.confirm()).rejects.toThrow('office_state_uncertain')
  const id = JSON.parse(proposed.output).changeId as string
  expect(f.records.get(id)?.state).toBe('pending')
  f.setWriteFailure(true)
  expect((await f.call('reconcile', { change_id: id })).isError).toBe(true)
  expect(f.records.get(id)?.state).toBe('pending')
  f.setWriteFailure(false)
  const reconciled = await f.call('reconcile', { change_id: id })
  expect(reconciled.isError, reconciled.output).not.toBe(true)
  expect(JSON.parse(reconciled.output)).toMatchObject({
    status: 'staged',
    newSlideId: 'new',
    hostWrite: false,
  })
  expect(f.records.get(id)).toMatchObject({ state: 'staged', newSlideId: 'new' })
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
  expect(f.adapter.commit).not.toHaveBeenCalled()
  expect((await f.call('reconcile', { change_id: id })).isError).toBe(true)
})

it('discards a staged page without deleting the original', async () => {
  const f = await fixture()
  const proposed = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  await f.confirm()
  const id = JSON.parse(proposed.output).changeId as string
  expect((await f.call('release', { change_id: id })).isError).toBe(true)
  await f.call('discard', { change_id: id })
  await f.confirm()
  expect(f.records.get(id)?.state).toBe('discarded')
  expect(f.adapter.discard).toHaveBeenCalledTimes(1)
  expect(f.adapter.commit).not.toHaveBeenCalled()
  await f.call('release', { change_id: id })
  await f.confirm()
  expect(f.records.get(id)?.backupReleasedAt).toMatch(/^\d{4}-/)
  expect(f.data().length).toBe(0)
  expect((await f.call('release', { change_id: id })).isError).toBe(true)
})

it('rejects a PC backup response with the wrong document before saving intent or inserting', async () => {
  const f = await fixture()
  const proposed = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  f.request.mockImplementationOnce(async (body: unknown) => {
    const input = body as Record<string, unknown>
    return new Response(
      JSON.stringify({
        ...input,
        documentId: 'other',
        status: 'ready',
        receivedBytes: input.sizeBytes,
      }),
    )
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.records.has(JSON.parse(proposed.output).changeId)).toBe(false)
  expect(f.adapter.stage).not.toHaveBeenCalled()
})

it('does not insert when the durable pending savepoint fails', async () => {
  const f = await fixture()
  await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  f.setWriteFailure(true)
  await expect(f.confirm()).rejects.toThrow()
  expect(f.adapter.stage).not.toHaveBeenCalled()
  expect(
    f.request.mock.calls.some(
      ([body]) => (body as Record<string, unknown>).operation === 'existing_page_backup_release',
    ),
  ).toBe(true)
  expect(f.data().length).toBe(0)
})

it('keeps the backup if the first savepoint persisted but its acknowledgement was lost', async () => {
  const f = await fixture()
  const proposed = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  f.setAckFailure()
  await expect(f.confirm()).rejects.toThrow('journal_ack_lost')
  expect(f.records.has(JSON.parse(proposed.output).changeId)).toBe(true)
  expect(
    f.request.mock.calls.some(
      ([body]) => (body as Record<string, unknown>).operation === 'existing_page_backup_release',
    ),
  ).toBe(false)
  expect(f.data().length).toBeGreaterThan(0)
  expect(f.adapter.stage).not.toHaveBeenCalled()
})

it('resumes a known inserted page without running insertion again', async () => {
  const f = await fixture()
  const proposed = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  const original = f.adapter.stage.getMockImplementation()!
  f.adapter.stage.mockImplementationOnce(async (...args) => {
    await original(...args)
    throw new Error('office_state_uncertain')
  })
  await expect(f.confirm()).rejects.toThrow()
  const id = JSON.parse(proposed.output).changeId as string
  expect(f.records.get(id)?.state).toBe('inserted')
  await f.call('resume', { change_id: id })
  await f.confirm()
  expect(f.records.get(id)?.state).toBe('staged')
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
})
it('does not report a staged page as verified after that page disappears', async () => {
  const f = await fixture()
  const proposed = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/home/user/rebuilt.pptx',
  })
  await f.confirm()
  f.removeStaged()
  const inspected = await f.call('inspect', { change_id: JSON.parse(proposed.output).changeId })
  expect(JSON.parse(inspected.output)).toMatchObject({
    state: 'staged',
    inspection: { status: 'baseline' },
    currentHostVerified: false,
    manualReview: true,
  })
})

it('retains a durable replacement source and reapplies to a fresh staged journal', async () => {
  const f = await fixture()
  const proposed = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/rebuilt.pptx',
  })
  await f.confirm()
  const oldId = JSON.parse(proposed.output).changeId
  expect(f.records.get(oldId)?.sourceBackup).toBeDefined()
})

async function undonePageFixture() {
  const f = await fixture()
  const proposed = await f.call('stage', {
    baseline_id: 'baseline',
    slide_id: 'old',
    path: '/rebuilt.pptx',
  })
  await f.confirm()
  const oldId = JSON.parse(proposed.output).changeId as string
  await f.call('commit', { change_id: oldId })
  await f.confirm()
  await f.call('undo', { change_id: oldId })
  await f.confirm()
  f.reopen()
  return { f, oldId }
}

it('reopens and reapplies using independent original/source backups and a separate commit', async () => {
  const { f, oldId } = await undonePageFixture()
  const oldRecord = structuredClone(f.records.get(oldId)!)
  f.changeSource(new Uint8Array())
  const proposal = await f.call('reapply', { change_id: oldId })
  expect(proposal.isError, proposal.output).not.toBe(true)
  const newId = JSON.parse(proposal.output).changeId
  expect(newId).not.toBe(oldId)
  expect(JSON.parse(proposal.output).reapplies).toBe(oldId)
  await f.confirm()
  const newRecord = f.records.get(newId)!
  expect(newRecord).toMatchObject({
    state: 'staged',
    oldSlideId: oldRecord.restoredSlideId,
    reapplies: oldId,
  })
  expect(newRecord.backup.backupId).not.toBe(oldRecord.backup.backupId)
  expect(newRecord.sourceBackup!.backupId).not.toBe(oldRecord.sourceBackup!.backupId)
  expect(f.records.get(oldId)).toEqual(oldRecord)
  await f.call('release', { change_id: oldId })
  await f.confirm()
  expect(f.backupStore.get(newRecord.sourceBackup!.backupId)!.data.length).toBeGreaterThan(0)
  await f.call('commit', { change_id: newId })
  await f.confirm()
  await f.call('undo', { change_id: newId })
  await f.confirm()
  expect(f.records.get(newId)?.state).toBe('undone')
  await f.call('reapply', { change_id: newId })
  await f.confirm()
})

it.each(['backup', 'sourceBackup'] as const)(
  'rejects missing/tampered/released %s without rebuilding old backups',
  async (field) => {
    const { f, oldId } = await undonePageFixture()
    const record = f.records.get(oldId)!
    const backup = f.backupStore.get(record[field]!.backupId)!
    const begins = f.request.mock.calls.filter(
      ([body]) => (body as Record<string, unknown>).operation === 'existing_page_backup_begin',
    ).length
    backup.data[0] ^= 1
    expect((await f.call('reapply', { change_id: oldId })).isError).toBe(true)
    backup.meta.status = 'released'
    expect(await f.call('reapply', { change_id: oldId })).toMatchObject({
      isError: true,
      output: 'presentation_page_backup_invalid',
    })
    f.backupStore.delete(record[field]!.backupId)
    expect((await f.call('reapply', { change_id: oldId })).isError).toBe(true)
    expect(
      f.request.mock.calls.filter(
        ([body]) => (body as Record<string, unknown>).operation === 'existing_page_backup_begin',
      ),
    ).toHaveLength(begins)
  },
)

it('rejects legacy and released page records and current package/order conflicts', async () => {
  const { f, oldId } = await undonePageFixture()
  const record = structuredClone(f.records.get(oldId)!)
  f.records.set(oldId, { ...record, sourceBackup: undefined })
  expect((await f.call('reapply', { change_id: oldId })).isError).toBe(true)
  f.records.set(oldId, { ...record, backupReleasedAt: new Date().toISOString() })
  expect((await f.call('reapply', { change_id: oldId })).isError).toBe(true)
  f.records.set(oldId, record)
  f.setOrder(['extra', record.restoredSlideId!])
  expect(await f.call('reapply', { change_id: oldId })).toMatchObject({
    isError: true,
    output: 'presentation_existing_page_conflict',
  })
  f.setOrder([record.restoredSlideId!])
  f.setCurrentPage(f.source())
  expect(await f.call('reapply', { change_id: oldId })).toMatchObject({
    isError: true,
    output: 'presentation_existing_page_conflict',
  })
})

it.each(['history', 'package', 'order', 'backup', 'sourceBackup'])(
  'rejects stale %s after proposal before copying backups or insertion',
  async (field) => {
    const { f, oldId } = await undonePageFixture()
    const record = f.records.get(oldId)!
    const proposed = await f.call('reapply', { change_id: oldId })
    if (field === 'history') f.records.set(oldId, { ...record, captures: undefined })
    else if (field === 'package') f.setCurrentPage(f.source())
    else if (field === 'order') f.setOrder(['extra', record.restoredSlideId!])
    else f.backupStore.get(record[field as 'backup' | 'sourceBackup']!.backupId)!.data[0] ^= 1
    await expect(f.confirm()).rejects.toThrow('proposal_stale')
    expect(f.records.has(JSON.parse(proposed.output).changeId)).toBe(false)
    expect(f.adapter.stage).toHaveBeenCalledTimes(1)
  },
)

it('keeps a lost reapply insertion receipt pending without replaying the host write', async () => {
  const { f, oldId } = await undonePageFixture()
  const old = structuredClone(f.records.get(oldId)!)
  const proposed = await f.call('reapply', { change_id: oldId })
  f.adapter.stage.mockImplementationOnce(async () => {
    f.setOrder(['restored', 'unknown'])
    throw new Error('office_state_uncertain')
  })
  await expect(f.confirm()).rejects.toThrow('office_state_uncertain')
  const newId = JSON.parse(proposed.output).changeId
  expect(f.records.get(newId)).toMatchObject({ state: 'pending', reapplies: oldId })
  expect(f.records.get(oldId)).toEqual(old)
  f.reopen()
  expect((await f.call('resume', { change_id: newId })).isError).toBe(true)
  expect(f.adapter.stage).toHaveBeenCalledTimes(2)
})

it('retries release after the original receipt succeeded and the source release failed', async () => {
  const { f, oldId } = await undonePageFixture()
  const record = f.records.get(oldId)!
  const request = f.request.getMockImplementation()!
  let fail = true
  f.request.mockImplementation(async (body) => {
    const input = body as Record<string, unknown>
    if (
      input.operation === 'existing_page_backup_release' &&
      input.backupId === record.sourceBackup!.backupId &&
      fail
    )
      throw new Error('office_state_uncertain')
    return request(body)
  })
  await f.call('release', { change_id: oldId })
  await expect(f.confirm()).rejects.toThrow('office_state_uncertain')
  expect(f.records.get(oldId)?.backupReleasedAt).toBeUndefined()
  expect(f.backupStore.get(record.backup.backupId)!.data.length).toBe(0)
  expect(f.backupStore.get(record.sourceBackup!.backupId)!.data.length).toBeGreaterThan(0)
  fail = false
  await f.call('release', { change_id: oldId })
  await f.confirm()
  expect(f.records.get(oldId)?.backupReleasedAt).toBeDefined()
  expect(f.backupStore.get(record.sourceBackup!.backupId)!.data.length).toBe(0)
})

it('retains actual historical restores source identity when targeting a restored native page', async () => {
  const f = await fixture()
  f.setCurrentPage(f.source())
  const prepared = await f.skill.executeTool({
    id: 'prepare',
    name: 'prepare_existing_presentation_original_page_restore',
    input: { source_kind: 'single', change_id: 'single', slide_id: 'old' },
  })
  const preparedInput = JSON.parse(prepared.output).nextInput
  const proposed = await f.call('stage', { baseline_id: 'baseline', ...preparedInput })
  expect(proposed.isError, proposed.output).not.toBe(true)
  await f.confirm()
  const oldId = JSON.parse(proposed.output).changeId
  await f.call('commit', { change_id: oldId })
  await f.confirm()
  await f.call('undo', { change_id: oldId })
  await f.confirm()
  const originalProvenance = f.records.get(oldId)!.restores
  const reapply = await f.call('reapply', { change_id: oldId })
  expect(reapply.isError, reapply.output).not.toBe(true)
  await f.confirm()
  const newRecord = f.records.get(JSON.parse(reapply.output).changeId)!
  expect(newRecord.oldSlideId).toBe('restored')
  expect(newRecord.restores).toEqual(originalProvenance)
  expect(newRecord.restores?.sourceHostSlideId).toBe('old')
})

it('fails safely on source-backup quota without host insertion or an ownerless ready original', async () => {
  const f = await fixture()
  const request = f.request.getMockImplementation()!
  let begins = 0
  f.request.mockImplementation(async (body) => {
    if (
      (body as Record<string, unknown>).operation === 'existing_page_backup_begin' &&
      ++begins === 2
    )
      throw new Error('quota_exceeded')
    return request(body)
  })
  await f.call('stage', { baseline_id: 'baseline', slide_id: 'old', path: '/rebuilt.pptx' })
  await expect(f.confirm()).rejects.toThrow('quota_exceeded')
  expect(f.adapter.stage).not.toHaveBeenCalled()
  expect(f.records.size).toBe(0)
  expect(f.data().length).toBe(0)
})

it.each([1, 2])(
  'classifies a real PC quota response for page backup %s before host insertion',
  async (blockedBegin) => {
    const f = await fixture()
    const backend = f.request.getMockImplementation()!
    let begins = 0
    f.request.mockImplementation(async (body) => {
      if (
        (body as Record<string, unknown>).operation === 'existing_page_backup_begin' &&
        ++begins === blockedBegin
      )
        return new Response(
          JSON.stringify({ error: 'quota_exceeded', detail: '/private/backup' }),
          { status: 409 },
        )
      return backend(body)
    })
    const staged = await f.call('stage', {
      baseline_id: 'baseline',
      slide_id: 'old',
      path: '/rebuilt.pptx',
    })
    expect(staged.isError, staged.output).not.toBe(true)
    const pending = f.confirm()
    await expect(pending).rejects.toThrow('presentation_existing_backup_capacity')
    expect(f.adapter.stage).not.toHaveBeenCalled()
    expect(f.records.size).toBe(0)
  },
)

it.each([
  ['existing_page_backup_begin', 'quota_exceeded: /private/secret'],
  ['existing_page_backup_begin', 'access_denied'],
  ['existing_page_backup_chunk', 'quota_exceeded'],
])(
  'filters page backup response %s %s without exposing PC body details',
  async (blockedOperation, error) => {
    const f = await fixture()
    const backend = f.request.getMockImplementation()!
    f.request.mockImplementation(async (body) =>
      (body as Record<string, unknown>).operation === blockedOperation
        ? new Response(JSON.stringify({ error, detail: '/private/secret' }), { status: 409 })
        : backend(body),
    )
    await f.call('stage', { baseline_id: 'baseline', slide_id: 'old', path: '/rebuilt.pptx' })
    await expect(f.confirm()).rejects.toThrow('presentation_page_backup_failed')
    expect(f.adapter.stage).not.toHaveBeenCalled()
    expect(f.records.size).toBe(0)
  },
)

it('prepares a pending native-add V2 original page for separately confirmed restoration without replaying additions', async () => {
  const f = await fixture()
  const backup = f.batchSourceRecord.backups![0]!
  const record: import('../src/skills/powerpoint/presentation-existing-batch').PresentationNativeAddBatch =
    {
      version: 2,
      kind: 'native_page_add',
      changeId: 'native-add',
      documentId: 'doc',
      baselineId: 'native-before',
      baselineDigest: backup.packageDigest,
      hostSlideId: 'old',
      slideIndex: 0,
      beforeSlideIds: ['old'],
      scope: { slideIds: ['old'] },
      intent: 'Add native title',
      preserved: ['Original page'],
      validation: ['Native readback'],
      risk: 'high',
      backups: [backup],
      operations: [
        {
          op: 'add_text_box',
          slide_index: 0,
          name: 'added-title',
          text: 'Title',
          left: 72,
          top: 72,
          width: 720,
          height: 72,
        },
      ],
      createdShapeIds: [],
      nextIndex: 0,
      inFlightIndex: 0,
      state: 'applying',
    }
  f.setBatchSourceRecord(record)
  f.setCurrentPage(f.source())
  const controller = createPresentationChangesController({
    available: () => false,
    existingAvailable: () => true,
    nativeRestorationAvailable: () => true,
    nativeRestoreFinalizationAvailable: () => true,
    artifact: () => undefined,
    documentId: async () => 'doc',
    listChangeHistory: () => [
      {
        id: 'existing_batch:native-add',
        kind: 'existing_batch',
        sequence: 1,
        legacy: false,
        record: f.readBatchSourceRecord(),
      },
      ...[...f.records.values()].map((page, index) => ({
        id: `existing_page:${page.changeId}`,
        kind: 'existing_page' as const,
        sequence: index + 2,
        legacy: false,
        record: structuredClone(page),
      })),
    ],
    executeTool: async (call, signal) => {
      if (call.name === 'read_presentation_baseline')
        return {
          output: JSON.stringify({
            ...f.baselineSnapshot(),
            qaPassed: false,
            coverage: { pagePackages: 'read' },
          }),
          mutated: false,
          summary: 'Synthetic fresh baseline',
        }
      if (call.name === 'finalize_slide_ir_addition_restore')
        return recovery!.executeTool(call, signal)
      return f.skill.executeTool(call, signal)
    },
  })
  await controller.refresh()
  await controller.run('existing_batch:native-add', 'undo')
  expect(controller.snapshot().error).toBeUndefined()
  expect(controller.snapshot().notice).toContain('确认')
  expect(f.preparedFiles.size).toBe(1)
  expect([...f.preparedFiles.values()][0]).toEqual(f.backup())
  expect(f.adapter.stage).not.toHaveBeenCalled()
  await f.confirm()
  const restore = [...f.records.values()][0]!
  expect(restore.restores).toMatchObject({
    sourceKind: 'batch',
    sourceChangeId: 'native-add',
    sourceHostSlideId: 'old',
  })
  expect(f.adapter.stage).toHaveBeenCalledOnce()
  const hostWrite = vi.fn()
  const metadataWrite = vi.fn(
    async (next: PresentationExistingBatch, expected: PresentationExistingBatch | undefined) => {
      expect(f.readBatchSourceRecord()).toEqual(expected)
      expect(validExistingBatchTransition(expected, next)).toBe(true)
      f.setBatchSourceRecord(next)
    },
  )
  const recovery = createPowerPointSkill({
    adapter: {
      exportPresentationPagePackage: f.exportPage,
      executeDeclarative: hostWrite,
    } as unknown as PowerPointAdapter,
    proposals: createStructuredProposalController(),
    nativeAddSavepoint: {
      documentId: async () => 'doc',
      request: f.request,
      readExistingBatch: () => f.readBatchSourceRecord(),
      writeExistingBatch: metadataWrite,
      readExistingPageChange: (id) => f.records.get(id),
    },
  })
  const close = () =>
    recovery!.executeTool({
      id: 'finalize-native',
      name: 'finalize_slide_ir_addition_restore',
      input: { change_id: record.changeId, restoration_change_id: restore.changeId },
    })
  expect(await close()).toMatchObject({
    isError: true,
    mutated: false,
    output: 'presentation_native_add_conflict',
  })
  expect(metadataWrite).not.toHaveBeenCalled()
  expect(f.readBatchSourceRecord()).toMatchObject({ state: 'applying', inFlightIndex: 0 })
  await f.call('commit', { change_id: restore.changeId })
  await f.confirm()
  expect(f.records.get(restore.changeId)?.state).toBe('applied')
  // The simulated host now exports the exact original package actually passed to the confirmed stage.
  f.setCurrentPage(binary(f.adapter.stage.mock.calls[0]![1]))
  await controller.refresh()
  expect(
    controller.snapshot().entries.find((entry) => entry.id === 'existing_batch:native-add')
      ?.actions,
  ).toContain('finalize')
  await controller.run('existing_batch:native-add', 'finalize')
  expect(controller.snapshot().error).toBeUndefined()
  expect(
    controller.snapshot().entries.find((entry) => entry.id === 'existing_batch:native-add'),
  ).toMatchObject({ state: 'undone', pageId: restore.newSlideId })
  const finalized = await close()
  expect(finalized.isError, finalized.output).not.toBe(true)
  expect(finalized.mutated).toBe(false)
  expect(JSON.parse(finalized.output)).toMatchObject({
    state: 'undone',
    restoredSlideId: restore.newSlideId,
    historicalOnly: true,
    visualQaVerified: false,
  })
  expect(f.readBatchSourceRecord()).toMatchObject({
    state: 'undone',
    nextIndex: 0,
    createdShapeIds: [],
    restoredSlideId: restore.newSlideId,
  })
  expect(f.readBatchSourceRecord()).not.toHaveProperty('inFlightIndex')
  expect(metadataWrite).toHaveBeenCalledTimes(2)
  await close()
  expect(metadataWrite).toHaveBeenCalledTimes(2)
  expect(hostWrite).not.toHaveBeenCalled()
  expect(f.adapter.stage).toHaveBeenCalledOnce()
  expect(f.adapter.commit).toHaveBeenCalledOnce()
  expect(record).toMatchObject({ state: 'applying', nextIndex: 0, inFlightIndex: 0 })
})
