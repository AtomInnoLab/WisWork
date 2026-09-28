import { expect, it, vi } from 'vitest'
import JSZip from 'jszip'
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

vi.mock('../src/skills/powerpoint/powerpoint-package', () => ({
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
  const batchSourceRecord: PresentationExistingBatch = {
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
  const meta = {
    backupId: '',
    documentId: 'doc',
    hostSlideId: 'old',
    slideIds: ['old'],
    sha256: '',
    sizeBytes: 0,
    receivedBytes: 0,
    status: 'uploading',
  }
  let data = new Uint8Array()
  let failCurrentBackup = false
  let onSourceRead: (() => void) | undefined
  const request = vi.fn(async (body: unknown) => {
    const input = body as Record<string, unknown>
    const op = input.operation
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
      data = new Uint8Array(meta.sizeBytes)
    } else if (op === 'existing_page_backup_chunk') {
      const part = binary(input.base64 as string)
      data.set(part, input.offset as number)
      meta.receivedBytes += part.length
    } else if (op === 'existing_page_backup_finish') meta.status = 'ready'
    else if (op === 'existing_page_backup_release') {
      data = new Uint8Array()
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
    inspect: vi.fn(async (): Promise<PresentationPageReplacementInspection> => ({
      status:
        slideIds.length === 2
          ? 'staged'
          : slideIds[0] === 'old'
            ? 'baseline'
            : slideIds[0] === 'new'
              ? 'applied'
              : 'undone',
      slideIds: [...slideIds],
    })),
    stage: vi.fn(
      async (_record: unknown, _base64: string, onInserted: (id: string) => Promise<void>) => {
        slideIds = ['old', 'new']
        await onInserted('new')
      },
    ),
    commit: vi.fn(async () => {
      slideIds = ['new']
    }),
    discard: vi.fn(async () => {
      slideIds = ['old']
    }),
    undo: vi.fn(
      async (_record: unknown, _base64: string, onRestored: (id: string) => Promise<void>) => {
        slideIds = ['new', 'restored']
        await onRestored('restored')
        slideIds = ['restored']
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
  const skill = createPresentationExistingPageEditingSkill({
    baseline,
    adapter,
    inspectPage,
    exportAdapter: {
      exportPresentationPagePackage: async () => ({
        slideId: 'old',
        slideIds: ['old'],
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
    preparedFiles,
    digest,
    data: () => data,
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
