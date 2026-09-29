import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationService } from '../../shell/src/main/presentation-service'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  describePagePackageBackup,
  saveChartPackageBackup,
} from '../src/skills/powerpoint/presentation-chart-backup'
import { officeOperationsForSlideIR } from '../src/skills/powerpoint/presentation-office-ir'
import { createPresentationNativeAddExecution } from '../src/skills/powerpoint/presentation-native-add-execution'
import type { PresentationNativeAddBatch } from '../src/skills/powerpoint/presentation-existing-batch'
const roots: string[] = []
afterEach(() => {
  for (const p of roots.splice(0)) rmSync(p, { recursive: true, force: true })
})
async function fixture(options: { fullOrder?: string[]; originalObjectCount?: number } = {}) {
  const fullOrder = options.fullOrder ?? ['host']
  const dir = mkdtempSync(join(tmpdir(), 'native-add-step-'))
  roots.push(dir)
  let service = createPresentationService({ userDataPath: dir })
  const request = async (body: unknown, signal?: AbortSignal) =>
    new Response(
      Buffer.from(await service(body, signal ?? new AbortController().signal)).toString('utf8'),
    )
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const additions = [
    {
      kind: 'text' as const,
      id: 'new-text',
      x: 1,
      y: 3,
      w: 4,
      h: 0.5,
      text: 'Exact text',
      fontSize: 20,
    },
    {
      kind: 'shape' as const,
      id: 'new-shape',
      x: 6,
      y: 3,
      w: 2,
      h: 1,
      shape: 'rect' as const,
      fill: '123456',
      lineColor: '654321',
    },
  ]
  const originals = Array.from({ length: options.originalObjectCount ?? 0 }, (_, i) => ({
    kind: 'shape' as const,
    id: `original-${i}`,
    x: 0,
    y: 0,
    w: 0.1,
    h: 0.1,
    shape: 'rect' as const,
    fill: 'FFFFFF',
    lineColor: 'FFFFFF',
  }))
  deck.slides[0]!.elements.push(...originals, ...additions)
  const { bytes } = await compilePresentationDeck(deck),
    zip = await JSZip.loadAsync(bytes)
  const xml = (await zip.file('ppt/slides/slide1.xml')!.async('string')).replace(
    /(<p:cNvPr[^>]*name="new-text"[^]*?<p:cNvSpPr)\/>/,
    '$1 txBox="1"/>',
  )
  const nodes = [...xml.matchAll(/<p:sp\b[^]*?<\/p:sp>/g)].map((m) => m[0]),
    added = additions.map((e) => nodes.find((n) => n.includes(`name="${e.id}"`))!)
  const baseline = added.reduce((s, n) => s.replace(n, ''), xml)
  const packages: string[] = []
  for (let i = 0; i <= 2; i++) {
    zip.file(
      'ppt/slides/slide1.xml',
      baseline.replace('</p:spTree>', added.slice(0, i).join('') + '</p:spTree>'),
    )
    packages.push(await zip.generateAsync({ type: 'base64' }))
  }
  const operations = officeOperationsForSlideIR(
    { ...deck.slides[0]!, elements: additions, claimIds: [] },
    deck.style,
    0,
  )
  const values = new Map<string, string>(),
    save = vi.fn(async () => {})
  const settings = {
    get: (k: string) => values.get(k),
    set: (k: string, v: string) => {
      values.set(k, v)
    },
    save,
    location: () => 'synthetic',
  }
  const binding = () => createPresentationDocumentBinding(settings, () => 'doc')
  const doc = await binding().documentId(),
    metadata = await describePagePackageBackup(packages[0]!)
  const backup = await saveChartPackageBackup({
    request,
    documentId: doc,
    hostSlideId: 'host',
    slideIds: fullOrder,
    base64: packages[0]!,
    backupId: 'backup',
  })
  const r: PresentationNativeAddBatch = {
    version: 2,
    kind: 'native_page_add',
    changeId: 'add',
    documentId: doc,
    baselineId: 'baseline',
    baselineDigest: metadata.packageDigest,
    hostSlideId: 'host',
    slideIndex: 0,
    beforeSlideIds: fullOrder,
    scope: { slideIds: ['host'] },
    intent: 'add',
    preserved: ['originals'],
    validation: ['readback'],
    risk: 'high',
    backups: [{ ...backup, hostSlideId: 'host', packageDigest: metadata.packageDigest }],
    operations,
    nextIndex: 0,
    createdShapeIds: [],
    state: 'applying',
  }
  await binding().writeExistingBatch(r, undefined)
  let count = 0,
    failAfterWrite = false,
    SDKMismatch = false,
    changeSecondExport = false,
    exports = 0
  const shape = (index: number) => ({
    id: `sdk-real-${index}`,
    name: operations[index]!.name,
    type: index === 0 ? 'TextBox' : 'GeometricShape',
    left: operations[index]!.left,
    top: operations[index]!.top,
    width: operations[index]!.width,
    height: operations[index]!.height,
  })
  const adapter = {
    verifySlides: vi.fn(async () => ({
      slideWidth: 960,
      slideHeight: 540,
      slides: [
        {
          slideId: 'host',
          slideIndex: 0,
          shapes: [],
          shapesTruncated: false,
          overflows: [],
          overlaps: [],
          overlapsTruncated: false,
        },
      ],
    })),
    exportPresentationPagePackage: vi.fn(async () => ({
      slideId: 'host',
      slideIds: fullOrder,
      base64: packages[changeSecondExport && ++exports % 2 === 0 ? Math.min(count + 1, 2) : count]!,
    })),
    listSlideShapes: vi.fn(async () => ({
      slideId: 'host',
      slideIndex: 0,
      shapes: [
        ...originals.map((e, i) => ({
          id: `sdk-original-${i}`,
          name: e.id,
          type: 'GeometricShape',
          left: 0,
          top: 0,
          width: 7.2,
          height: 7.2,
        })),
        ...Array.from({ length: count }, (_, i) => ({
          ...shape(i),
          ...(SDKMismatch ? { name: 'wrong' } : {}),
        })),
      ],
    })),
    executeDeclarative: vi.fn(async () => {
      expect(binding().readExistingBatch('add')).toHaveProperty('inFlightIndex', count)
      count++
      if (failAfterWrite) throw Error('host receipt lost')
      return { createdShapeIds: [`sdk-real-${count - 1}`] }
    }),
  }
  const dependencies: Parameters<typeof createPresentationNativeAddExecution>[0] = {
    documentId: () => binding().documentId(),
    readExistingBatch: (id) => binding().readExistingBatch(id),
    writeExistingBatch: (next, expected) => binding().writeExistingBatch(next, expected),
    request,
    adapter,
  }
  const controller = () => createPresentationNativeAddExecution(dependencies)
  return {
    controller,
    request,
    dependencies,
    binding,
    save,
    adapter,
    r,
    setCount: (n: number) => {
      count = n
    },
    loseHostAck: () => {
      failAfterWrite = true
    },
    allowHostAck: () => {
      failAfterWrite = false
    },
    wrongSDK: () => {
      SDKMismatch = true
    },
    changeExport: () => {
      changeSecondExport = true
    },
    restartPC: () => {
      service = createPresentationService({ userDataPath: dir })
    },
  }
}

import { createPresentationNativeAddRestoration } from '../src/skills/powerpoint/presentation-native-add-restoration'
import type { PresentationExistingPageChange } from '../src/skills/powerpoint/presentation-existing-page'
async function restorationFixture() {
  const f = await fixture(),
    baseline = await f.adapter.exportPresentationPagePackage()
  const sourceBackup = await saveChartPackageBackup({
    request: f.request,
    documentId: f.r.documentId,
    hostSlideId: 'host',
    slideIds: ['host'],
    backupId: 'restoration-source',
    base64: baseline.base64,
  })
  let page: PresentationExistingPageChange = {
    version: 1,
    changeId: 'restore',
    documentId: f.r.documentId,
    baselineId: 'restore-baseline',
    baselineDigest: f.r.baselineDigest,
    scope: { slideIds: ['host'] },
    oldSlideId: 'host',
    beforeSlideIds: ['host'],
    originalPackageDigest: f.r.baselineDigest,
    replacementPackageDigest: f.r.baselineDigest,
    sourceSlideId: '256#',
    restores: {
      sourceKind: 'batch',
      sourceChangeId: 'add',
      sourceHostSlideId: 'host',
      originalBackupId: f.r.backups[0]!.backupId,
      originalPackageDigest: f.r.baselineDigest,
    },
    sourceBackup,
    backup: { backupId: 'current-page', sha256: '1'.repeat(64), sizeBytes: 1 },
    state: 'applied',
    newSlideId: 'restored',
  }
  const exportPage = vi.fn(async () => ({
    slideId: 'restored',
    slideIds: ['restored'],
    base64: baseline.base64,
  }))
  const finalizer = () =>
    createPresentationNativeAddRestoration({
      ...f.dependencies,
      readExistingPageChange: () => page,
      exportPresentationPagePackage: exportPage,
    })
  return {
    ...f,
    finalizer,
    exportPage,
    page,
    setPage: (p: PresentationExistingPageChange) => {
      page = p
    },
  }
}
it('closes only verified applied restoration and retains its actual identity across retry/reopen', async () => {
  const f = await restorationFixture()
  const next = await f.finalizer().finalize('add', 'restore')
  expect(next).toMatchObject({
    state: 'undone',
    restoredSlideId: 'restored',
    nextIndex: 0,
    createdShapeIds: [],
  })
  expect(f.exportPage).toHaveBeenCalledTimes(4)
  f.restartPC()
  expect(await f.finalizer().finalize('add', 'restore')).toEqual(next)
  expect(f.exportPage).toHaveBeenCalledTimes(4)
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it.each(['staged', 'pending', 'commit_pending'] as const)(
  'does not close %s restoration',
  async (state) => {
    const f = await restorationFixture()
    f.setPage({
      ...f.page,
      state,
      ...(state === 'pending' ? { newSlideId: undefined } : {}),
    })
    await expect(f.finalizer().finalize('add', 'restore')).rejects.toThrow()
    expect(f.binding().readExistingBatch('add')!.state).toBe('applying')
    expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  },
)
it('rejects wrong source identity and live order', async () => {
  const f = await restorationFixture()
  f.setPage({
    ...f.page,
    restores: { ...f.page.restores!, sourceChangeId: 'other' },
  })
  await expect(f.finalizer().finalize('add', 'restore')).rejects.toThrow()
  f.setPage(f.page)
  f.exportPage.mockResolvedValue({
    ...(await f.exportPage()),
    slideIds: ['other'],
  })
  await expect(f.finalizer().finalize('add', 'restore')).rejects.toThrow()
  expect(f.binding().readExistingBatch('add')!.state).toBe('applying')
})
it('retains undoing after terminal save failure and completes by metadata only', async () => {
  const f = await restorationFixture()
  f.save.mockResolvedValueOnce().mockRejectedValueOnce(Error('save failed'))
  await expect(f.finalizer().finalize('add', 'restore')).rejects.toThrow('save failed')
  expect(f.binding().readExistingBatch('add')!.state).toBe('undoing')
  expect((await f.finalizer().finalize('add', 'restore')).state).toBe('undone')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('checks live content again after undoing and cancellation before writes', async () => {
  const f = await restorationFixture()
  f.exportPage.mockImplementation(async () => ({
    slideId: 'restored',
    slideIds: f.binding().readExistingBatch('add')!.state === 'undoing' ? ['other'] : ['restored'],
    base64: (await f.adapter.exportPresentationPagePackage()).base64,
  }))
  await expect(f.finalizer().finalize('add', 'restore')).rejects.toThrow()
  expect(f.binding().readExistingBatch('add')!.state).toBe('undoing')
  const signal = new AbortController()
  signal.abort()
  await expect(f.finalizer().finalize('add', 'restore', signal.signal)).rejects.toThrow('cancelled')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('rejects a changed live package and missing backup before metadata writes', async () => {
  const f = await restorationFixture()
  f.setCount(1)
  f.exportPage.mockImplementation(async () => ({
    ...(await f.adapter.exportPresentationPagePackage()),
    slideId: 'restored',
    slideIds: ['restored'],
  }))
  await expect(f.finalizer().finalize('add', 'restore')).rejects.toThrow(
    'presentation_native_add_conflict',
  )
  expect(f.binding().readExistingBatch('add')!.state).toBe('applying')
  const missing = createPresentationNativeAddRestoration({
    ...f.dependencies,
    readExistingPageChange: () => f.page,
    exportPresentationPagePackage: f.exportPage,
    request: async () => new Response('{}', { status: 404 }),
  })
  await expect(missing.finalize('add', 'restore')).rejects.toThrow()
  expect(f.binding().readExistingBatch('add')!.state).toBe('applying')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('reconciles a durable terminal write whose acknowledgment was lost without rewriting it', async () => {
  const f = await restorationFixture()
  const finalizer = createPresentationNativeAddRestoration({
    ...f.dependencies,
    readExistingPageChange: () => f.page,
    exportPresentationPagePackage: f.exportPage,
    writeExistingBatch: async (next, expected) => {
      await f.dependencies.writeExistingBatch(next, expected)
      if (next.state === 'undone') throw Error('lost ACK')
    },
  })
  await expect(finalizer.finalize('add', 'restore')).rejects.toThrow('lost ACK')
  expect(f.binding().readExistingBatch('add')!.state).toBe('undone')
  const saveCount = f.save.mock.calls.length
  expect((await f.finalizer().finalize('add', 'restore')).state).toBe('undone')
  expect(f.save).toHaveBeenCalledTimes(saveCount)
  f.setPage({ ...f.page, newSlideId: 'other-restored' })
  await expect(f.finalizer().finalize('add', 'restore')).rejects.toThrow()
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('preserves the recorded SDK prefix when closing a partially completed addition', async () => {
  const f = await restorationFixture()
  await f.controller().step('add')
  const result = await f.finalizer().finalize('add', 'restore')
  expect(result).toMatchObject({
    state: 'undone',
    nextIndex: 1,
    createdShapeIds: ['sdk-real-0'],
    restoredSlideId: 'restored',
  })
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(1)
})

import { createPresentationNativeModifyRestoration } from '../src/skills/powerpoint/presentation-native-modify-restoration'
import type { PresentationNativeModifyBatch } from '../src/skills/powerpoint/presentation-existing-batch'
async function genericRestorationFixture() {
  const f = await restorationFixture()
  const values = new Map<string, string>()
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save: async () => {},
    location: () => 'synthetic',
  }
  const binding = () => createPresentationDocumentBinding(settings, () => 'doc')
  const original = f.r
  const record: PresentationNativeModifyBatch = {
    version: 3,
    kind: 'native_page_modify',
    changeId: 'modify',
    documentId: original.documentId,
    baselineId: original.baselineId,
    baselineDigest: original.baselineDigest,
    beforeSlideIds: ['host'],
    scope: { slideIds: ['host'] },
    intent: 'Delete a native object',
    preserved: [],
    validation: [],
    risk: 'high',
    backups: original.backups,
    operations: [{ op: 'delete_shape', slide_index: 0, shape_id: 'native-shape' }],
    pages: [{ hostSlideId: 'host', slideIndex: 0, expectedPackageDigest: original.baselineDigest }],
    nextIndex: 0,
    state: 'applying',
  }
  await binding().writeExistingBatch(record, undefined)
  await binding().writeExistingBatch({ ...record, inFlightIndex: 0 }, record)
  f.setPage({ ...f.page, restores: { ...f.page.restores!, sourceChangeId: 'modify' } })
  const finalizer = () =>
    createPresentationNativeModifyRestoration({
      ...f.dependencies,
      readExistingBatch: (id) => binding().readExistingBatch(id),
      writeExistingBatch: (next, expected) => binding().writeExistingBatch(next, expected),
      readExistingPageChange: () => ({
        ...f.page,
        restores: { ...f.page.restores!, sourceChangeId: 'modify' },
      }),
      listExistingPageChanges: () => [
        { ...f.page, restores: { ...f.page.restores!, sourceChangeId: 'modify' } },
      ],
      exportPresentationPagePackage: f.exportPage,
    })
  return { ...f, binding, finalizer }
}
it('closes a generic uncertain deletion only after verified package restoration and survives reopen', async () => {
  const f = await genericRestorationFixture()
  const restored = await f.finalizer().finalize('modify', ['restore'])
  expect(restored).toMatchObject({
    version: 3,
    state: 'undone',
    restoredSlideIds: { host: 'restored' },
  })
  expect(restored.inFlightIndex).toBeUndefined()
  f.restartPC()
  expect(f.binding().readExistingBatch('modify')).toMatchObject({
    state: 'undone',
    restoredSlideIds: { host: 'restored' },
  })
  expect(await f.finalizer().finalize('modify', ['restore'])).toEqual(restored)
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('rejects missing, duplicated or changed generic restore evidence without closing the savepoint', async () => {
  const f = await genericRestorationFixture()
  await expect(f.finalizer().finalize('modify', [])).rejects.toThrow()
  await expect(f.finalizer().finalize('modify', ['restore', 'restore'])).rejects.toThrow()
  f.exportPage.mockResolvedValueOnce({
    slideId: 'restored',
    slideIds: ['other'],
    base64: (await f.adapter.exportPresentationPagePackage()).base64,
  })
  await expect(f.finalizer().finalize('modify', ['restore'])).rejects.toThrow()
  expect(f.binding().readExistingBatch('modify')).toMatchObject({
    state: 'applying',
    inFlightIndex: 0,
  })
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})

it('allows an exact original-page recovery record while its generic source remains uncertain', async () => {
  const f = await genericRestorationFixture()
  const { newSlideId: _inserted, ...identity } = f.page
  const page: PresentationExistingPageChange = {
    ...identity,
    changeId: 'generic-stage',
    state: 'pending',
    restores: { ...f.page.restores!, sourceChangeId: 'modify' },
  }
  await f.binding().writeExistingPageChange(page, undefined)
  expect(f.binding().readExistingPageChange('generic-stage')).toMatchObject({
    state: 'pending',
    restores: { sourceChangeId: 'modify' },
  })
  expect(f.binding().readExistingBatch('modify')).toMatchObject({
    state: 'applying',
    inFlightIndex: 0,
  })
})
it('does not let unrelated page work bypass an uncertain generic savepoint', async () => {
  const f = await genericRestorationFixture()
  const { newSlideId: _inserted, ...identity } = f.page
  const page: PresentationExistingPageChange = {
    ...identity,
    changeId: 'unrelated-stage',
    state: 'pending',
    restores: { ...f.page.restores!, sourceChangeId: 'other-modify' },
  }
  await expect(f.binding().writeExistingPageChange(page, undefined)).rejects.toThrow(
    'presentation_change_history_pending',
  )
  expect(f.binding().readExistingPageChange('unrelated-stage')).toBeUndefined()
})

async function multiGenericRestorationFixture() {
  const f = await fixture({ fullOrder: ['host', 'other'] })
  const original = (await f.adapter.exportPresentationPagePackage()).base64
  const otherBackup = await saveChartPackageBackup({
    request: f.request,
    documentId: f.r.documentId,
    hostSlideId: 'other',
    slideIds: ['host', 'other'],
    base64: original,
    backupId: 'original-other',
  })
  const values = new Map<string, string>()
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save: async () => {},
    location: () => 'synthetic',
  }
  const binding = () => createPresentationDocumentBinding(settings, () => 'doc')
  const record: PresentationNativeModifyBatch = {
    version: 3,
    kind: 'native_page_modify',
    changeId: 'multi-modify',
    documentId: f.r.documentId,
    baselineId: 'baseline',
    baselineDigest: f.r.baselineDigest,
    beforeSlideIds: ['host', 'other'],
    scope: { slideIds: ['host', 'other'] },
    intent: 'Two uncertain native changes',
    preserved: [],
    validation: [],
    risk: 'high',
    backups: [
      f.r.backups[0]!,
      { ...otherBackup, hostSlideId: 'other', packageDigest: f.r.baselineDigest },
    ],
    operations: [
      { op: 'delete_shape', slide_index: 0, shape_id: 'shape' },
      { op: 'delete_shape', slide_index: 1, shape_id: 'shape' },
    ],
    pages: [
      { hostSlideId: 'host', slideIndex: 0, expectedPackageDigest: f.r.baselineDigest },
      { hostSlideId: 'other', slideIndex: 1, expectedPackageDigest: f.r.baselineDigest },
    ],
    nextIndex: 0,
    state: 'applying',
  }
  await binding().writeExistingBatch(record, undefined)
  await binding().writeExistingBatch({ ...record, inFlightIndex: 0 }, record)
  const pages: PresentationExistingPageChange[] = []
  for (const hostSlideId of ['host', 'other']) {
    const order = hostSlideId === 'host' ? ['host', 'other'] : ['restored-host', 'other']
    const sourceBackup = await saveChartPackageBackup({
      request: f.request,
      documentId: f.r.documentId,
      hostSlideId,
      slideIds: order,
      base64: original,
      backupId: `source-${hostSlideId}`,
    })
    pages.push({
      version: 1,
      changeId: `restore-${hostSlideId}`,
      documentId: f.r.documentId,
      baselineId: 'restore-baseline',
      baselineDigest: f.r.baselineDigest,
      scope: { slideIds: [hostSlideId] },
      oldSlideId: hostSlideId,
      beforeSlideIds: order,
      originalPackageDigest: f.r.baselineDigest,
      replacementPackageDigest: f.r.baselineDigest,
      sourceSlideId: '256#',
      restores: {
        sourceKind: 'batch',
        sourceChangeId: 'multi-modify',
        sourceHostSlideId: hostSlideId,
        originalBackupId: record.backups.find((backup) => backup.hostSlideId === hostSlideId)!
          .backupId,
        originalPackageDigest: f.r.baselineDigest,
      },
      sourceBackup,
      backup: { backupId: `edited-${hostSlideId}`, sha256: '1'.repeat(64), sizeBytes: 1 },
      state: 'applied',
      newSlideId: `restored-${hostSlideId}`,
    })
  }
  const exportPage = vi.fn(async (slideId: string) => ({
    slideId,
    slideIds: ['restored-host', 'restored-other'],
    base64: original,
  }))
  const controller = () =>
    createPresentationNativeModifyRestoration({
      documentId: () => binding().documentId(),
      readExistingBatch: (id) => binding().readExistingBatch(id),
      writeExistingBatch: (next, expected) => binding().writeExistingBatch(next, expected),
      readExistingPageChange: (id) => pages.find((page) => page.changeId === id),
      listExistingPageChanges: () => pages,
      request: f.request,
      exportPresentationPagePackage: exportPage,
    })
  return { ...f, binding, pages, controller, exportPage }
}
it('verifies every restored generic page in its sequential deck order and retries regardless of supplied ID order', async () => {
  const f = await multiGenericRestorationFixture()
  const restored = await f.controller().finalize('multi-modify', ['restore-other', 'restore-host'])
  expect(restored).toMatchObject({
    state: 'undone',
    restoredSlideIds: { host: 'restored-host', other: 'restored-other' },
  })
  expect(await f.controller().finalize('multi-modify', ['restore-host', 'restore-other'])).toEqual(
    restored,
  )
  expect(f.exportPage.mock.calls.some(([id]) => id === 'restored-host')).toBe(true)
  expect(f.exportPage.mock.calls.some(([id]) => id === 'restored-other')).toBe(true)
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('rejects a generic restore chain that skipped a prior page identity or has an additional pending restore', async () => {
  const f = await multiGenericRestorationFixture()
  f.pages[1] = { ...f.pages[1]!, beforeSlideIds: ['host', 'other'] }
  await expect(
    f.controller().finalize('multi-modify', ['restore-host', 'restore-other']),
  ).rejects.toThrow('presentation_native_modify_restore_conflict')
  expect(f.binding().readExistingBatch('multi-modify')).toMatchObject({
    state: 'applying',
    inFlightIndex: 0,
  })
  f.pages[1] = { ...f.pages[1]!, beforeSlideIds: ['restored-host', 'other'] }
  f.pages.push({
    ...f.pages[0]!,
    changeId: 'extra-pending',
    state: 'pending',
    newSlideId: undefined,
  })
  await expect(
    f.controller().finalize('multi-modify', ['restore-host', 'restore-other']),
  ).rejects.toThrow('presentation_native_modify_restore_conflict')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
