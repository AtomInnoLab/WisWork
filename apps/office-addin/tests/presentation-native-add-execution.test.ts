import { afterEach, expect, it, vi } from 'vitest'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill'
import type { PowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
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
it('executes one native operation at a time with durable inFlight and SDK identities across restart', async () => {
  const f = await fixture()
  expect((await f.controller().inspect('add')).observation.status).toBe('none')
  const first = await f.controller().step('add')
  expect(first.createdShapeIds).toEqual(['sdk-real-0'])
  expect(first.nextIndex).toBe(1)
  f.restartPC()
  const end = await f.controller().step('add')
  expect(end.state).toBe('applied')
  expect(end.createdShapeIds).toEqual(['sdk-real-0', 'sdk-real-1'])
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(2)
})
it('claims lost host ACK by strict package and actual SDK readback without replay', async () => {
  const f = await fixture()
  f.loseHostAck()
  await expect(f.controller().step('add')).rejects.toThrow()
  expect(f.binding().readExistingBatch('add')).toHaveProperty('inFlightIndex', 0)
  f.restartPC()
  const claimed = await f.controller().step('add')
  expect(claimed.nextIndex).toBe(1)
  expect(claimed.createdShapeIds).toEqual(['sdk-real-0'])
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(1)
})
it('keeps an unwritten inFlight pending and rejects automatic replay', async () => {
  const f = await fixture()
  await f.binding().writeExistingBatch({ ...f.r, inFlightIndex: 0 }, f.r)
  expect((await f.controller().inspect('add')).observation.status).toBe('none')
  await expect(f.controller().step('add')).rejects.toThrow('presentation_native_add_pending')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('fails before host write when durable inFlight save fails', async () => {
  const f = await fixture()
  f.save.mockRejectedValueOnce(Error('save_failed'))
  await expect(f.controller().step('add')).rejects.toThrow('save_failed')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  expect(f.binding().readExistingBatch('add')).toEqual(f.r)
})
it('rejects SDK mismatch and package changes between SDK read and second export without adopting IDs', async () => {
  const f = await fixture()
  f.setCount(1)
  await f.binding().writeExistingBatch({ ...f.r, inFlightIndex: 0 }, f.r)
  f.wrongSDK()
  await expect(f.controller().inspect('add')).rejects.toThrow('presentation_native_add_conflict')
  expect(f.binding().readExistingBatch('add')).toHaveProperty('nextIndex', 0)
  const g = await fixture()
  g.changeExport()
  await expect(g.controller().inspect('add')).rejects.toThrow('presentation_native_add_conflict')
  expect(g.adapter.executeDeclarative).not.toHaveBeenCalled()
})

it('recovers a saved host result after final receipt save rollback without executing again', async () => {
  const f = await fixture()
  f.save.mockResolvedValueOnce(undefined).mockRejectedValueOnce(Error('terminal_save_failed'))
  await expect(f.controller().step('add')).rejects.toThrow('terminal_save_failed')
  expect(f.binding().readExistingBatch('add')).toHaveProperty('inFlightIndex', 0)
  const result = await f.controller().step('add')
  expect(result.createdShapeIds).toEqual(['sdk-real-0'])
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(1)
})

it('uses full exported order and complete target SDK shapes without diagnostic truncation restrictions', async () => {
  const f = await fixture({
    fullOrder: ['host', ...Array.from({ length: 24 }, (_, i) => `other-${i}`)],
    originalObjectCount: 110,
  })
  f.adapter.verifySlides.mockImplementation(async () => ({
    slideWidth: 960,
    slideHeight: 540,
    truncated: true,
    slides: [],
  }))
  const proof = await f.controller().inspect('add')
  expect(proof.observation.status).toBe('none')
  expect(f.adapter.verifySlides).not.toHaveBeenCalled()
  expect((await f.controller().step('add')).createdShapeIds).toEqual(['sdk-real-0'])
  expect((await f.adapter.listSlideShapes()).shapes.length).toBeGreaterThan(100)
})
it('rejects another valid record returned for the requested change ID before any host write', async () => {
  const f = await fixture()
  const controller = createPresentationNativeAddExecution({
    ...f.dependencies,
    readExistingBatch: () => ({ ...f.r, changeId: 'other' }),
  })
  await expect(controller.inspect('add')).rejects.toThrow('presentation_existing_batch_missing')
  await expect(controller.step('add')).rejects.toThrow('presentation_existing_batch_missing')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('keeps durable inFlight when cancellation arrives after one actual native write', async () => {
  const f = await fixture(),
    abort = new AbortController()
  const execute = f.adapter.executeDeclarative.getMockImplementation()!
  f.adapter.executeDeclarative.mockImplementationOnce(async () => {
    const result = await execute()
    abort.abort()
    return result
  })
  await expect(f.controller().step('add', abort.signal)).rejects.toThrow('cancelled')
  expect(f.binding().readExistingBatch('add')).toHaveProperty('inFlightIndex', 0)
  const result = await f.controller().step('add')
  expect(result.createdShapeIds).toEqual(['sdk-real-0'])
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(1)
})
it('rejects a concurrent persisted record change during SDK read without writing progress', async () => {
  const f = await fixture()
  f.adapter.listSlideShapes.mockImplementationOnce(async () => {
    await f.binding().writeExistingBatch({ ...f.r, inFlightIndex: 0 }, f.r)
    return { slideId: 'host', slideIndex: 0, shapes: [] }
  })
  await expect(f.controller().inspect('add')).rejects.toThrow('presentation_existing_batch_stale')
  expect(f.binding().readExistingBatch('add')).toHaveProperty('nextIndex', 0)
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})

it('reopens the real skill to inspect a lost host ACK and confirm only the remaining native write', async () => {
  const f = await fixture()
  const firstProposals = createStructuredProposalController()
  const firstSkill = createPowerPointSkill({
    adapter: f.adapter as unknown as PowerPointAdapter,
    proposals: firstProposals,
    nativeAddSavepoint: {
      documentId: f.dependencies.documentId,
      request: f.dependencies.request,
      readExistingBatch: (id) => f.binding().readExistingBatch(id),
      writeExistingBatch: (next, expected) => f.binding().writeExistingBatch(next, expected),
    },
  })
  await firstSkill.executeTool({
    id: 'initial',
    name: 'resume_slide_ir_addition',
    input: { change_id: 'add' },
  })
  const firstPending = firstProposals.pending()!
  const firstDecision = firstProposals.waitForDecision(firstPending.id)
  f.loseHostAck()
  await expect(firstProposals.confirm(firstPending.id)).rejects.toThrow('host receipt lost')
  expect((await firstDecision).status).toBe('failed')
  expect(f.binding().readExistingBatch('add')).toMatchObject({ nextIndex: 0, inFlightIndex: 0 })
  f.restartPC()
  f.allowHostAck()
  const proposals = createStructuredProposalController()
  const skill = createPowerPointSkill({
    adapter: f.adapter as unknown as PowerPointAdapter,
    proposals,
    nativeAddSavepoint: {
      documentId: f.dependencies.documentId,
      request: f.dependencies.request,
      readExistingBatch: (id) => f.binding().readExistingBatch(id),
      writeExistingBatch: (next, expected) => f.binding().writeExistingBatch(next, expected),
    },
  })
  const inspect = await skill.executeTool({
    id: 'inspect',
    name: 'inspect_slide_ir_addition',
    input: { change_id: 'add' },
  })
  expect(inspect.isError).not.toBe(true)
  expect(inspect.mutated).toBe(false)
  expect(JSON.parse(inspect.output)).toMatchObject({
    nextIndex: 0,
    createdShapeIds: ['sdk-real-0'],
    observation: { completedCount: 1 },
  })
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(1)
  expect(f.binding().readExistingBatch('add')).toHaveProperty('inFlightIndex', 0)
  const resume = await skill.executeTool({
    id: 'resume',
    name: 'resume_slide_ir_addition',
    input: { change_id: 'add' },
  })
  expect(resume.isError).not.toBe(true)
  expect(resume.mutated).toBe(false)
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(1)
  const pending = proposals.pending()!
  expect(pending.preview).toMatchObject({ receiptRecovery: true, completed: 1, total: 2 })
  const decision = proposals.waitForDecision(pending.id)
  await proposals.confirm(pending.id)
  expect(await decision).toEqual({ status: 'confirmed' })
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(2)
  expect(f.binding().readExistingBatch('add')).toMatchObject({
    state: 'applied',
    nextIndex: 2,
    createdShapeIds: ['sdk-real-0', 'sdk-real-1'],
  })
  expect(f.binding().readExistingBatch('add')).not.toHaveProperty('inFlightIndex')
})

it.each(['name', 'type', 'geometry'] as const)(
  'refuses SDK %s mismatch while the real exported native package is unchanged',
  async (field) => {
    const f = await fixture()
    f.loseHostAck()
    await expect(f.controller().step('add')).rejects.toThrow('host receipt lost')
    const original = f.adapter.listSlideShapes.getMockImplementation()!
    f.adapter.listSlideShapes.mockImplementation(async () => {
      const result = await original()
      result.shapes[0] = {
        ...result.shapes[0]!,
        ...(field === 'name'
          ? { name: 'another-object' }
          : field === 'type'
            ? { type: 'Picture' }
            : { left: result.shapes[0]!.left + 1 }),
      }
      return result
    })
    await expect(f.controller().inspect('add')).rejects.toThrow('presentation_native_add_conflict')
    await expect(f.controller().step('add')).rejects.toThrow('presentation_native_add_conflict')
    expect(f.binding().readExistingBatch('add')).toMatchObject({
      nextIndex: 0,
      inFlightIndex: 0,
      createdShapeIds: [],
    })
    expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(1)
  },
)
