import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationService } from '../../shell/src/main/presentation-service'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { describePagePackageBackup } from '../src/skills/powerpoint/presentation-chart-backup'
import { officeOperationsForSlideIR } from '../src/skills/powerpoint/presentation-office-ir'
import { createPresentationNativeAddExecution } from '../src/skills/powerpoint/presentation-native-add-execution'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { createPresentationNativeAddProposal } from '../src/skills/powerpoint/presentation-native-add-proposal'
import type {
  PowerPointAdapter,
  VerifySlidesResult,
} from '../src/skills/powerpoint/browser-powerpoint-adapter'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill'
import type {
  PresentationExistingBatch,
  PresentationNativeAddBatch,
} from '../src/skills/powerpoint/presentation-existing-batch'
const roots: string[] = []
afterEach(() => {
  for (const p of roots.splice(0)) rmSync(p, { recursive: true, force: true })
})
async function fixture(withClaims = false) {
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
  deck.slides[0]!.elements.push(...additions)
  const { bytes } = await compilePresentationDeck(deck),
    zip = await JSZip.loadAsync(bytes)
  const xml = (await zip.file('ppt/slides/slide1.xml')!.async('string')).replace(
    /(<p:cNvPr[^>]*name="new-text"[^]*?<p:cNvSpPr)\/>/,
    '$1 txBox="1"/>',
  )
  const nodes = [...xml.matchAll(/<p:sp\b[^]*?<\/p:sp>/g)].map((m) => m[0]),
    added = additions.map((e) => nodes.find((n) => n.includes(`name="${e.id}"`))!)
  if (withClaims)
    // Synthetic Office export: native role and the SDK operation's explicit top alignment.
    added.push(
      nodes
        .find((n) => n.includes('name="source-attribution"'))!
        .replace('<p:cNvSpPr/>', '<p:cNvSpPr txBox="1"/>')
        .replace('anchor="ctr"', 'anchor="t"'),
    )
  const originalFooter = nodes.find((n) => n.includes('name="source-attribution"'))!
  const baseline = added.reduce(
    (s, n) => s.replace(n, ''),
    withClaims ? xml.replace(originalFooter, '') : xml,
  )
  const packages: string[] = []
  for (let i = 0; i <= added.length; i++) {
    zip.file(
      'ppt/slides/slide1.xml',
      baseline.replace('</p:spTree>', added.slice(0, i).join('') + '</p:spTree>'),
    )
    packages.push(await zip.generateAsync({ type: 'base64' }))
  }
  const operations = officeOperationsForSlideIR(
    { ...deck.slides[0]!, elements: additions, claimIds: withClaims ? ['source-1'] : [] },
    deck.style,
    0,
    withClaims ? deck.claims : [],
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
  const backup = { backupId: 'unused', ...metadata }
  const r: PresentationNativeAddBatch = {
    version: 2,
    kind: 'native_page_add',
    changeId: 'add',
    documentId: doc,
    baselineId: 'baseline',
    baselineDigest: metadata.packageDigest,
    hostSlideId: 'host',
    slideIndex: 0,
    beforeSlideIds: ['host'],
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
  let activeId = ''
  let count = 0,
    failAfterWrite = false,
    SDKMismatch = false,
    changeSecondExport = false,
    exports = 0
  const shape = (index: number) => ({
    id: `sdk-real-${index}`,
    name: operations[index]!.name,
    type: operations[index]!.op === 'add_text_box' ? 'TextBox' : 'GeometricShape',
    left: operations[index]!.left,
    top: operations[index]!.top,
    width: operations[index]!.width,
    height: operations[index]!.height,
  })
  const adapter = {
    verifySlides: vi.fn(async (): Promise<VerifySlidesResult> => ({
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
      slideIds: ['host'],
      base64: packages[changeSecondExport && ++exports % 2 === 0 ? Math.min(count + 1, 2) : count]!,
    })),
    listSlideShapes: vi.fn(async () => ({
      slideId: 'host',
      slideIndex: 0,
      shapes: Array.from({ length: count }, (_, i) => ({
        ...shape(i),
        ...(SDKMismatch ? { name: 'wrong' } : {}),
      })),
    })),
    executeDeclarative: vi.fn(async () => {
      expect(binding().readExistingBatch(activeId)).toHaveProperty('inFlightIndex', count)
      count++
      if (failAfterWrite) throw Error('host receipt lost')
      return { createdShapeIds: [`sdk-real-${count - 1}`] }
    }),
  }
  const controller = () =>
    createPresentationNativeAddExecution({
      documentId: () => binding().documentId(),
      readExistingBatch: (id) => binding().readExistingBatch(id),
      writeExistingBatch: (next, expected) => binding().writeExistingBatch(next, expected),
      request,
      adapter,
    })
  return {
    controller,
    slideIR: { ...deck.slides[0]!, elements: additions, claimIds: withClaims ? ['source-1'] : [] },
    claims: withClaims ? deck.claims : [],
    style: deck.style,
    request,
    operations,
    setId: (id: string) => {
      activeId = id
    },
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

async function proposalFixture(withClaims = false) {
  const f = await fixture(withClaims),
    proposals = createStructuredProposalController()
  const factory = (overrides = {}) =>
    createPresentationNativeAddProposal({
      proposals,
      documentId: () => f.binding().documentId(),
      request: f.request,
      adapter: f.adapter,
      readExistingBatch: (id) => f.binding().readExistingBatch(id),
      writeExistingBatch: (next, expected) => f.binding().writeExistingBatch(next, expected),
      execution: f.controller(),
      ...overrides,
    })
  return { ...f, proposals, factory }
}
it('confirms a real PC savepoint and durable settings before each acknowledged host addition', async () => {
  const f = await proposalFixture(),
    p = await f.factory().propose(f.operations)
  f.setId(p.preview.changeId as string)
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  const decision = f.proposals.waitForDecision(p.id)
  await f.proposals.confirm(p.id)
  expect(await decision).toEqual({ status: 'confirmed' })
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(2)
  expect(f.binding().readExistingBatch(p.preview.changeId as string)).toMatchObject({
    state: 'applied',
    nextIndex: 2,
    createdShapeIds: ['sdk-real-0', 'sdk-real-1'],
  })
})
it('does no host write when the PC backup cannot be saved', async () => {
  const f = await proposalFixture(),
    p = await f
      .factory({ request: async () => new Response('{}', { status: 500 }) })
      .propose(f.operations)
  f.setId(p.preview.changeId as string)
  const decision = f.proposals.waitForDecision(p.id)
  await expect(f.proposals.confirm(p.id)).rejects.toThrow()
  expect((await decision).status).toBe('failed')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  expect(f.binding().readExistingBatch(p.preview.changeId as string)).toBeUndefined()
})
it('does no host write when initial durable intent fails', async () => {
  const f = await proposalFixture(),
    p = await f
      .factory({
        writeExistingBatch: async () => {
          throw Error('lost settings receipt')
        },
      })
      .propose(f.operations)
  f.setId(p.preview.changeId as string)
  const decision = f.proposals.waitForDecision(p.id)
  await expect(f.proposals.confirm(p.id)).rejects.toThrow()
  expect((await decision).status).toBe('failed')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('refuses changed package baseline on confirmation', async () => {
  const f = await proposalFixture(),
    p = await f.factory().propose(f.operations)
  f.setId(p.preview.changeId as string)
  f.setCount(1)
  const decision = f.proposals.waitForDecision(p.id)
  await expect(f.proposals.confirm(p.id)).rejects.toThrow()
  expect((await decision).status).toBe('failed')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})

it('uses the complete exported order even when diagnostic verification is truncated', async () => {
  const f = await proposalFixture(),
    ids = ['host', ...Array.from({ length: 24 }, (_, index) => `host-${index}`)]
  const original = f.adapter.exportPresentationPagePackage
  f.adapter.exportPresentationPagePackage = vi.fn(async (...args: Parameters<typeof original>) => ({
    ...(await original(...args)),
    slideIds: ids,
  }))
  f.adapter.verifySlides.mockResolvedValue({
    slideWidth: 960,
    slideHeight: 540,
    slides: [],
    truncated: true,
  } as VerifySlidesResult)
  const p = await f.factory().propose(f.operations)
  expect(p.preview.slideId).toBe('host')
  expect(f.adapter.verifySlides).not.toHaveBeenCalled()
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})
it('does not replay host writes after a native write receipt is lost', async () => {
  const f = await proposalFixture(),
    p = await f.factory().propose(f.operations)
  f.setId(p.preview.changeId as string)
  f.loseHostAck()
  const decision = f.proposals.waitForDecision(p.id)
  await expect(f.proposals.confirm(p.id)).rejects.toThrow()
  expect((await decision).status).toBe('failed')
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(1)
  expect(f.binding().readExistingBatch(p.preview.changeId as string)).toMatchObject({
    state: 'applying',
    nextIndex: 0,
    inFlightIndex: 0,
  })
})
it('does not write a native object after an ambiguous initial intent save', async () => {
  const f = await proposalFixture(),
    p = await f
      .factory({
        writeExistingBatch: async (
          next: PresentationExistingBatch,
          expected: PresentationExistingBatch | undefined,
        ) => {
          await f.binding().writeExistingBatch(next, expected)
          throw Error('lost receipt')
        },
      })
      .propose(f.operations)
  f.setId(p.preview.changeId as string)
  const decision = f.proposals.waitForDecision(p.id)
  await expect(f.proposals.confirm(p.id)).rejects.toThrow()
  expect((await decision).status).toBe('failed')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  expect(f.binding().readExistingBatch(p.preview.changeId as string)).toMatchObject({
    state: 'applying',
    nextIndex: 0,
  })
})
it('refuses pre-aborted proposal construction without creating confirmation or writing host state', async () => {
  const f = await proposalFixture(),
    controller = new AbortController()
  controller.abort()
  await expect(f.factory().propose(f.operations, undefined, controller.signal)).rejects.toThrow(
    'cancelled',
  )
  expect(f.proposals.pending()).toBeUndefined()
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
})

it('routes shared SlideIR through the real skill confirmation, then inspects durable SDK identities', async () => {
  const f = await proposalFixture(true),
    skill = createPowerPointSkill({
      adapter: f.adapter as unknown as PowerPointAdapter,
      proposals: f.proposals,
      nativeAddSavepoint: {
        documentId: () => f.binding().documentId(),
        request: f.request,
        readExistingBatch: (id) => f.binding().readExistingBatch(id),
        writeExistingBatch: (next, expected) => f.binding().writeExistingBatch(next, expected),
      },
    })
  const result = await skill.executeTool({
    id: 'add',
    name: 'add_slide_ir_objects',
    input: { slide_index: 0, slide: f.slideIR, style: f.style, claims: f.claims },
  })
  const p = f.proposals.pending()!
  expect(result.mutated).toBe(false)
  expect(p).toBeDefined()
  f.setId(p.preview.changeId as string)
  const decision = f.proposals.waitForDecision(p.id)
  await f.proposals.confirm(p.id)
  expect(await decision).toEqual({ status: 'confirmed' })
  const inspected = await skill.executeTool({
    id: 'inspect',
    name: 'inspect_slide_ir_addition',
    input: { change_id: p.preview.changeId },
  })
  expect(inspected.mutated).toBe(false)
  expect(inspected.output).toContain('sdk-real-0')
  expect(inspected.output).toContain('sdk-real-1')
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(f.operations.length)
})

it('routes a pure same-page declarative addition through the durable controller with its real tool origin', async () => {
  const f = await proposalFixture()
  const originalRead = f.adapter.listSlideShapes
  f.adapter.listSlideShapes = vi.fn(async (...args: Parameters<typeof originalRead>) => {
    const result = await originalRead(...args)
    return {
      ...result,
      shapes: result.shapes.map((shape) => ({
        ...shape,
        left: shape.left + 0.00003,
        top: shape.top - 0.00003,
      })),
    }
  })
  const skill = createPowerPointSkill({
    adapter: f.adapter as unknown as PowerPointAdapter,
    proposals: f.proposals,
    nativeAddSavepoint: {
      documentId: () => f.binding().documentId(),
      request: f.request,
      readExistingBatch: (id) => f.binding().readExistingBatch(id),
      writeExistingBatch: (next, expected) => f.binding().writeExistingBatch(next, expected),
    },
  })
  const result = await skill.executeTool({
    id: 'generic',
    name: 'execute_office_js',
    input: { program: { version: 1, operations: f.operations } },
  })
  expect(result.isError, result.output).not.toBe(true)
  const p = f.proposals.pending()!
  expect(p.toolName).toBe('execute_office_js')
  expect(p.operation).toBe('execute_office_js')
  expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  f.setId(p.preview.changeId as string)
  await f.proposals.confirm(p.id)
  expect(f.binding().readExistingBatch(p.preview.changeId as string)).toMatchObject({
    state: 'applied',
    nextIndex: 2,
    createdShapeIds: ['sdk-real-0', 'sdk-real-1'],
  })
  expect(f.adapter.executeDeclarative).toHaveBeenCalledTimes(2)
})
it('rejects undeclared text style and mixed or cross-page native additions before host or savepoint writes', async () => {
  for (const kind of ['missing-style', 'mixed', 'cross-page']) {
    const f = await proposalFixture()
    const writeExistingBatch = vi.fn(
      (next: PresentationExistingBatch, expected: PresentationExistingBatch | undefined) =>
        f.binding().writeExistingBatch(next, expected),
    )
    const request = vi.fn(f.request)
    const skill = createPowerPointSkill({
      adapter: f.adapter as unknown as PowerPointAdapter,
      proposals: f.proposals,
      nativeAddSavepoint: {
        documentId: () => f.binding().documentId(),
        request,
        readExistingBatch: (id) => f.binding().readExistingBatch(id),
        writeExistingBatch,
      },
    })
    const operations = structuredClone(f.operations) as any[]
    if (kind === 'missing-style') delete operations[0].fontFace
    if (kind === 'mixed')
      operations.push({ op: 'set_shape_text', slide_index: 0, shape_id: 'old', text: 'changed' })
    if (kind === 'cross-page') operations[1].slide_index = 1
    expect(
      await skill.executeTool({
        id: kind,
        name: 'execute_office_js',
        input: { program: { version: 1, operations } },
      }),
    ).toMatchObject({ isError: true, mutated: false, output: 'office_api_unsupported' })
    expect(f.proposals.pending()).toBeUndefined()
    expect(request).not.toHaveBeenCalled()
    expect(writeExistingBatch).not.toHaveBeenCalled()
    expect(f.adapter.executeDeclarative).not.toHaveBeenCalled()
  }
})
