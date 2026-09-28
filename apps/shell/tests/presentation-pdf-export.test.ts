import { mkdtempSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it, vi } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { createPresentationService } from '../src/main/presentation-service'
import { libreOfficeCommands } from '../src/main/presentation-page-render'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'

const sofficeAvailable = libreOfficeCommands().some(
  (command) => spawnSync(command, ['--version'], { timeout: 5_000 }).status === 0,
)

const bytes = new Uint8Array([1, 2, 3])
const report = {
  deckId: 'deck',
  slideCount: 1,
  elementCount: 1,
  geometry: [],
  checks: {
    structure: 'passed' as const,
    geometry: 'passed' as const,
    render: 'not_run' as const,
    sources: 'not_verified' as const,
    roundTrip: 'not_run' as const,
  },
}
const deck = {
  version: 1,
  id: 'deck',
  title: 'One',
  style: { fontFace: 'Arial', background: 'FFFFFF', textColor: '111111', accentColor: '3366FF' },
  assets: [],
  claims: [],
  slides: [
    {
      id: 'slide',
      title: 'Title',
      elements: [{ id: 'text', kind: 'text', text: 'Hello', x: 1, y: 1, w: 4, h: 1 }],
    },
  ],
}
const decode = (value: Uint8Array) => JSON.parse(Buffer.from(value).toString('utf8'))
const signal = () => new AbortController().signal

it('exports PDF only for an exact compiled request bound to the same document', async () => {
  const document = await PDFDocument.create()
  document.addPage([960, 540])
  const pdf = Buffer.from(await document.save())
  const renderPdf = vi.fn(async () => pdf)
  const service = createPresentationService({
    userDataPath: mkdtempSync(join(tmpdir(), 'presentation-pdf-')),
    compile: async () => ({ bytes, report }),
    renderPdf,
  })
  const request = {
    operation: 'export_pdf',
    documentId: 'office:/opaque?document',
    projectId: 'deck',
    requestId: 'first',
  }
  expect(decode(await service(request, signal()))).toEqual({ error: 'not_found' })
  await service({ ...request, operation: 'compile', deck }, signal())
  expect(decode(await service(request, signal()))).toEqual({
    status: 'exported',
    source: 'compiled',
    projectId: 'deck',
    requestId: 'first',
    slideCount: 1,
    pdfBase64: pdf.toString('base64'),
  })
  expect(Buffer.from(renderPdf.mock.calls[0]![0])).toEqual(Buffer.from(bytes))
  expect(decode(await service({ ...request, documentId: 'another' }, signal()))).toEqual({
    error: 'document_mismatch',
  })
  expect(decode(await service({ ...request, requestId: 'another' }, signal()))).toEqual({
    error: 'not_found',
  })
  expect(renderPdf).toHaveBeenCalledOnce()
})

it('rejects a PDF that lost a compiled page', async () => {
  const document = await PDFDocument.create()
  document.addPage([960, 540])
  const service = createPresentationService({
    userDataPath: mkdtempSync(join(tmpdir(), 'presentation-pdf-')),
    compile: async () => ({ bytes, report: { ...report, slideCount: 2 } }),
    renderPdf: async () => document.save(),
  })
  const request = {
    operation: 'export_pdf',
    documentId: 'office:/opaque?document',
    projectId: 'deck',
    requestId: 'first',
  }
  await service({ ...request, operation: 'compile', deck }, signal())
  expect(decode(await service(request, signal()))).toEqual({ error: 'renderer_unavailable' })
})

it.skipIf(!sofficeAvailable)(
  'compiles and exports the same real eight-page project through the PC service',
  async () => {
    const service = createPresentationService({
      userDataPath: mkdtempSync(join(tmpdir(), 'presentation-pdf-')),
    })
    const request = {
      documentId: 'office:/eight-pages',
      requestId: 'eight',
    }
    const deck = benchmarkDeck()
    const compiled = decode(
      await service({ operation: 'compile', ...request, projectId: deck.id, deck }, signal()),
    )
    expect(compiled.status).toBe('compiled')
    const exported = decode(
      await service({ operation: 'export_pdf', ...request, projectId: deck.id }, signal()),
    )
    expect(exported.status).toBe('exported')
    expect(exported.slideCount).toBe(8)
    expect((await PDFDocument.load(Buffer.from(exported.pdfBase64, 'base64'))).getPageCount()).toBe(
      8,
    )
  },
  75_000,
)

it('exports all completed production pages in order and rejects incomplete production', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'presentation-pdf-production-'))
  const onePage = await PDFDocument.create()
  onePage.addPage([960, 540])
  const rendered = await onePage.save()
  const renderPdf = vi.fn(async () => rendered)
  const service = createPresentationService({
    userDataPath,
    compile: compilePresentationDeck,
    renderPdf,
  })
  const deck = benchmarkPlannedDeck()
  const request = {
    documentId: 'office:/production',
    projectId: deck.id,
    requestId: 'pages',
    source: 'production',
  }
  expect(
    decode(
      await service(
        {
          operation: 'save_plan',
          documentId: request.documentId,
          projectId: deck.id,
          expectedRevision: 0,
          plan: benchmarkPlan(),
        },
        signal(),
      ),
    ).revision,
  ).toBe(1)
  expect(
    decode(
      await service(
        {
          operation: 'production_begin',
          documentId: request.documentId,
          projectId: deck.id,
          requestId: 'pages',
          planRevision: 1,
          deck,
        },
        signal(),
      ),
    ).status,
  ).toBe('pending')
  expect(decode(await service({ operation: 'export_pdf', ...request }, signal()))).toEqual({
    error: 'page_not_ready',
  })
  expect(
    decode(
      await service(
        {
          operation: 'production_run',
          documentId: request.documentId,
          projectId: deck.id,
          requestId: 'pages',
        },
        signal(),
      ),
    ).status,
  ).toBe('compiled')
  const exported = decode(await service({ operation: 'export_pdf', ...request }, signal()))
  expect(exported.status).toBe('exported')
  expect(exported.slideCount).toBe(8)
  expect((await PDFDocument.load(Buffer.from(exported.pdfBase64, 'base64'))).getPageCount()).toBe(8)
  expect(renderPdf).toHaveBeenCalledTimes(8)
})

it.skipIf(!sofficeAvailable)(
  'renders and merges real production pages into one PDF',
  async () => {
    const service = createPresentationService({
      userDataPath: mkdtempSync(join(tmpdir(), 'presentation-pdf-production-')),
    })
    const plan = benchmarkPlan()
    const deck = benchmarkPlannedDeck()
    plan.slides = plan.slides.slice(0, 2)
    deck.slides = deck.slides.slice(0, 2)
    const common = { documentId: 'office:/two-pages', projectId: deck.id, requestId: 'pages' }
    expect(
      decode(
        await service(
          {
            operation: 'save_plan',
            documentId: common.documentId,
            projectId: deck.id,
            expectedRevision: 0,
            plan,
          },
          signal(),
        ),
      ).revision,
    ).toBe(1)
    await service({ operation: 'production_begin', ...common, planRevision: 1, deck }, signal())
    expect(decode(await service({ operation: 'production_run', ...common }, signal())).status).toBe(
      'compiled',
    )
    const exported = decode(
      await service({ operation: 'export_pdf', ...common, source: 'production' }, signal()),
    )
    expect(exported.status).toBe('exported')
    expect(exported.source).toBe('production')
    expect((await PDFDocument.load(Buffer.from(exported.pdfBase64, 'base64'))).getPageCount()).toBe(
      2,
    )
  },
  120_000,
)
