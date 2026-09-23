import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import type { PresentationCompileReport } from '@wiswork/pptx-engine/presentation'
import { createPresentationService } from '../src/main/presentation-service.js'

vi.mock('@wiswork/pptx-engine/presentation-compiler', () => ({ compilePresentationDeck: vi.fn() }))

const report: PresentationCompileReport = {
  deckId: 'deck',
  slideCount: 1,
  elementCount: 1,
  geometry: [],
  checks: {
    structure: 'passed',
    geometry: 'passed',
    render: 'not_run',
    sources: 'not_verified',
    roundTrip: 'not_run',
  },
}
const input = {
  operation: 'compile',
  documentId: 'office:/opaque?document',
  requestId: 'first',
  deck: {
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
  },
}
const decode = (bytes: Uint8Array) => JSON.parse(Buffer.from(bytes).toString('utf8'))
const signal = () => new AbortController().signal
const root = () => mkdtempSync(join(tmpdir(), 'presentation-service-'))
const result = () => ({ bytes: new Uint8Array([1, 2, 3]), report })

describe('presentation service', () => {
  it('deduplicates concurrent calls and persists full result across recreation', async () => {
    const userDataPath = root()
    const compile = vi.fn(async () => result())
    const service = createPresentationService({ userDataPath, compile })
    const [a, b] = await Promise.all([service(input, signal()), service(input, signal())])
    expect(decode(a)).toEqual({
      projectId: 'deck',
      requestId: 'first',
      status: 'compiled',
      pptxBase64: 'AQID',
      report,
    })
    expect(decode(b)).toEqual(decode(a))
    const reload = createPresentationService({ userDataPath, compile })
    expect(decode(await reload(input, signal()))).toEqual(decode(a))
    expect(compile).toHaveBeenCalledTimes(1)
    expect(
      decode(
        await reload(
          { operation: 'get', projectId: 'deck', documentId: input.documentId },
          signal(),
        ),
      ),
    ).toEqual(decode(a))
    expect(
      decode(await reload({ ...input, deck: { ...input.deck, title: 'Changed' } }, signal())),
    ).toEqual({ error: 'request_conflict' })
  })
  it('preserves the last ready output on failed or cancelled compilation and can retry', async () => {
    const userDataPath = root()
    const compile = vi.fn(async () => result())
    const service = createPresentationService({ userDataPath, compile })
    const ready = decode(await service(input, signal()))
    compile.mockRejectedValueOnce(new Error('/secret/arbitrary/path'))
    const second = { ...input, requestId: 'second' }
    expect(decode(await service(second, signal()))).toEqual({ error: 'compile_failed' })
    const controller = new AbortController()
    compile.mockImplementationOnce(async () => {
      controller.abort()
      return result()
    })
    expect(decode(await service(second, controller.signal))).toEqual({ error: 'aborted' })
    expect(
      decode(
        await service(
          { operation: 'get', projectId: 'deck', documentId: input.documentId },
          signal(),
        ),
      ),
    ).toEqual(ready)
    expect(
      decode(await createPresentationService({ userDataPath, compile })(second, signal()))
        .requestId,
    ).toBe('second')
  })
  it('denies document mismatch, traversal and oversized output with bounded errors', async () => {
    const compile = vi.fn(async () => result())
    const service = createPresentationService({ userDataPath: root(), compile })
    expect(decode(await service({ ...input, outputPath: '/tmp/arbitrary' }, signal()))).toEqual({
      error: 'invalid_request',
    })
    expect(
      decode(
        await service(
          { operation: 'get', projectId: 'deck', documentId: input.documentId, requestId: 'first' },
          signal(),
        ),
      ),
    ).toEqual({ error: 'invalid_request' })
    expect(
      decode(await service({ ...input, deck: { ...input.deck, arbitraryCode: 'evil' } }, signal())),
    ).toEqual({ error: 'invalid_deck' })
    await service(input, signal())
    expect(decode(await service({ ...input, documentId: 'other' }, signal()))).toEqual({
      error: 'document_mismatch',
    })
    for (const projectId of ['../x', '/tmp/x', '..', 'a/b']) {
      expect(
        decode(await service({ operation: 'get', documentId: 'doc', projectId }, signal())),
      ).toEqual({ error: 'invalid_request' })
    }
    compile.mockResolvedValueOnce({ bytes: new Uint8Array(11 * 1024 * 1024), report })
    expect(decode(await service({ ...input, requestId: 'big' }, signal()))).toEqual({
      error: 'output_too_large',
    })
    expect(
      decode(
        await service({ operation: 'get', projectId: 'missing', documentId: 'doc' }, signal()),
      ),
    ).toEqual({ error: 'not_found' })
  })
})
