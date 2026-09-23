import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationService } from '../src/main/presentation-service'

it('persists real eight-page OOXML source IDs and restores their original ordering', async () => {
  const userDataPath = await mkdtemp(join(tmpdir(), 'presentation-pages-'))
  try {
    const deck = benchmarkDeck()
    const service = createPresentationService({ userDataPath })
    const read = (bytes: Uint8Array) => JSON.parse(Buffer.from(bytes).toString('utf8'))
    const signal = new AbortController().signal
    const result = read(
      await service(
        { operation: 'compile', documentId: 'page-document', requestId: 'pages-first', deck },
        signal,
      ),
    )
    expect(result.error).toBeUndefined()
    const zip = await JSZip.loadAsync(Buffer.from(result.pptxBase64, 'base64'))
    const xml = await zip.file('ppt/presentation.xml')!.async('string')
    const ids = [...xml.matchAll(/<p:sldId\s+id="(\d+)"/g)].map((match) => `${match[1]}#`)
    expect(result.pages).toEqual(
      deck.slides.map((slide, index) => ({
        id: slide.id,
        title: slide.title,
        sourceSlideId: ids[index],
      })),
    )
    expect(result.pages).toHaveLength(8)
    const reload = createPresentationService({ userDataPath })
    const restored = read(
      await reload({ operation: 'get', documentId: 'page-document', projectId: deck.id }, signal),
    )
    expect(restored.pages).toEqual(result.pages)
    expect(restored.pptxBase64).toBe(result.pptxBase64)
  } finally {
    await rm(userDataPath, { recursive: true, force: true })
  }
})
