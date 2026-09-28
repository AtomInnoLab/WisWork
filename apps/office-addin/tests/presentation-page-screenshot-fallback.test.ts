import { expect, it, vi } from 'vitest'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import type { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
import { createPresentationPageScreenshotFallback } from '../src/skills/powerpoint/presentation-page-screenshot-fallback'

const png =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII='
const hostError = () =>
  Object.assign(new Error('office_read_failed'), { code: 'office_screenshot_unavailable' })

async function fixture() {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const base64 = Buffer.from((await compilePresentationDeck(deck)).bytes).toString('base64')
  const exportPage = vi.fn(async () => ({ slideId: 'host-1', slideIds: ['host-1'], base64 }))
  const inspectPage = vi.fn(
    async (_slideId: string, _signal?: AbortSignal, replacement?: string) => {
      if (!replacement) throw hostError()
      return {
        slideId: 'host-1',
        screenshot: { mime: 'image/png', base64: replacement, renderer: 'libreoffice' },
      }
    },
  )
  const adapter = {
    exportPresentationPagePackage: exportPage,
    inspectPresentationPage: inspectPage,
  } as unknown as BrowserPowerPointAdapter
  let received = new Uint8Array(0)
  let meta: Record<string, unknown> = {}
  let releaseFails = false
  let chunkFails = false
  const operations: string[] = []
  const request = vi.fn(async (body: unknown) => {
    const data = body as Record<string, unknown>
    operations.push(data.operation as string)
    switch (data.operation) {
      case 'existing_page_backup_begin':
        meta = { ...data, status: 'uploading', receivedBytes: 0 }
        break
      case 'existing_page_backup_chunk': {
        if (chunkFails) return new Response(JSON.stringify({ error: 'busy' }), { status: 503 })
        const part = Uint8Array.from(atob(data.base64 as string), (c) => c.charCodeAt(0))
        const next = new Uint8Array(received.length + part.length)
        next.set(received)
        next.set(part, received.length)
        received = next
        meta.receivedBytes = received.length
        break
      }
      case 'existing_page_backup_finish':
        meta.status = 'ready'
        break
      case 'existing_page_backup_render':
        expect(Buffer.from(received).toString('base64')).toBe(base64)
        return new Response(
          JSON.stringify({
            backupId: data.backupId,
            hostSlideId: 'host-1',
            sha256: meta.sha256,
            renderer: 'libreoffice',
            mime: 'image/png',
            base64: png,
          }),
        )
      case 'existing_page_backup_release':
        if (releaseFails) return new Response('{}', { status: 503 })
        meta.status = 'released'
        break
      case 'existing_page_backup_abandon':
        meta.status = 'abandoned'
        break
    }
    return new Response(JSON.stringify(meta))
  })
  const inspect = createPresentationPageScreenshotFallback({
    adapter,
    request,
    documentId: async () => 'document-1',
  })
  return {
    inspect,
    inspectPage,
    exportPage,
    request,
    operations,
    base64,
    failRelease: () => {
      releaseFails = true
    },
    failChunk: () => {
      chunkFails = true
    },
  }
}

it('uses a bounded PC PNG only after export identity and semantic content are rechecked', async () => {
  const f = await fixture()
  expect((await f.inspect('host-1')).screenshot.base64).toBe(png)
  expect(f.inspectPage).toHaveBeenCalledTimes(2)
  expect(f.inspectPage).toHaveBeenLastCalledWith('host-1', undefined, png)
  expect(f.exportPage).toHaveBeenCalledTimes(3)
  expect(f.operations).toContain('existing_page_backup_render')
  expect(f.operations.at(-1)).toBe('existing_page_backup_release')
})

it('rejects a changed host page after PC rendering', async () => {
  const f = await fixture()
  f.exportPage.mockResolvedValueOnce({ slideId: 'host-1', slideIds: ['host-1'], base64: f.base64 })
  // The second export proves that the original host page changed during rendering.
  f.exportPage.mockResolvedValueOnce({ slideId: 'host-1', slideIds: ['host-2'], base64: f.base64 })
  await expect(f.inspect('host-1')).rejects.toThrow('presentation_qa_stale')
  expect(f.inspectPage).toHaveBeenCalledTimes(1)
})

it('rejects a page change while rereading native geometry after fallback rendering', async () => {
  const f = await fixture()
  f.exportPage.mockResolvedValueOnce({ slideId: 'host-1', slideIds: ['host-1'], base64: f.base64 })
  f.exportPage.mockResolvedValueOnce({ slideId: 'host-1', slideIds: ['host-1'], base64: f.base64 })
  f.exportPage.mockResolvedValueOnce({ slideId: 'host-1', slideIds: ['host-2'], base64: f.base64 })
  await expect(f.inspect('host-1')).rejects.toThrow('presentation_qa_stale')
  expect(f.inspectPage).toHaveBeenCalledTimes(2)
})

it('does not hide an unconfirmed temporary backup release', async () => {
  const f = await fixture()
  f.failRelease()
  await expect(f.inspect('host-1')).rejects.toThrow('presentation_page_backup_cleanup_failed')
  expect(f.operations.filter((op) => op === 'existing_page_backup_release')).toHaveLength(2)
  expect(f.inspectPage).toHaveBeenCalledTimes(1)
})

it('abandons an incomplete upload before reporting the screenshot as unavailable', async () => {
  const f = await fixture()
  f.failChunk()
  await expect(f.inspect('host-1')).rejects.toThrow('office_read_failed')
  expect(f.operations).toContain('existing_page_backup_abandon')
  expect(f.operations).not.toContain('existing_page_backup_render')
})
