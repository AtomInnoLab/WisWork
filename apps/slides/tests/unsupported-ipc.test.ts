import { describe, expect, it } from 'vitest'
import {
  registerUnsupportedCloudIpc,
  registerUnsupportedMediaIpc,
} from '../src/main/unsupported-ipc'

interface FakeEvent {
  sender: { id: number }
}
type Handler = (event: FakeEvent, ...args: unknown[]) => unknown

function harness() {
  const handlers = new Map<string, Handler>()
  const ipcMain = { handle: (channel: string, handler: Handler) => handlers.set(channel, handler) }
  const trusted = new Set([7])
  registerUnsupportedMediaIpc(ipcMain, (id) => trusted.has(id))
  registerUnsupportedCloudIpc(ipcMain, (id) => trusted.has(id))
  const invoke = (channel: string, senderId: number, ...args: unknown[]) =>
    handlers.get(channel)!({ sender: { id: senderId } }, ...args)
  return { invoke }
}

const imagePayload = { prompt: 'draw a chart' }
const mediaPayload = { mediaUrls: ['https://example.test/a.png'], requirements: 'summarize' }
const cloudPayload = { brief: 'one page', width: 1280, height: 720 }

function expectUnsupported(value: unknown): void {
  expect(value).toMatchObject({ ok: false, errorCode: 'unsupported_feature' })
}

describe('unsupported main-process IPC', () => {
  it.each([
    'ai:generate-image',
    'ai:analyze-media',
    'slides:cloud-gen-status',
    'slides:cloud-page-generate',
  ])('%s rejects an untrusted sender', async (channel) => {
    const { invoke } = harness()
    const payload =
      channel === 'ai:generate-image'
        ? imagePayload
        : channel === 'ai:analyze-media'
          ? mediaPayload
          : channel === 'slides:cloud-page-generate'
            ? cloudPayload
            : undefined
    await expect(
      Promise.resolve().then(() =>
        payload === undefined ? invoke(channel, 99) : invoke(channel, 99, payload),
      ),
    ).rejects.toMatchObject({ code: 'untrusted_sender' })
  })

  it('returns the stable result for every valid payload', async () => {
    const { invoke } = harness()
    expectUnsupported(await invoke('ai:generate-image', 7, imagePayload))
    expectUnsupported(await invoke('ai:analyze-media', 7, mediaPayload))
    expect(await invoke('slides:cloud-gen-status', 7)).toEqual({
      enabled: false,
      errorCode: 'unsupported_feature',
    })
    expectUnsupported(await invoke('slides:cloud-page-generate', 7, cloudPayload))
    expectUnsupported(await invoke('slides:cloud-page-generate', 7, { brief: 'legacy page' }))
    expectUnsupported(
      await invoke('slides:cloud-page-generate', 7, { brief: 'width only', width: 1280 }),
    )
    expectUnsupported(
      await invoke('slides:cloud-page-generate', 7, { brief: 'height only', height: 720 }),
    )
  })

  it.each([
    ['ai:generate-image', null],
    ['ai:generate-image', { ...imagePayload, extra: true }],
    ['ai:generate-image', { prompt: 'x'.repeat(32_001) }],
    ['ai:analyze-media', { ...mediaPayload, extra: true }],
    ['ai:analyze-media', { mediaUrls: [], requirements: 'x' }],
    ['slides:cloud-gen-status', {}],
    ['slides:cloud-page-generate', { ...cloudPayload, extra: true }],
    ['slides:cloud-page-generate', { brief: 'x'.repeat(32_001) }],
    ['slides:cloud-page-generate', { brief: 'x', width: 0 }],
    ['slides:cloud-page-generate', { brief: 'x', height: '720' }],
  ])('%s rejects malformed, extra, or oversized payloads', async (channel, payload) => {
    const { invoke } = harness()
    await expect(
      Promise.resolve().then(() => invoke(channel as string, 7, payload)),
    ).rejects.toMatchObject({ code: expect.stringMatching(/invalid_payload|payload_too_large/) })
  })
})
