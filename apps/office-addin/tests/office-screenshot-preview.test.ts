import { afterEach, describe, expect, it, vi } from 'vitest'
import { PNG } from 'pngjs'
import {
  OFFICE_SCREENSHOT_PREVIEW_BYTES,
  OFFICE_SCREENSHOT_SOURCE_BYTES,
} from '@wiswork/agent-core'
import { prepareOfficeScreenshotPreview } from '../src/agent/office-screenshot-preview'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill'
import type { PowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
import { createOfficeAgentSession } from '../src/agent/use-office-agent'
import { createOfficeDiagnostics } from '../src/diagnostics/office-diagnostics'

const png = {
  mime: 'image/png',
  base64:
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4AWP4DwQACfsD/c8LaHIAAAAASUVORK5CYII=',
}
const jpeg =
  '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD5/ooooA//2Q=='

afterEach(() => {
  vi.unstubAllGlobals()
  vi.useRealTimers()
})

describe('Office model screenshot previews', () => {
  it('records failed legacy screenshot serialization before marking its tool card complete', async () => {
    let handler: ((call: any) => Promise<{ output: string; isError?: boolean }>) | undefined
    const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'test' })
    const session = createOfficeAgentSession({
      transport: { stream: () => ({ cancel() {} }) },
      skill: {
        id: 'test',
        systemPrompt: '',
        tools: [{ name: 'screenshot_slide', description: '', inputSchema: {} }],
        executeTool: async () => ({
          output: '{"visualAvailableToModel":true}',
          summary: 'Screenshot',
          mutated: false,
        }),
      },
      proposals: createStructuredProposalController(),
      diagnostics,
      remoteTools: {
        setToolHandler(next) {
          handler = next
        },
      },
    })
    try {
      expect(
        await handler!({
          turnId: 'turn_12345678',
          callId: 'call_12345678',
          generation: 1,
          toolName: 'screenshot_slide',
          input: { slide_index: 0 },
          signal: new AbortController().signal,
        }),
      ).toEqual({ output: 'office_screenshot_unavailable', isError: true })
      expect(session.snapshot().timeline.find((event) => event.kind === 'tool')).toMatchObject({
        state: 'error',
        output: 'office_screenshot_unavailable',
      })
      expect(diagnostics.snapshot().events.at(-1)).toMatchObject({
        tool: 'screenshot_slide',
        error_code: 'office_read_failed',
      })
    } finally {
      session.dispose()
    }
  })
  it('validates a small native image and preserves its exact bytes', async () => {
    const close = vi.fn()
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width: 1, height: 1, close })),
    )
    expect(await prepareOfficeScreenshotPreview(png)).toEqual(png)
    expect(close).toHaveBeenCalledOnce()
  })

  it.each(['invalid', 'oversized source', 'pixel bomb', 'unavailable decoder'])(
    'fails closed for %s',
    async (scenario) => {
      const create = vi.fn(async () => ({ width: 1, height: 1, close() {} }))
      vi.stubGlobal('createImageBitmap', scenario === 'unavailable decoder' ? undefined : create)
      let source = png
      if (scenario === 'invalid') source = { ...png, base64: 'iVBORw0KGgoAAAA=' }
      if (scenario === 'oversized source')
        source = {
          ...png,
          base64: Buffer.alloc(OFFICE_SCREENSHOT_SOURCE_BYTES + 1).toString('base64'),
        }
      if (scenario === 'pixel bomb') {
        const bytes = Buffer.from(png.base64, 'base64')
        bytes.writeUInt32BE(100_000, 16)
        source = { ...png, base64: bytes.toString('base64') }
      }
      await expect(prepareOfficeScreenshotPreview(source)).rejects.toThrow(
        'office_screenshot_unavailable',
      )
      expect(create).not.toHaveBeenCalled()
    },
  )

  it.each(['success', 'too large', 'cancel'])(
    'bounds canvas attempts and keeps the original screenshot untouched: %s',
    async (scenario) => {
      const original = new PNG({ width: 300, height: 300 })
      let seed = 123
      for (let index = 0; index < original.data.length; index++) {
        seed = Math.imul(seed, 1664525) + 1013904223
        original.data[index] = seed >>> 24
      }
      const source = { mime: 'image/png', base64: PNG.sync.write(original).toString('base64') }
      expect(Buffer.from(source.base64, 'base64').length).toBeGreaterThan(
        OFFICE_SCREENSHOT_PREVIEW_BYTES,
      )
      const before = source.base64
      const close = vi.fn()
      vi.stubGlobal('createImageBitmap', async () => ({ width: 300, height: 300, close }))
      const controller = new AbortController()
      const toBlob = vi.fn((callback: (value: Blob) => void) => {
        if (scenario === 'cancel') controller.abort()
        callback(
          new Blob(
            [
              scenario === 'too large'
                ? Buffer.alloc(OFFICE_SCREENSHOT_PREVIEW_BYTES + 1)
                : Buffer.from(jpeg, 'base64'),
            ],
            { type: 'image/jpeg' },
          ),
        )
      })
      const canvas = {
        width: 0,
        height: 0,
        getContext: () => ({ fillStyle: '', fillRect() {}, drawImage() {} }),
        toBlob,
      }
      vi.stubGlobal('document', { createElement: () => canvas })
      const result = prepareOfficeScreenshotPreview(source, controller.signal)
      if (scenario === 'success') expect(await result).toEqual({ mime: 'image/jpeg', base64: jpeg })
      else
        await expect(result).rejects.toThrow(
          scenario === 'cancel' ? 'tool_cancelled' : 'office_screenshot_unavailable',
        )
      expect(toBlob).toHaveBeenCalledTimes(scenario === 'too large' ? 3 : 1)
      expect(close).toHaveBeenCalledOnce()
      expect(canvas.width).toBe(0)
      expect(source.base64).toBe(before)
    },
  )

  it.each(['cancel', 'timeout'])(
    'closes a late native decoder after %s',
    async (scenario) => {
      let finish!: (value: { width: number; height: number; close(): void }) => void
      const create = vi.fn(
        () =>
          new Promise((resolve) => {
            finish = resolve
          }),
      )
      vi.stubGlobal('createImageBitmap', create)
      const controller = new AbortController()
      const result = prepareOfficeScreenshotPreview(png, controller.signal)
      const rejected = expect(result).rejects.toThrow(
        scenario === 'cancel' ? 'tool_cancelled' : 'office_screenshot_unavailable',
      )
      await vi.waitFor(() => expect(create).toHaveBeenCalledOnce())
      if (scenario === 'cancel') controller.abort()
      else {
        // The decoder timeout is bounded without replacing native image parsing.
        await new Promise((resolve) => setTimeout(resolve, 10_050))
      }
      await rejected
      const close = vi.fn()
      finish({ width: 1, height: 1, close })
      await vi.waitFor(() => expect(close).toHaveBeenCalledOnce())
    },
    15_000,
  )

  it.each(['preparation', 'envelope'])(
    'keeps the screenshot gate dirty after %s failure',
    async (scenario) => {
      const proposals = createStructuredProposalController()
      const skill = createPowerPointSkill({
        platform: 'Mac',
        proposals,
        adapter: {
          snapshotSlide: async () => ({ slideId: 's1', fingerprint: 'same' }),
          editSlideText: async () => undefined,
          readSlideText: async () => ({ text: 'New' }),
          screenshotSlide: async () => png,
        } as unknown as PowerPointAdapter,
        prepareScreenshot: async () => {
          if (scenario === 'preparation') throw new Error('office_screenshot_unavailable')
          return {
            ...png,
            base64: Buffer.alloc(OFFICE_SCREENSHOT_PREVIEW_BYTES + 1).toString('base64'),
          }
        },
      })
      await skill.executeTool({
        id: 'edit',
        name: 'edit_slide_text',
        input: { slide_index: 0, shape_id: 'title', text: 'New' },
      })
      const result = await skill.executeTool({
        id: 'shot',
        name: 'screenshot_slide',
        input: { slide_index: 0 },
      })
      expect(result).toMatchObject({ isError: true, output: 'office_screenshot_unavailable' })
      expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toContain(
        'screenshot_slide',
      )
    },
  )

  it('describes the prepared model preview while retaining the original native UI image', async () => {
    const image = { mime: 'image/jpeg', base64: jpeg }
    const skill = createPowerPointSkill({
      platform: 'Mac',
      proposals: createStructuredProposalController(),
      adapter: { screenshotSlide: async () => png } as unknown as PowerPointAdapter,
      prepareScreenshot: async () => image,
    })
    const result = await skill.executeTool({
      id: 'shot',
      name: 'screenshot_slide',
      input: { slide_index: 0 },
    })
    expect(result.isError).not.toBe(true)
    expect(result.modelContent).toEqual([{ type: 'image', image }])
    expect(JSON.parse(result.output)).toMatchObject({
      mime: image.mime,
      bytes: Buffer.from(jpeg, 'base64').length,
      visualAvailableToModel: true,
    })
    expect(result.display?.items?.[0]?.url).toBe(`data:${png.mime};base64,${png.base64}`)
  })

  it('does not clear a newer mutation while an older screenshot is being prepared', async () => {
    let complete!: (value: typeof png) => void
    const prepare = vi.fn(
      () =>
        new Promise<typeof png>((resolve) => {
          complete = resolve
        }),
    )
    const skill = createPowerPointSkill({
      platform: 'Mac',
      proposals: createStructuredProposalController(),
      adapter: {
        screenshotSlide: async () => png,
        snapshotSlide: async () => ({ slideId: 's1', fingerprint: 'same' }),
        editSlideText: async () => undefined,
      } as unknown as PowerPointAdapter,
      prepareScreenshot: prepare,
    })
    const screenshot = skill.executeTool({
      id: 'shot',
      name: 'screenshot_slide',
      input: { slide_index: 0 },
    })
    await vi.waitFor(() => expect(prepare).toHaveBeenCalledOnce())
    await skill.executeTool({
      id: 'edit',
      name: 'edit_slide_text',
      input: { slide_index: 0, shape_id: 'title', text: 'Newer' },
    })
    complete(png)
    expect(await screenshot).toMatchObject({
      isError: true,
      output: 'office_screenshot_unavailable',
    })
    expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toContain(
      'screenshot_slide',
    )
  })
})
