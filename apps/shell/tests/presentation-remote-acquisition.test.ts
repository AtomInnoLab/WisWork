import { mkdtempSync, rmSync, readdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { afterEach, expect, it, vi } from 'vitest'
import { createPresentationAttachmentService } from '../src/main/presentation-attachments'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function setup(fetchImage: (url: string, signal: AbortSignal) => Promise<Response | null>) {
  const root = mkdtempSync(join(tmpdir(), 'remote-acquisition-'))
  roots.push(root)
  const service = createPresentationAttachmentService({ userDataPath: root, fetchImage })
  const call = (body: Record<string, unknown>, signal = new AbortController().signal) =>
    service({ documentId: 'doc', ...body }, signal)
  return { call, root }
}
it('records an active download cancellation and rejects corrupt history without retrying', async () => {
  const controller = new AbortController()
  const fetchImage = vi.fn(async () => {
    controller.abort()
    return null
  })
  const { call, root } = setup(fetchImage)
  await expect(
    call(
      { operation: 'attachment_import_url', url: 'https://8.8.8.8/image?token=secret' },
      controller.signal,
    ),
  ).rejects.toThrow('aborted')
  expect(await call({ operation: 'attachment_acquisition_history' })).toMatchObject({
    records: [{ state: 'rejected', error: 'aborted' }],
  })
  expect(fetchImage).toHaveBeenCalledTimes(1)
  const dir = join(root, 'presentation-acquisition-history')
  writeFileSync(join(dir, readdirSync(dir)[0]!), 'sensitive disk exception')
  await expect(call({ operation: 'attachment_acquisition_history' })).rejects.toThrow(
    /^invalid_state$/,
  )
  await expect(
    call({ operation: 'attachment_import_url', url: 'https://8.8.8.8/image' }),
  ).rejects.toThrow(/^invalid_state$/)
  expect(fetchImage).toHaveBeenCalledTimes(1)
})
it('saves the staged animated attachment as rejected and reuses its download', async () => {
  const gif = Buffer.from('R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==', 'base64')
  // Two image descriptors make the inspected GIF animated without requiring image normalization.
  const animated = Buffer.concat([gif.subarray(0, gif.length - 1), gif.subarray(19)])
  const fetchImage = vi.fn(
    async () => new Response(animated, { headers: { 'content-type': 'image/gif' } }),
  )
  const { call } = setup(fetchImage)
  const request = {
    operation: 'attachment_import_url',
    url: 'https://8.8.8.8/animated.gif',
    stageAnimated: true,
  }
  const first = await call(request)
  expect(first).toMatchObject({ status: 'failed', error: 'animated_image_unsupported' })
  await call(request)
  expect(fetchImage).toHaveBeenCalledTimes(1)
  expect(await call({ operation: 'attachment_acquisition_history' })).toMatchObject({
    totalAttempts: 2,
    records: [
      {
        state: 'rejected',
        error: 'animated_image_unsupported',
        attachmentId: (first as { attachmentId: string }).attachmentId,
      },
      { state: 'rejected', error: 'animated_image_unsupported' },
    ],
  })
})
