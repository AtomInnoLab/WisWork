import {
  OFFICE_SCREENSHOT_PREVIEW_BYTES,
  OFFICE_SCREENSHOT_SOURCE_BYTES,
  officeScreenshotBytes,
  type AgentImage,
} from '@wiswork/agent-core'
import { validateSkillPackageImage } from '../skills/shared/skill-package.js'
import { MAX_IMAGE_DIMENSION, MAX_IMAGE_PIXELS } from '../skills/shared/import-media.js'

const PREVIEW_TIMEOUT_MS = 10_000

function bounded<T>(
  pending: Promise<T>,
  signal?: AbortSignal,
  late?: (value: T) => void,
): Promise<T> {
  return new Promise((resolve, reject) => {
    let settled = false
    const finish = (value?: T, error?: Error) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      signal?.removeEventListener('abort', abort)
      if (error) reject(error)
      else resolve(value!)
    }
    const abort = () => finish(undefined, new Error('tool_cancelled'))
    const timer = setTimeout(
      () => finish(undefined, new Error('office_screenshot_unavailable')),
      PREVIEW_TIMEOUT_MS,
    )
    signal?.addEventListener('abort', abort, { once: true })
    if (signal?.aborted) abort()
    pending.then(
      (value) => (settled ? late?.(value) : finish(value)),
      () => finish(undefined, new Error('office_screenshot_unavailable')),
    )
  })
}

/** Native UI keeps its original PNG; only the model preview is normalized. */
export async function prepareOfficeScreenshotPreview(
  image: AgentImage,
  outerSignal?: AbortSignal,
): Promise<AgentImage> {
  if (outerSignal?.aborted) throw new Error('tool_cancelled')
  const bytes = officeScreenshotBytes(image, OFFICE_SCREENSHOT_SOURCE_BYTES)
  if (image.mime !== 'image/png') throw new Error('office_screenshot_unavailable')
  const header = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const width = header.getUint32(16)
  const height = header.getUint32(20)
  if (
    !width ||
    !height ||
    width > MAX_IMAGE_DIMENSION ||
    height > MAX_IMAGE_DIMENSION ||
    width * height > MAX_IMAGE_PIXELS
  )
    throw new Error('office_screenshot_unavailable')
  const controller = new AbortController()
  const signal = controller.signal
  const abort = () => controller.abort()
  outerSignal?.addEventListener('abort', abort, { once: true })
  let preview: AgentImage | undefined
  try {
    await bounded(
      validateSkillPackageImage('png', bytes, async (source, mime) => {
        if (signal?.aborted || typeof createImageBitmap !== 'function')
          throw new Error('office_screenshot_unavailable')
        const bitmap = await bounded(
          createImageBitmap(new Blob([new Uint8Array(source).buffer], { type: mime })),
          signal,
          (lateBitmap) => lateBitmap.close(),
        )
        try {
          if (signal?.aborted) throw new Error('tool_cancelled')
          if (bytes.byteLength <= OFFICE_SCREENSHOT_PREVIEW_BYTES) preview = image
          else {
            if (typeof document === 'undefined') throw new Error('office_screenshot_unavailable')
            const canvas = document.createElement('canvas')
            try {
              // Three fixed attempts; never reduce below a 960px long edge to force a pass.
              for (const [edge, quality] of [
                [1600, 0.82],
                [1280, 0.72],
                [960, 0.65],
              ]) {
                if (signal?.aborted) throw new Error('tool_cancelled')
                const scale = Math.min(1, edge! / Math.max(width, height))
                canvas.width = Math.max(1, Math.round(width * scale))
                canvas.height = Math.max(1, Math.round(height * scale))
                const context = canvas.getContext('2d')
                if (!context) throw new Error('office_screenshot_unavailable')
                context.fillStyle = '#FFFFFF'
                context.fillRect(0, 0, canvas.width, canvas.height)
                context.drawImage(bitmap as CanvasImageSource, 0, 0, canvas.width, canvas.height)
                const blob = await bounded(
                  new Promise<Blob | null>((resolve) =>
                    canvas.toBlob(resolve, 'image/jpeg', quality),
                  ),
                  signal,
                )
                if (
                  !blob ||
                  blob.type !== 'image/jpeg' ||
                  blob.size > OFFICE_SCREENSHOT_PREVIEW_BYTES
                )
                  continue
                const encoded = new Uint8Array(await bounded(blob.arrayBuffer(), signal))
                let binary = ''
                for (let offset = 0; offset < encoded.length; offset += 32 * 1024)
                  binary += String.fromCharCode(...encoded.subarray(offset, offset + 32 * 1024))
                preview = { mime: 'image/jpeg', base64: btoa(binary) }
                officeScreenshotBytes(preview, OFFICE_SCREENSHOT_PREVIEW_BYTES)
                break
              }
            } finally {
              canvas.width = canvas.height = 0
            }
          }
          if (!preview) throw new Error('office_screenshot_unavailable')
          return bitmap
        } catch (error) {
          bitmap.close()
          throw error
        }
      }),
      signal,
    )
    if (signal?.aborted) throw new Error('tool_cancelled')
    if (!preview) throw new Error('office_screenshot_unavailable')
    return preview
  } catch {
    throw new Error(outerSignal?.aborted ? 'tool_cancelled' : 'office_screenshot_unavailable')
  } finally {
    controller.abort()
    outerSignal?.removeEventListener('abort', abort)
  }
}
