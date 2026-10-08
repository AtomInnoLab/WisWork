import {
  MAX_IMAGE_IMPORT_BYTES,
  validateBoundedImageBytes,
  type BoundedImage,
} from '../shared/import-media.js'

export type PowerPointImageFit = 'cover' | 'contain'

/** ImageCoercion stretches both dimensions; prepare matching pixels before the Office write. */
export async function preparePowerPointImage(
  image: BoundedImage,
  geometry: { width: number; height: number },
  fit: PowerPointImageFit,
  signal?: AbortSignal,
): Promise<BoundedImage> {
  if (signal?.aborted) throw new Error('cancelled')
  if (Math.abs(image.width * geometry.height - image.height * geometry.width) < 0.000001)
    return image
  if (typeof document === 'undefined' || typeof createImageBitmap !== 'function')
    throw new Error('office_api_unsupported')
  let failure: Error | undefined
  let rejectStopped: (error: Error) => void
  const stopped = new Promise<never>((_resolve, reject) => {
    rejectStopped = reject
  })
  const stop = (code: string) => {
    failure = new Error(code)
    rejectStopped(failure)
  }
  const check = () => {
    if (failure) throw failure
  }
  const abort = () => stop('cancelled')
  signal?.addEventListener('abort', abort, { once: true })
  const timer = setTimeout(() => stop('image_fetch_unavailable'), 10_000)
  let release: (() => void) | undefined
  const prepare = async () => {
    const bytes = Uint8Array.from(atob(image.base64), (value) => value.charCodeAt(0))
    const bitmap = await createImageBitmap(new Blob([bytes.buffer], { type: image.mime }))
    let canvas: HTMLCanvasElement | undefined
    let released = false
    release = () => {
      if (released) return
      released = true
      bitmap.close()
      if (canvas) canvas.width = canvas.height = 0
    }
    try {
      check()
      if (bitmap.width !== image.width || bitmap.height !== image.height)
        throw new Error('invalid_image')
      canvas = document.createElement('canvas')
      const sourceScale = (fit === 'cover' ? Math.min : Math.max)(
        image.width / geometry.width,
        image.height / geometry.height,
      )
      const scale = Math.min(2, 1536 / Math.max(geometry.width, geometry.height), sourceScale)
      canvas.width = Math.max(1, Math.round(geometry.width * scale))
      canvas.height = Math.max(1, Math.round(geometry.height * scale))
      const context = canvas.getContext('2d')
      if (!context) throw new Error('office_api_unsupported')
      if (fit === 'cover') {
        const cropScale = Math.min(image.width / canvas.width, image.height / canvas.height)
        const width = canvas.width * cropScale
        const height = canvas.height * cropScale
        context.drawImage(
          bitmap,
          (image.width - width) / 2,
          (image.height - height) / 2,
          width,
          height,
          0,
          0,
          canvas.width,
          canvas.height,
        )
      } else {
        const fitScale = Math.min(canvas.width / image.width, canvas.height / image.height)
        const width = image.width * fitScale
        const height = image.height * fitScale
        // A fresh PNG canvas is transparent; letterboxing preserves the slide background.
        context.drawImage(
          bitmap,
          0,
          0,
          image.width,
          image.height,
          (canvas.width - width) / 2,
          (canvas.height - height) / 2,
          width,
          height,
        )
      }
      const blob = await new Promise<Blob | null>((resolve) => canvas!.toBlob(resolve, 'image/png'))
      check()
      if (!blob || blob.type !== 'image/png') throw new Error('invalid_image')
      if (blob.size > MAX_IMAGE_IMPORT_BYTES) throw new Error('image_limit')
      const output = await validateBoundedImageBytes(new Uint8Array(await blob.arrayBuffer()))
      check()
      return output
    } finally {
      release?.()
    }
  }
  try {
    return await Promise.race([prepare(), stopped])
  } finally {
    clearTimeout(timer)
    signal?.removeEventListener('abort', abort)
    release?.()
  }
}
