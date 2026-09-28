import { spawn } from 'node:child_process'
import { mkdtemp, open, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'

const MAX_PNG_BYTES = 64 * 1024
const MAX_RENDERED_PNG_BYTES = 4 * 1024 * 1024
const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10])

/** Read no more than the renderer's file budget and reject oversized PNG dimensions before decode. */
export async function readBoundedRenderedPng(path: string): Promise<Uint8Array> {
  const handle = await open(path, 'r')
  try {
    const bytes = Buffer.alloc(MAX_RENDERED_PNG_BYTES + 1)
    let length = 0
    while (length < bytes.length) {
      const part = await handle.read(bytes, length, bytes.length - length, length)
      if (!part.bytesRead) break
      length += part.bytesRead
    }
    if (
      length < 33 ||
      length > MAX_RENDERED_PNG_BYTES ||
      !bytes.subarray(0, 8).equals(PNG_SIGNATURE) ||
      bytes.readUInt32BE(8) !== 13 ||
      bytes.toString('ascii', 12, 16) !== 'IHDR'
    )
      throw new Error('renderer_unavailable')
    const width = bytes.readUInt32BE(16)
    const height = bytes.readUInt32BE(20)
    if (width < 1 || height < 1 || width > 8192 || height > 8192)
      throw new Error('renderer_unavailable')
    return bytes.subarray(0, length)
  } finally {
    await handle.close()
  }
}

/** Optional local fallback for a validated single-slide PPTX; never claims PowerPoint fidelity. */
export async function convertSinglePagePackageToPng(
  pptx: Uint8Array,
  signal: AbortSignal,
): Promise<Uint8Array> {
  if (signal.aborted) throw new Error('aborted')
  if (!pptx.length || pptx.length > 8 * 1024 * 1024) throw new Error('renderer_unavailable')
  const dir = await mkdtemp(join(tmpdir(), 'wiswork-page-render-'))
  try {
    const source = join(dir, 'page.pptx')
    await writeFile(source, pptx, { mode: 0o600 })
    await new Promise<void>((resolve, reject) => {
      const process = spawn(
        'soffice',
        [
          `-env:UserInstallation=${pathToFileURL(join(dir, 'profile')).href}`,
          '--headless',
          '--convert-to',
          'png',
          '--outdir',
          dir,
          source,
        ],
        { stdio: 'ignore' },
      )
      const timeout = setTimeout(() => process.kill('SIGKILL'), 15_000)
      const abort = () => process.kill('SIGKILL')
      signal.addEventListener('abort', abort, { once: true })
      process.once('error', () => reject(new Error('renderer_unavailable')))
      process.once('close', (code) =>
        code === 0 && !signal.aborted
          ? resolve()
          : reject(new Error(signal.aborted ? 'aborted' : 'renderer_unavailable')),
      )
      process.once('close', () => {
        clearTimeout(timeout)
        signal.removeEventListener('abort', abort)
      })
    })
    if (signal.aborted) throw new Error('aborted')
    const raw = await readBoundedRenderedPng(join(dir, 'page.png'))
    return raw
  } catch (error) {
    if (error instanceof Error && error.message === 'aborted') throw error
    throw new Error('renderer_unavailable', { cause: error })
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}

export async function renderSinglePagePackage(
  pptx: Uint8Array,
  signal: AbortSignal,
): Promise<Uint8Array> {
  try {
    const raw = await convertSinglePagePackageToPng(pptx, signal)
    const { nativeImage } = await import('electron')
    const image = nativeImage.createFromBuffer(Buffer.from(raw))
    if (image.isEmpty()) throw new Error('renderer_unavailable')
    const size = image.getSize()
    if (size.width < 1 || size.height < 1 || size.width > 8192 || size.height > 8192)
      throw new Error('renderer_unavailable')
    for (const width of [960, 640, 480, 320, 240]) {
      if (signal.aborted) throw new Error('aborted')
      const png = (size.width > width ? image.resize({ width }) : image).toPNG()
      if (png.length <= MAX_PNG_BYTES) return png
    }
    throw new Error('output_too_large')
  } catch (error) {
    if (error instanceof Error && ['aborted', 'output_too_large'].includes(error.message))
      throw error
    throw new Error('renderer_unavailable', { cause: error })
  }
}
