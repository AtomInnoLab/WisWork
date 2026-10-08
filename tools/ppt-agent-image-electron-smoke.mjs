import { spawn } from 'node:child_process'
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { build } from 'esbuild'
import electron from 'electron'
import pngjs from 'pngjs'

const root = resolve(fileURLToPath(new URL('..', import.meta.url)))
const temp = await mkdtemp(join(tmpdir(), 'ppt-image-electron-'))
try {
  await mkdir(join(temp, 'pc-data'))
  const releaseFixture = join(temp, 'release-smoke.png')
  const releasePng = new pngjs.PNG({ width: 1, height: 1 })
  releasePng.data.set(Buffer.from([11, 22, 33, 255]))
  await writeFile(releaseFixture, pngjs.PNG.sync.write(releasePng))
  const bundle = join(temp, 'presentation-image.cjs')
  await build({
    entryPoints: [join(root, 'apps/shell/src/main/presentation-image.ts')],
    outfile: bundle,
    bundle: true,
    platform: 'node',
    format: 'cjs',
    external: ['electron'],
  })
  const attachmentsBundle = join(temp, 'presentation-attachments.cjs')
  await build({
    entryPoints: [join(root, 'apps/shell/src/main/presentation-attachments.ts')],
    outfile: attachmentsBundle,
    bundle: true,
    platform: 'node',
    format: 'cjs',
    external: ['electron'],
  })
  const fixture = join(root, 'apps/shell/tests/fixtures/presentation-image')
  const driver = join(temp, 'driver.cjs')
  await writeFile(
    driver,
    `
const { app, BrowserWindow, nativeImage } = require('electron')
const { readFileSync } = require('node:fs')
const { createHash } = require('node:crypto')
const { join } = require('node:path')
const { normalizePresentationImage, normalizePresentationImageFirstFrame } = require(${JSON.stringify(bundle)})
const { createPresentationAttachmentService } = require(${JSON.stringify(attachmentsBundle)})
const fixture = ${JSON.stringify(fixture)}
const releaseFixture = ${JSON.stringify(releaseFixture)}
const userDataPath = ${JSON.stringify(join(temp, 'pc-data'))}
app.whenReady().then(async () => {
  const names = ['static.gif', 'static.webp', 'extended.webp', 'static.gif', 'static.webp']
  const results = await Promise.all(names.map(name => normalizePresentationImage(readFileSync(join(fixture, name)))))
  for (const image of results) {
    if (image.width !== 2 || image.height !== 3 || image.bytes.length === 0) throw Error('bad normalized image')
  }
  const release = await normalizePresentationImage(readFileSync(releaseFixture))
  if (release.width !== 1 || release.height !== 1 || !release.bytes.length) throw Error('release fixture normalization failed')
  const releasePixel = nativeImage.createFromBuffer(Buffer.from(release.bytes)).toBitmap()
  if (releasePixel[0] !== 33 || releasePixel[1] !== 22 || releasePixel[2] !== 11 || releasePixel[3] !== 255)
    throw Error('release fixture pixel changed')
  const attachment = createPresentationAttachmentService({ userDataPath })
  const raw = readFileSync(releaseFixture)
  const attachmentId = createHash('sha256').update(raw).digest('hex')
  const signal = new AbortController().signal
  const call = (operation, extra = {}) => attachment({ operation, documentId: 'electron-smoke-doc', attachmentId, ...extra }, signal)
  const begun = await call('attachment_begin', { name: 'release-smoke.png', sizeBytes: raw.length, sha256: attachmentId })
  if (begun.status !== 'uploading' || begun.receivedBytes !== 0) throw Error('Electron attachment begin failed')
  const chunk = await call('attachment_chunk', { offset: 0, base64: raw.toString('base64') })
  if (chunk.receivedBytes !== raw.length) throw Error('Electron attachment chunk failed')
  const finished = await call('attachment_finish')
  if (finished.status !== 'ready' || finished.kind !== 'image' || finished.width !== 1 || finished.height !== 1)
    throw Error('Electron attachment image parse failed')
  const asset = await call('attachment_asset')
  if (asset.id !== attachmentId || createHash('sha256').update(Buffer.from(asset.base64, 'base64')).digest('hex') !== finished.assetSha256)
    throw Error('Electron attachment asset digest mismatch')
  const deleted = await call('attachment_delete')
  if (!deleted.deleted || (await attachment({ operation: 'attachment_list_assets', documentId: 'electron-smoke-doc' }, signal)).attachments.length !== 0)
    throw Error('Electron attachment cleanup failed')
  for (const name of ['animated.gif', 'animated.webp', 'animated.png']) {
    try { await normalizePresentationImage(readFileSync(join(fixture, name))); throw Error('animation accepted') }
    catch (error) { if (error.message !== 'animated_image_unsupported') throw error }
    const source = readFileSync(join(fixture, name))
    const first = await normalizePresentationImageFirstFrame(source)
    const repeated = await normalizePresentationImageFirstFrame(source)
    if (first.width !== 2 || first.height !== 3 || !Buffer.from(first.bytes).equals(Buffer.from(repeated.bytes)))
      throw Error('first frame was not stable')
    const pixel = nativeImage.createFromBuffer(Buffer.from(first.bytes)).toBitmap()
    if (pixel[2] < 240 || pixel[0] > 20) throw Error('first frame was not the red frame')
  }
  const main = new BrowserWindow({ show: false })
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(Error('decoder prevented window-all-closed')), 1000)
    app.once('window-all-closed', () => { clearTimeout(timer); resolve() })
    main.close()
  })
  console.log('Electron image smoke passed: 5 static decodes, release-upload PNG and PC attachment round trip, 3 animations rejected by default and converted to stable first frames, app lifecycle preserved')
  app.quit()
}).catch(error => { console.error(error); app.exit(1) })
`,
  )
  const command = process.platform === 'linux' ? 'xvfb-run' : electron
  const args = process.platform === 'linux' ? ['-a', electron, '--no-sandbox', driver] : [driver]
  await new Promise((done, reject) => {
    const child = spawn(command, args, { stdio: 'inherit' })
    const timer = setTimeout(() => {
      child.kill('SIGKILL')
      reject(new Error('Electron image smoke timed out'))
    }, 20_000)
    child.once('error', (error) => {
      clearTimeout(timer)
      reject(error)
    })
    child.once('exit', (code) => {
      clearTimeout(timer)
      if (code === 0) done()
      else reject(new Error(`Electron image smoke exited ${code}`))
    })
  })
} finally {
  await rm(temp, { recursive: true, force: true })
}
