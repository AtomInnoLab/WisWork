import { spawn } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { build } from 'esbuild'
import electron from 'electron'

const root = resolve(fileURLToPath(new URL('..', import.meta.url)))
const temp = await mkdtemp(join(tmpdir(), 'ppt-image-electron-'))
try {
  const bundle = join(temp, 'presentation-image.cjs')
  await build({
    entryPoints: [join(root, 'apps/shell/src/main/presentation-image.ts')],
    outfile: bundle,
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
const { join } = require('node:path')
const { normalizePresentationImage, normalizePresentationImageFirstFrame } = require(${JSON.stringify(bundle)})
const fixture = ${JSON.stringify(fixture)}
app.whenReady().then(async () => {
  const names = ['static.gif', 'static.webp', 'extended.webp', 'static.gif', 'static.webp']
  const results = await Promise.all(names.map(name => normalizePresentationImage(readFileSync(join(fixture, name)))))
  for (const image of results) {
    if (image.width !== 2 || image.height !== 3 || image.bytes.length === 0) throw Error('bad normalized image')
  }
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
  console.log('Electron image smoke passed: 5 static decodes, 3 animations rejected by default and converted to stable first frames, app lifecycle preserved')
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
