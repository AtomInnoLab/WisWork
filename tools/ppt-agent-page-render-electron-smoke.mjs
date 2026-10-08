import { spawn } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { build } from 'esbuild'
import electron from 'electron'
import PptxGenJS from 'pptxgenjs'

const root = resolve(fileURLToPath(new URL('..', import.meta.url)))
const temp = await mkdtemp(join(tmpdir(), 'ppt-page-render-electron-'))
try {
  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_WIDE'
  pptx.addSlide().addText('WisWork fallback preview', { x: 1, y: 1, w: 8, h: 1 })
  const packagePath = join(temp, 'page.pptx')
  await writeFile(packagePath, await pptx.write({ outputType: 'nodebuffer' }))
  const bundle = join(temp, 'presentation-page-render.cjs')
  await build({
    entryPoints: [join(root, 'apps/shell/src/main/presentation-page-render.ts')],
    outfile: bundle,
    bundle: true,
    platform: 'node',
    format: 'cjs',
    external: ['electron'],
  })
  const driver = join(temp, 'driver.cjs')
  await writeFile(
    driver,
    `
const { app } = require('electron')
const { readFile } = require('node:fs/promises')
const { renderSinglePagePackage } = require(${JSON.stringify(bundle)})
app.whenReady().then(async () => {
  const bytes = await readFile(${JSON.stringify(packagePath)})
  const png = Buffer.from(await renderSinglePagePackage(bytes, new AbortController().signal))
  if (png.length < 33 || png.length > 64 * 1024 || png.subarray(0, 8).toString('hex') !== '89504e470d0a1a0a')
    throw Error('invalid bounded PNG')
  const width = png.readUInt32BE(16), height = png.readUInt32BE(20)
  if (width < 1 || width > 960 || height < 1 || height > 8192) throw Error('invalid PNG dimensions')
  console.log('Electron page fallback smoke passed:', png.length, width, height)
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
      reject(new Error('Electron page fallback smoke timed out'))
    }, 30_000)
    child.once('error', (error) => {
      clearTimeout(timer)
      reject(error)
    })
    child.once('exit', (code) => {
      clearTimeout(timer)
      if (code === 0) done()
      else reject(new Error(`Electron page fallback smoke exited ${code}`))
    })
  })
} finally {
  await rm(temp, { recursive: true, force: true })
}
