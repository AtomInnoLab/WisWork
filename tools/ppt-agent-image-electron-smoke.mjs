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
const { app } = require('electron')
const { readFileSync } = require('node:fs')
const { join } = require('node:path')
const { normalizePresentationImage } = require(${JSON.stringify(bundle)})
const fixture = ${JSON.stringify(fixture)}
app.whenReady().then(async () => {
  const names = ['static.gif', 'static.webp', 'extended.webp', 'static.gif', 'static.webp']
  const results = await Promise.all(names.map(name => normalizePresentationImage(readFileSync(join(fixture, name)))))
  for (const image of results) {
    if (image.width !== 2 || image.height !== 3 || image.bytes.length === 0) throw Error('bad normalized image')
  }
  for (const name of ['animated.gif', 'animated.webp']) {
    try { await normalizePresentationImage(readFileSync(join(fixture, name))); throw Error('animation accepted') }
    catch (error) { if (error.message !== 'parse_failed') throw error }
  }
  console.log('Electron image smoke passed: 5 consecutive/concurrent static decodes, 2 animations rejected')
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
