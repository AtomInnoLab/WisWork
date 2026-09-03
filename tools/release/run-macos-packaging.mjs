import { spawn } from 'node:child_process'
import { pathToFileURL } from 'node:url'

const MAX_CAPTURED_OUTPUT = 512 * 1024
const MAX_ATTEMPTS = 3
const BUSY_DEVICE =
  /Unable to detach device cleanly: hdiutil: couldn't eject "(disk[0-9]+)" - Resource busy/

function appendBounded(current, chunk) {
  const next = current + chunk
  return next.length <= MAX_CAPTURED_OUTPUT ? next : next.slice(-MAX_CAPTURED_OUTPUT)
}

function runVisible(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { env: process.env, stdio: ['inherit', 'pipe', 'pipe'] })
    let output = ''
    const forward = (destination) => (chunk) => {
      destination.write(chunk)
      output = appendBounded(output, chunk.toString('utf8'))
    }
    child.stdout.on('data', forward(process.stdout))
    child.stderr.on('data', forward(process.stderr))
    child.once('error', reject)
    child.once('close', (code) => resolve({ code: code ?? 1, output }))
  })
}

export async function runMacosPackaging(electronArgs, options = {}) {
  const run = options.run ?? runVisible
  const builderArgs = ['electron-builder', '--config', 'electron-builder.cjs', ...electronArgs]

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt += 1) {
    const result = await run('npx', builderArgs)
    if (result.code === 0) return 0

    const match = BUSY_DEVICE.exec(result.output)
    if (!match || attempt === MAX_ATTEMPTS) return result.code

    const device = `/dev/${match[1]}`
    process.stderr.write(
      `[desktop-release] transient DMG detach failure on ${device}; cleaning it up before retry ${attempt + 1}/${MAX_ATTEMPTS}\n`,
    )
    const detached = await run('hdiutil', ['detach', '-force', device])
    if (detached.code !== 0) return detached.code
  }

  return 1
}

async function main() {
  if (process.platform !== 'darwin') {
    throw new Error('macOS packaging runner requires darwin')
  }
  process.exitCode = await runMacosPackaging(process.argv.slice(2))
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) {
  await main()
}
