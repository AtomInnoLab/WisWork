import { createHash, randomBytes } from 'node:crypto'
import { createReadStream } from 'node:fs'
import { lstat, open, rename, rm } from 'node:fs/promises'
import { basename, resolve, win32 } from 'node:path'
import { pathToFileURL } from 'node:url'
import { dump } from 'js-yaml'

function fail(message) {
  throw new Error(message)
}

export function parseArgs(argv) {
  const values = new Map()
  for (let index = 0; index < argv.length; index += 2) {
    const option = argv[index]
    const value = argv[index + 1]
    if (!['--artifact', '--version', '--output'].includes(option) || !value || values.has(option)) {
      fail('expected --artifact, --version, and --output exactly once')
    }
    values.set(option, value)
  }
  if (values.size !== 3) fail('expected --artifact, --version, and --output exactly once')
  const version = values.get('--version')
  if (!/^\d+\.\d+\.\d+$/.test(version)) fail('version must be exact semver')
  return Object.freeze({
    artifact: resolve(values.get('--artifact')),
    output: resolve(values.get('--output')),
    version,
  })
}

function safeBasename(path) {
  const name = basename(path)
  if (name !== win32.basename(name) || name === '.' || name === '..' || name.includes('\0')) {
    fail('artifact name is unsafe')
  }
  return name
}

async function sha512(path) {
  const hash = createHash('sha512')
  for await (const chunk of createReadStream(path)) hash.update(chunk)
  return hash.digest('base64')
}

export async function generateUpdateMetadata({ artifact, output, version }) {
  const info = await lstat(artifact).catch(() => null)
  if (!info?.isFile() || info.isSymbolicLink()) fail('artifact must be a regular file')
  if (!Number.isSafeInteger(info.size) || info.size <= 0) fail('artifact size is invalid')
  const name = safeBasename(artifact)
  if (!name.endsWith('.zip') && !name.endsWith('.exe')) fail('artifact type is unsupported')
  const digest = await sha512(artifact)
  const document = {
    version,
    files: [{ url: name, sha512: digest, size: info.size }],
    path: name,
    sha512: digest,
    releaseDate: new Date().toISOString(),
  }
  const temporary = `${output}.${randomBytes(6).toString('hex')}.tmp`
  const handle = await open(temporary, 'wx')
  try {
    await handle.writeFile(dump(document, { lineWidth: -1 }), 'utf8')
    await handle.sync()
  } finally {
    await handle.close()
  }
  try {
    await rename(temporary, output)
  } catch (error) {
    await rm(temporary, { force: true })
    throw error
  }
  return document
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) {
  generateUpdateMetadata(parseArgs(process.argv.slice(2))).catch(() => {
    process.stderr.write(`${JSON.stringify({ code: 'UPDATE_METADATA_GENERATION_FAILED' })}\n`)
    process.exitCode = 1
  })
}
