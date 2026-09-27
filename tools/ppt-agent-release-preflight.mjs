import { readFile, readdir } from 'node:fs/promises'
import { createHash } from 'node:crypto'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const defaultDist = resolve(dirname(fileURLToPath(import.meta.url)), '../apps/office-addin/dist')

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex')
}

function checkManifest(manifest, expectedOrigin) {
  const domains = [...manifest.matchAll(/<AppDomain>([^<]+)<\/AppDomain>/g)].map(
    (match) => match[1],
  )
  if (domains.length !== 1 || domains[0] !== expectedOrigin)
    throw new Error('manifest origin mismatch')
  for (const tag of ['SourceLocation', 'IconUrl', 'HighResolutionIconUrl', 'FunctionFile']) {
    const elements = [...manifest.matchAll(new RegExp(`<${tag}\\b[^>]*>`, 'g'))]
    if (tag === 'SourceLocation' && elements.length !== 1)
      throw new Error('manifest SourceLocation mismatch')
    for (const element of elements) {
      const value = element[0].match(/\bDefaultValue="([^"]+)"/)?.[1]
      if (!value) throw new Error(`manifest ${tag} missing URL`)
      let url
      try {
        url = new URL(value)
      } catch {
        throw new Error(`manifest ${tag} invalid URL`)
      }
      if (
        url.origin !== expectedOrigin ||
        (tag === 'SourceLocation' && url.pathname !== '/taskpane.html') ||
        (tag === 'IconUrl' && url.pathname !== '/assets/icon.png')
      )
        throw new Error(`manifest ${tag} origin mismatch`)
    }
  }
}

export async function inspectOfficeBuild(dist, expectedOrigin) {
  const origin = new URL(expectedOrigin)
  if (origin.protocol !== 'https:' || origin.origin !== expectedOrigin)
    throw new Error('expected origin must be an HTTPS origin without a path')
  const files = await readdir(dist, { recursive: true })
  const metadata = JSON.parse(await readFile(resolve(dist, 'version.json'), 'utf8'))
  if (!/^[A-Za-z0-9_.-]{3,96}$/.test(metadata.buildId || ''))
    throw new Error('invalid version.json buildId')
  const html = await readFile(resolve(dist, 'taskpane.html'), 'utf8')
  const manifest = await readFile(resolve(dist, 'manifest.xml'), 'utf8')
  checkManifest(manifest, expectedOrigin)
  if (html.includes('__WISWORK_CONNECT_ORIGINS__')) throw new Error('unresolved connect policy')
  const scripts = [...html.matchAll(/<script\b[^>]*\bsrc="(\/assets\/[^"?]+\.js)"/g)].map((match) =>
    match[1].slice(1),
  )
  if (scripts.length !== 1 || !/^assets\/taskpane-[A-Za-z0-9_-]+\.js$/.test(scripts[0]))
    throw new Error('missing hashed taskpane entry')
  const script = await readFile(resolve(dist, scripts[0]))
  if (!script.toString('utf8').includes(metadata.buildId))
    throw new Error('buildId differs from compiled taskpane')
  for (const path of files) {
    if (path.endsWith('.map')) throw new Error('source map in release artifact')
  }
  return { buildId: metadata.buildId, script: scripts[0], scriptSha256: sha256(script) }
}

export async function inspectRelayHealth(relayOrigin, fetcher = fetch) {
  const origin = new URL(relayOrigin)
  if (!['http:', 'https:'].includes(origin.protocol) || origin.origin !== relayOrigin)
    throw new Error('relay origin must be an HTTP(S) origin without a path')
  if (
    origin.protocol !== 'https:' &&
    !['localhost', '127.0.0.1', '[::1]'].includes(origin.hostname)
  )
    throw new Error('remote relay requires HTTPS')
  const response = await fetcher(new URL('/office-relay/health', origin), {
    cache: 'no-store',
    signal: AbortSignal.timeout(5_000),
  })
  if (!response.ok || (await response.text()).trim() !== 'ok')
    throw new Error(`relay health failed: ${response.status}`)
}

export async function inspectDeployedOffice(origin, build, fetcher = fetch) {
  const requested = [
    new URL('/version.json', origin),
    new URL('/taskpane.html', origin),
    new URL(`/${build.script}`, origin),
  ]
  const responses = []
  for (const url of requested) {
    const response = await fetcher(url, { cache: 'no-store', signal: AbortSignal.timeout(5_000) })
    if (!response.ok)
      throw new Error(`deployed asset unavailable: ${url.pathname} (${response.status})`)
    responses.push(Buffer.from(await response.arrayBuffer()))
  }
  let metadata
  try {
    metadata = JSON.parse(responses[0].toString('utf8'))
  } catch {
    throw new Error('invalid deployed version.json')
  }
  if (
    metadata.buildId !== build.buildId ||
    !responses[1].toString('utf8').includes(`src="/${build.script}"`) ||
    sha256(responses[2]) !== build.scriptSha256
  )
    throw new Error('deployed Office assets differ from release artifact')
}

async function main(args) {
  const options = Object.fromEntries(
    args.flatMap((arg, index) =>
      arg.startsWith('--') && args[index + 1] && !args[index + 1].startsWith('--')
        ? [[arg, args[index + 1]]]
        : [],
    ),
  )
  if (
    args.length !== Object.keys(options).length * 2 ||
    !options['--origin'] ||
    !options['--relay-origin'] ||
    Object.keys(options).some(
      (key) => !['--origin', '--relay-origin', '--dist', '--deployed'].includes(key),
    ) ||
    (options['--deployed'] && options['--deployed'] !== '1')
  )
    throw new Error(
      'usage: node tools/ppt-agent-release-preflight.mjs --origin https://office.example --relay-origin https://relay.example [--dist path] [--deployed 1]',
    )
  const build = await inspectOfficeBuild(
    resolve(options['--dist'] || defaultDist),
    options['--origin'],
  )
  if (options['--deployed']) await inspectDeployedOffice(options['--origin'], build)
  await inspectRelayHealth(options['--relay-origin'])
  process.stdout.write(
    `PPT Agent release preflight passed: build ${build.buildId}, ${build.script}, Relay healthy\n`,
  )
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main(process.argv.slice(2)).catch((error) => {
    process.stderr.write(`PPT Agent release preflight failed: ${error.message}\n`)
    process.exitCode = 1
  })
}
