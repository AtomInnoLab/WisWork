import { readFile, readdir } from 'node:fs/promises'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const defaultDist = resolve(dirname(fileURLToPath(import.meta.url)), '../apps/office-addin/dist')

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
  if (
    !manifest.includes(`${expectedOrigin}/taskpane.html`) ||
    !manifest.includes(`<AppDomain>${expectedOrigin}</AppDomain>`) ||
    manifest.includes('localhost')
  )
    throw new Error('manifest origin mismatch')
  if (html.includes('__WISWORK_CONNECT_ORIGINS__')) throw new Error('unresolved connect policy')
  const scripts = [...html.matchAll(/<script\b[^>]*\bsrc="(\/assets\/[^"?]+\.js)"/g)].map((match) =>
    match[1].slice(1),
  )
  if (scripts.length !== 1 || !/^assets\/taskpane-[A-Za-z0-9_-]+\.js$/.test(scripts[0]))
    throw new Error('missing hashed taskpane entry')
  const script = await readFile(resolve(dist, scripts[0]), 'utf8')
  if (!script.includes(metadata.buildId)) throw new Error('buildId differs from compiled taskpane')
  for (const path of files) {
    if (path.endsWith('.map')) throw new Error('source map in release artifact')
  }
  return { buildId: metadata.buildId, script: scripts[0] }
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
    responses.push(await response.text())
  }
  let metadata
  try {
    metadata = JSON.parse(responses[0])
  } catch {
    throw new Error('invalid deployed version.json')
  }
  if (
    metadata.buildId !== build.buildId ||
    !responses[1].includes(`src="/${build.script}"`) ||
    !responses[2].includes(build.buildId)
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
