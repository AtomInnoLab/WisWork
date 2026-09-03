import { createHash } from 'node:crypto'
import {
  existsSync,
  lstatSync,
  readdirSync,
  readFileSync,
  readlinkSync,
  mkdirSync,
  writeFileSync,
  openSync,
  closeSync,
} from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { execFileSync, spawnSync } from 'node:child_process'

const modules = ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'latex', 'shell']
// Shell embeds every editor's main process; markdown imports Docs icons.
export function sourceRoots(name) {
  return (name === 'shell' ? modules : name === 'markdown' ? ['markdown', 'docs'] : [name]).map(
    (module) => `apps/${module}`,
  )
}
const ignored = new Set([
  'node_modules',
  'out',
  'dist',
  'target',
  '.git',
  'release',
  'release-dogfood',
  'release-preview',
])
export function fingerprint(root, paths) {
  const hash = createHash('sha256')
  function visit(rel) {
    const full = join(root, rel)
    hash.update(JSON.stringify(rel))
    if (!existsSync(full)) {
      hash.update('missing')
      return
    }
    const stat = lstatSync(full)
    hash.update(String(stat.mode))
    if (stat.isSymbolicLink()) hash.update(readlinkSync(full))
    else if (stat.isDirectory()) {
      for (const name of readdirSync(full).sort()) if (!ignored.has(name)) visit(join(rel, name))
    } else hash.update(readFileSync(full))
  }
  paths.forEach(visit)
  return hash.digest('hex')
}
export function selectBuilds(current, previous) {
  return Object.keys(current).filter(
    (name) =>
      !current[name].output ||
      current[name].source !== previous[name]?.source ||
      current[name].output !== previous[name]?.output,
  )
}
export function parseArgs(args) {
  for (const arg of args)
    if (!['--dry-run', '--no-launch'].includes(arg)) throw new Error(`Unknown argument: ${arg}`)
  return { dryRun: args.includes('--dry-run'), launch: !args.includes('--no-launch') }
}
function main() {
  const started = Date.now()
  const options = parseArgs(process.argv.slice(2))
  const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
  const cache = join(root, '.cache/dogfood')
  const output = join(root, 'apps/shell/release-dogfood')
  const builtAt = new Date().toISOString()
  const commit = execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim()
  const version = JSON.parse(readFileSync(join(root, 'apps/shell/package.json'), 'utf8')).version
  const dirty =
    execFileSync('git', ['status', '--porcelain'], { cwd: root, encoding: 'utf8' }).trim().length >
    0
  const log = join(cache, `build-${builtAt.replaceAll(':', '-')}.log`)
  // Hash every shared package and root configuration conservatively; app-local changes
  // invalidate only that app. Output hashes also detect dev builds overwriting the cache.
  const shared = fingerprint(root, [
    'packages',
    'package.json',
    'package-lock.json',
    ...readdirSync(root).filter((n) => /^(tsconfig|vite|\.npmrc)/.test(n)),
  ])
  const snapshot = () =>
    Object.fromEntries(
      modules.map((name) => [
        name,
        {
          source: createHash('sha256')
            .update(shared + process.arch + process.version + fingerprint(root, sourceRoots(name)))
            .digest('hex'),
          output: existsSync(join(root, `apps/${name}/out`))
            ? fingerprint(root, [`apps/${name}/out`])
            : null,
        },
      ]),
    )
  let previous = {}
  try {
    previous = JSON.parse(readFileSync(join(cache, 'fingerprints.json'), 'utf8'))
  } catch {
    /* cold build */
  }
  const builds = selectBuilds(snapshot(), previous)
  console.log(
    JSON.stringify(
      {
        mode: 'dogfood',
        unsigned: true,
        dirty,
        commit,
        version,
        builtAt,
        arch: process.arch,
        builds,
        output,
        log,
        userData: 'WisWork Dogfood',
        enhancedCache:
          'WisWork Dogfood/components/enhanced-mode (verified by component manager; never copied)',
      },
      null,
      2,
    ),
  )
  if (options.dryRun) return
  if (process.platform !== 'darwin') throw new Error('Dogfood packaging requires macOS')
  mkdirSync(cache, { recursive: true })
  const env = {
    ...process.env,
    CSC_IDENTITY_AUTO_DISCOVERY: 'false',
    WISWORK_UNSIGNED_MAC_BUILD: '1',
    WISWORK_ITERATION_COMMIT: commit,
    WISWORK_ITERATION_BUILT_AT: builtAt,
  }
  for (const key of Object.keys(env))
    if (/^(CSC_|APPLE_|WISWORK_UPDATE_|WISWORK_MAC_X64|WISWORK_SLIDES_ACCEPTANCE_E2E)/.test(key))
      delete env[key]
  env.CSC_IDENTITY_AUTO_DISCOVERY = 'false'
  const fd = openSync(log, 'a')
  function run(cmd, args, cwd = root) {
    console.log(`Running ${cmd} ${args.join(' ')} (log: ${log})`)
    const result = spawnSync(cmd, args, { cwd, env, stdio: ['ignore', fd, fd] })
    if (result.error || result.status !== 0)
      throw new Error(`${cmd} failed; see ${log}`, { cause: result.error })
  }
  try {
    run('npm', ['run', 'notices'])
    for (const name of builds) run('npm', ['run', 'build', '-w', `@wiswork/${name}`])
    // Native sidecar must be host-native even when JS outputs are reused.
    run('npm', ['run', 'native:build', '-w', '@wiswork/sheets'])
    run(process.execPath, [
      'tools/fetch-tectonic.mjs',
      '--platform',
      `darwin-${process.arch}`,
      '--output',
      join(root, 'apps/latex/native/tectonic'),
    ])
    run(process.execPath, ['tools/optional-runtime-policy.mjs', '--mode', 'source'])
    run(
      join(root, 'node_modules/.bin/electron-builder'),
      [
        '--config',
        'electron-builder.dogfood.cjs',
        '--mac',
        '--dir',
        `--${process.arch}`,
        '--publish',
        'never',
      ],
      join(root, 'apps/shell'),
    )
    run(process.execPath, [
      'tools/optional-runtime-policy.mjs',
      '--mode',
      'post-package',
      '--artifact-dir',
      output,
    ])
    writeFileSync(join(cache, 'fingerprints.json'), JSON.stringify(snapshot(), null, 2))
    writeFileSync(
      join(cache, 'last-build.json'),
      JSON.stringify({ commit, version, builtAt, output, log }, null, 2),
    )
    const app = join(output, process.arch === 'arm64' ? 'mac-arm64' : 'mac', 'WisWork Dogfood.app')
    if (!existsSync(app)) throw new Error(`Missing artifact: ${app}`)
    console.log(
      `Artifact: ${app}\nLog: ${log}\nElapsed: ${((Date.now() - started) / 1000).toFixed(1)}s`,
    )
    if (options.launch) run('open', ['-n', app])
  } finally {
    closeSync(fd)
  }
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    main()
  } catch (error) {
    console.error(error.message)
    process.exitCode = 1
  }
}
