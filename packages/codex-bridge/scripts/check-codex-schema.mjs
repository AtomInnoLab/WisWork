import { spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { mkdir, mkdtemp, readFile, rm } from 'node:fs/promises'
import { isAbsolute, join } from 'node:path'
import { tmpdir } from 'node:os'
import { fileURLToPath } from 'node:url'

const MAX_OUTPUT_BYTES = 65_536
const TIMEOUT_MS = 10_000
const packageRoot = fileURLToPath(new URL('..', import.meta.url))
const manifestPath = join(packageRoot, 'src/generated/schema-manifest.json')
const bindingsPath = join(packageRoot, 'src/generated/codex-app-server-0.147.ts')

function fail(code) {
  process.stderr.write(`${code}\n`)
  process.exitCode = 1
}

function minimalEnv(codexHome) {
  const env = {
    CODEX_HOME: codexHome,
    NO_PROXY: '127.0.0.1,localhost',
    no_proxy: '127.0.0.1,localhost',
  }
  if (process.platform === 'win32') {
    for (const key of ['SystemRoot', 'WINDIR', 'COMSPEC', 'PATHEXT', 'TEMP', 'TMP']) {
      if (process.env[key] !== undefined) env[key] = process.env[key]
    }
  }
  return env
}

function run(executable, args, cwd, env, label) {
  const result = spawnSync(executable, args, {
    cwd,
    env,
    encoding: 'utf8',
    timeout: TIMEOUT_MS,
    maxBuffer: MAX_OUTPUT_BYTES,
    windowsHide: true,
  })
  if (result.error?.code === 'ETIMEDOUT') throw new Error(`codex_${label}_timeout`)
  if (result.error?.code === 'ENOBUFS') throw new Error(`codex_${label}_output_limit`)
  if (result.error || result.status !== 0) throw new Error(`codex_${label}_command_failed`)
  return result.stdout.trim()
}

function sha256(buffer) {
  return createHash('sha256').update(buffer).digest('hex')
}

function notificationMethods(schema) {
  if (!Array.isArray(schema.oneOf)) throw new Error('codex_schema_invalid')
  const methods = schema.oneOf.map((entry) => entry?.properties?.method?.enum?.[0])
  if (!methods.every((method) => typeof method === 'string')) {
    throw new Error('codex_schema_invalid')
  }
  return [...new Set(methods)].sort()
}

function boundNotificationMethods(source) {
  const block =
    /KNOWN_SERVER_NOTIFICATION_METHODS = Object\.freeze\(\[([\s\S]*?)\] as const\)/.exec(
      source,
    )?.[1]
  if (block === undefined) throw new Error('codex_bindings_invalid')
  return [...block.matchAll(/^\s*'([^']+)',$/gm)].map((match) => match[1]).sort()
}

async function main() {
  const executable = process.argv[2]
  if (typeof executable !== 'string' || !isAbsolute(executable)) {
    throw new Error('codex_executable_must_be_absolute')
  }
  const manifest = JSON.parse(await readFile(manifestPath, 'utf8'))
  const temporaryRoot = await mkdtemp(join(tmpdir(), 'wiswork-codex-schema-'))
  const codexHome = join(temporaryRoot, 'home')
  const output = join(temporaryRoot, 'schema')
  const typeOutput = join(temporaryRoot, 'typescript')
  try {
    await mkdir(codexHome)
    const env = minimalEnv(codexHome)
    const version = run(executable, ['--version'], temporaryRoot, env, 'version')
    if (version !== manifest.version) throw new Error('codex_version_mismatch')
    run(
      executable,
      ['app-server', 'generate-json-schema', '--out', output],
      temporaryRoot,
      env,
      'schema',
    )
    run(executable, ['app-server', 'generate-ts', '--out', typeOutput], temporaryRoot, env, 'types')
    for (const [relativePath, expected] of Object.entries(manifest.sha256)) {
      const actual = sha256(await readFile(join(output, relativePath)))
      if (actual !== expected) throw new Error('codex_schema_drift')
    }
    for (const [relativePath, expected] of Object.entries(manifest.generatedTypeSha256)) {
      const actual = sha256(await readFile(join(typeOutput, relativePath)))
      if (actual !== expected) throw new Error('codex_generated_type_drift')
    }
    const notificationSchema = JSON.parse(
      await readFile(join(output, 'ServerNotification.json'), 'utf8'),
    )
    const bindings = await readFile(bindingsPath, 'utf8')
    if (sha256(bindings) !== manifest.bindingsSha256) {
      throw new Error('codex_binding_drift')
    }
    if (
      JSON.stringify(notificationMethods(notificationSchema)) !==
      JSON.stringify(boundNotificationMethods(bindings))
    ) {
      throw new Error('codex_binding_drift')
    }
    for (const expected of Object.values(manifest.sha256)) {
      if (!bindings.includes(expected)) throw new Error('codex_binding_metadata_drift')
    }
    process.stdout.write('codex_schema_ok\n')
  } finally {
    await rm(temporaryRoot, { recursive: true, force: true })
  }
}

main().catch((error) => fail(error instanceof Error ? error.message : 'codex_schema_check_failed'))
