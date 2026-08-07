import { basename } from 'node:path'
import { parse, type TomlTable } from 'smol-toml'
import type { ProjectPathPolicy } from './path-policy.js'
import type { MainFileDiscovery, MainFileDiscoveryOptions } from './types.js'

export interface MainFileProjectReader {
  listTextFiles(): Promise<string[]>
  readText(path: string): Promise<string>
}

async function existingFile(policy: ProjectPathPolicy, path: string | undefined) {
  if (!path) return undefined
  try {
    await policy.resolveExisting(path, 'file')
    return policy.normalize(path)
  } catch {
    return undefined
  }
}

interface ParsedTectonicInputs {
  officialOutputs: string[][]
  legacyInputs: string[]
}

function stringInputs(value: unknown): string[] {
  if (typeof value === 'string') return [value]
  if (Array.isArray(value))
    return value.filter((entry): entry is string => typeof entry === 'string')
  return []
}

function parseTectonicInputs(toml: string): ParsedTectonicInputs {
  const config = parse(toml)
  const output = Array.isArray(config.output) ? config.output : []
  const officialOutputs = output
    .filter((entry): entry is TomlTable => typeof entry === 'object' && entry !== null)
    .map((entry) => stringInputs(entry.inputs))
  return { officialOutputs, legacyInputs: stringInputs(config.inputs) }
}

async function resolveTectonicInputs(
  policy: ProjectPathPolicy,
  paths: string[],
  sourcePrefix: string,
): Promise<string[]> {
  const resolved = await Promise.all(
    paths.map((path) => existingFile(policy, sourcePrefix ? `${sourcePrefix}/${path}` : path)),
  )
  return resolved.filter((path): path is string => path !== undefined && path.endsWith('.tex'))
}

export async function discoverMainFile(
  policy: ProjectPathPolicy,
  options: MainFileDiscoveryOptions = {},
  reader?: MainFileProjectReader,
): Promise<MainFileDiscovery> {
  const savedCandidate = options.savedMainFile?.toLowerCase().endsWith('.tex')
    ? options.savedMainFile
    : undefined
  const saved = await existingFile(policy, savedCandidate)
  if (saved) return { kind: 'found', path: saved, source: 'saved' }

  const read =
    reader?.readText.bind(reader) ??
    (async (path: string) => {
      const { readFile } = await import('node:fs/promises')
      return readFile(await policy.resolveExisting(path, 'file'), 'utf8')
    })

  const tectonicPath = await existingFile(policy, 'Tectonic.toml')
  if (tectonicPath) {
    try {
      const parsed = parseTectonicInputs(await read(tectonicPath))
      const officialSets = await Promise.all(
        parsed.officialOutputs.map((inputs) => resolveTectonicInputs(policy, inputs, 'src')),
      )
      const validOfficialSets = officialSets.filter((candidates) => candidates.length > 0)
      const officialCandidates = [...new Set(validOfficialSets.flat())].sort()
      if (parsed.officialOutputs.length > 1 || officialCandidates.length > 1) {
        return { kind: 'selection-required', candidates: officialCandidates }
      }
      if (officialCandidates.length === 1) {
        return { kind: 'found', path: officialCandidates[0]!, source: 'tectonic' }
      }

      const legacyCandidates = [
        ...new Set(await resolveTectonicInputs(policy, parsed.legacyInputs, '')),
      ].sort()
      if (legacyCandidates.length === 1) {
        return { kind: 'found', path: legacyCandidates[0]!, source: 'tectonic' }
      }
      if (legacyCandidates.length > 1) {
        return { kind: 'selection-required', candidates: legacyCandidates }
      }
    } catch {
      // An invalid or unreadable config does not prevent lower-priority discovery.
    }
  }

  const main = await existingFile(policy, 'main.tex')
  if (main) return { kind: 'found', path: main, source: 'main' }

  const files = reader ? await reader.listTextFiles() : await listTexFiles(policy)
  const candidates: string[] = []
  for (const path of files.filter((path) => path.endsWith('.tex'))) {
    try {
      if (/^[ \t]*\\documentclass(?:\[[^\]]*\])?\s*\{/m.test(await read(path))) {
        candidates.push(path)
      }
    } catch {
      // Invalid, oversized, or concurrently removed files are not candidates.
    }
  }
  candidates.sort()
  if (candidates.length === 1) {
    return { kind: 'found', path: candidates[0]!, source: 'documentclass' }
  }
  if (candidates.length > 1) return { kind: 'selection-required', candidates }
  return { kind: 'not-found', candidates: [] }
}

async function listTexFiles(policy: ProjectPathPolicy): Promise<string[]> {
  const { readdir } = await import('node:fs/promises')
  const results: string[] = []
  async function walk(relativeDir: string): Promise<void> {
    await policy.assertRootUnchanged()
    const absoluteDir = relativeDir
      ? await policy.resolveExisting(relativeDir, 'directory')
      : policy.realRootPath
    const entries = await readdir(absoluteDir, { withFileTypes: true })
    for (const entry of entries) {
      const path = relativeDir ? `${relativeDir}/${entry.name}` : entry.name
      if (entry.isSymbolicLink()) continue
      if (entry.isDirectory()) await walk(path)
      else if (entry.isFile() && basename(path).endsWith('.tex')) results.push(path)
    }
  }
  await walk('')
  return results.sort()
}
