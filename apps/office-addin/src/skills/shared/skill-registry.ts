import { InMemoryVfs } from './vfs.js'
import type { ParsedSkillArchive } from './skill-package.js'

export interface SkillMetadata {
  name: string
  description: string
  version?: string
}
export interface ParsedSkill {
  metadata: SkillMetadata
  body: string
  source: string
}

export const MAX_SKILL_BYTES = 64 * 1024
const NAME = /^[a-z0-9][a-z0-9-]{0,62}$/

export function parseSkillPackage(source: string): ParsedSkill {
  if (new TextEncoder().encode(source).byteLength > MAX_SKILL_BYTES)
    throw new Error('invalid_skill_package')
  const match = /^---\r?\n([\s\S]*?)\r?\n---\r?\n([\s\S]*)$/.exec(source)
  if (!match) throw new Error('invalid_skill_package')
  const values: Record<string, string> = {}
  for (const line of match[1].split(/\r?\n/)) {
    const field = /^([a-z]+):\s*(.+)$/.exec(line)
    if (
      !field ||
      !['name', 'description', 'version'].includes(field[1]) ||
      Object.hasOwn(values, field[1])
    ) {
      throw new Error('invalid_skill_package')
    }
    values[field[1]] = field[2].trim()
  }
  if (!NAME.test(values.name ?? '') || !values.description || values.description.length > 240) {
    throw new Error('invalid_skill_package')
  }
  if (values.version && !/^\d+\.\d+\.\d+$/.test(values.version))
    throw new Error('invalid_skill_package')
  return {
    metadata: {
      name: values.name,
      description: values.description,
      ...(values.version ? { version: values.version } : {}),
    },
    body: match[2],
    source,
  }
}

export class SkillRegistry {
  readonly #skills = new Map<string, SkillMetadata>()
  constructor(private readonly vfs: InMemoryVfs) {}

  install(source: string, files: Record<string, string> = {}): SkillMetadata {
    const parsed = parseSkillPackage(source)
    if (this.#skills.has(parsed.metadata.name)) throw new Error('skill_already_installed')
    const root = `/home/skills/${parsed.metadata.name}`
    const entries: Array<readonly [string, string]> = [[`${root}/SKILL.md`, source]]
    for (const [path, content] of Object.entries(files)) {
      if (path.startsWith('/') || path.split('/').includes('..'))
        throw new Error('invalid_skill_package')
      entries.push([`${root}/${path}`, content])
    }
    this.vfs.mountReadOnlyBatch(entries)
    this.#skills.set(parsed.metadata.name, Object.freeze({ ...parsed.metadata }))
    return { ...parsed.metadata }
  }

  installPackage(pkg: ParsedSkillArchive): SkillMetadata {
    const metadata = pkg.skill.metadata
    if (this.#skills.has(metadata.name)) throw new Error('skill_already_installed')
    const root = `/home/skills/${metadata.name}`
    this.vfs.mountReadOnlyBatch(pkg.files.map((file) => [`${root}/${file.path}`, file.bytes]))
    this.#skills.set(metadata.name, Object.freeze({ ...metadata }))
    return { ...metadata }
  }

  uninstall(name: string): void {
    if (!this.#skills.has(name)) throw new Error('skill_not_installed')
    this.vfs.unmountReadOnlyTree(`/home/skills/${name}`)
    this.#skills.delete(name)
  }

  prompt(): string {
    return [...this.#skills.values()]
      .sort((a, b) => a.name.localeCompare(b.name))
      .map((skill) => `${skill.name}: ${skill.description} (/home/skills/${skill.name}/SKILL.md)`)
      .join('\n')
  }

  list(): SkillMetadata[] {
    return [...this.#skills.values()]
      .sort((a, b) => a.name.localeCompare(b.name))
      .map((skill) => ({ ...skill }))
  }

  clear(): void {
    for (const name of this.#skills.keys()) this.vfs.unmountReadOnlyTree(`/home/skills/${name}`)
    this.#skills.clear()
  }
}
