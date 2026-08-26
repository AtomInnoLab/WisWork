const encoder = new TextEncoder()
const decoder = new TextDecoder('utf-8', { fatal: true })
export const MAX_VFS_FILE_BYTES = 20 * 1024 * 1024
export const MAX_VFS_TOTAL_BYTES = 64 * 1024 * 1024
export const MAX_VFS_READ_BYTES = 2 * 1024 * 1024

export interface VfsLimits {
  maxFileBytes: number
  maxTotalBytes: number
  maxFiles: number
  maxNameBytes: number
}

const DEFAULT_LIMITS: VfsLimits = {
  maxFileBytes: MAX_VFS_FILE_BYTES,
  maxTotalBytes: MAX_VFS_TOTAL_BYTES,
  maxFiles: 256,
  maxNameBytes: 128,
}

function denied(): never {
  throw new Error('vfs_path_denied')
}
function limited(): never {
  throw new Error('vfs_limit')
}

export class InMemoryVfs {
  readonly #limits: VfsLimits
  readonly #files = new Map<string, Uint8Array>()
  readonly #readOnly = new Set<string>()

  constructor(limits: Partial<VfsLimits> = {}) {
    this.#limits = { ...DEFAULT_LIMITS, ...limits }
  }

  normalize(path: string): string {
    if (typeof path !== 'string' || path.includes('\0') || path.includes('\\')) denied()
    if (
      path !== '/home/user' &&
      path !== '/home/skills' &&
      !path.startsWith('/home/user/') &&
      !path.startsWith('/home/skills/')
    )
      denied()
    const parts: string[] = []
    for (const part of path.split('/')) {
      if (!part || part === '.') continue
      if (part === '..') {
        if (parts.length <= 2) denied()
        parts.pop()
      } else {
        if (encoder.encode(part).byteLength > this.#limits.maxNameBytes) limited()
        parts.push(part)
      }
    }
    const normalized = `/${parts.join('/')}`
    if (
      normalized !== '/home/user' &&
      normalized !== '/home/skills' &&
      !normalized.startsWith('/home/user/') &&
      !normalized.startsWith('/home/skills/')
    )
      denied()
    return normalized
  }

  writeFile(path: string, content: string | Uint8Array): void {
    const normalized = this.normalize(path)
    if (
      normalized === '/home/user' ||
      this.#readOnly.has(normalized) ||
      normalized.startsWith('/home/skills/')
    )
      denied()
    this.#set(normalized, content)
  }

  writeBatch(entries: ReadonlyArray<readonly [string, string | Uint8Array]>): void {
    const prepared = entries.map(([path, content]) => {
      const normalized = this.normalize(path)
      if (normalized === '/home/user' || normalized.startsWith('/home/skills/')) denied()
      const bytes = typeof content === 'string' ? encoder.encode(content) : content.slice()
      if (bytes.byteLength > this.#limits.maxFileBytes) limited()
      return [normalized, bytes] as const
    })
    if (new Set(prepared.map(([path]) => path)).size !== prepared.length) limited()
    if (prepared.some(([path]) => this.#readOnly.has(path))) denied()
    const newFiles = prepared.filter(([path]) => !this.#files.has(path)).length
    if (this.#files.size + newFiles > this.#limits.maxFiles) limited()
    const replaced = prepared.reduce(
      (sum, [path]) => sum + (this.#files.get(path)?.byteLength ?? 0),
      0,
    )
    const added = prepared.reduce((sum, [, bytes]) => sum + bytes.byteLength, 0)
    const total = [...this.#files.values()].reduce((sum, value) => sum + value.byteLength, 0)
    if (total - replaced + added > this.#limits.maxTotalBytes) limited()
    for (const [path, bytes] of prepared) this.#files.set(path, bytes)
  }

  mountReadOnly(path: string, content: string | Uint8Array): void {
    this.mountReadOnlyBatch([[path, content]])
  }

  mountReadOnlyBatch(entries: ReadonlyArray<readonly [string, string | Uint8Array]>): void {
    const prepared = entries.map(([path, content]) => {
      const normalized = this.normalize(path)
      if (!normalized.startsWith('/home/skills/') || this.#readOnly.has(normalized)) denied()
      const bytes = typeof content === 'string' ? encoder.encode(content) : content.slice()
      if (bytes.byteLength > this.#limits.maxFileBytes) limited()
      return [normalized, bytes] as const
    })
    if (new Set(prepared.map(([path]) => path)).size !== prepared.length) limited()
    const newFiles = prepared.filter(([path]) => !this.#files.has(path)).length
    if (this.#files.size + newFiles > this.#limits.maxFiles) limited()
    const replaced = prepared.reduce(
      (sum, [path]) => sum + (this.#files.get(path)?.byteLength ?? 0),
      0,
    )
    const added = prepared.reduce((sum, [, bytes]) => sum + bytes.byteLength, 0)
    const total = [...this.#files.values()].reduce((sum, value) => sum + value.byteLength, 0)
    if (total - replaced + added > this.#limits.maxTotalBytes) limited()
    for (const [path, bytes] of prepared) {
      this.#files.set(path, bytes)
      this.#readOnly.add(path)
    }
  }

  #set(path: string, content: string | Uint8Array): void {
    const bytes = typeof content === 'string' ? encoder.encode(content) : content.slice()
    if (bytes.byteLength > this.#limits.maxFileBytes) limited()
    const previous = this.#files.get(path)?.byteLength ?? 0
    const total = [...this.#files.values()].reduce((sum, value) => sum + value.byteLength, 0)
    if (total - previous + bytes.byteLength > this.#limits.maxTotalBytes) limited()
    if (!this.#files.has(path) && this.#files.size >= this.#limits.maxFiles) limited()
    this.#files.set(path, bytes)
  }

  readBytes(path: string, options: { offset?: number; maxBytes?: number } = {}): Uint8Array {
    const bytes = this.#files.get(this.normalize(path))
    if (!bytes) throw new Error('vfs_not_found')
    const offset = options.offset ?? 0
    const maximum = options.maxBytes ?? this.#limits.maxFileBytes
    if (!Number.isInteger(offset) || offset < 0 || !Number.isInteger(maximum) || maximum < 0)
      limited()
    return bytes.slice(offset, offset + maximum)
  }

  readText(path: string, options: { offset?: number; maxBytes?: number } = {}): string {
    const allBytes = this.readBytes(path)
    const offset = options.offset ?? 0
    const maximum = options.maxBytes ?? this.#limits.maxFileBytes
    if (!Number.isInteger(offset) || offset < 0 || !Number.isInteger(maximum) || maximum < 0)
      limited()
    if (offset < allBytes.byteLength && (allBytes[offset] & 0xc0) === 0x80) {
      throw new Error('vfs_invalid_utf8_boundary')
    }
    const bytes = allBytes.slice(offset, offset + maximum)
    try {
      return decoder.decode(bytes)
    } catch {
      if (offset + maximum < allBytes.byteLength) {
        for (let trim = 1; trim <= Math.min(3, bytes.byteLength); trim += 1) {
          try {
            return decoder.decode(bytes.slice(0, -trim))
          } catch {
            // UTF-8 code points are at most four bytes; keep trimming the incomplete suffix.
          }
        }
      }
      throw new Error('vfs_invalid_utf8')
    }
  }

  list(path: string): string[] {
    const prefix = `${this.normalize(path).replace(/\/$/, '')}/`
    return [...this.#files.keys()].filter((name) => name.startsWith(prefix)).sort()
  }

  unmountReadOnlyTree(path: string): void {
    const root = this.normalize(path)
    if (!root.startsWith('/home/skills/')) denied()
    const prefix = `${root.replace(/\/$/, '')}/`
    for (const name of [...this.#files.keys()]) {
      if (name === root || name.startsWith(prefix)) {
        if (!this.#readOnly.has(name)) denied()
        this.#files.delete(name)
        this.#readOnly.delete(name)
      }
    }
  }

  clear(): void {
    this.#files.clear()
    this.#readOnly.clear()
  }
}
