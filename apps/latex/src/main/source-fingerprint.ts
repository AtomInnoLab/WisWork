import { createHash } from 'node:crypto'
import { constants } from 'node:fs'
import { lstat, open, readdir, realpath } from 'node:fs/promises'
import { join, relative, sep } from 'node:path'

const MAX_ENTRIES = 10_000
const MAX_FILE_BYTES = 100 * 1024 * 1024
const MAX_TOTAL_BYTES = 500 * 1024 * 1024

export async function fingerprintProjectDirectory(root: string): Promise<string> {
  const canonicalRoot = await realpath(root)
  const hash = createHash('sha256')
  let entries = 0
  let totalBytes = 0

  const walk = async (directory: string): Promise<void> => {
    const before = await lstat(directory)
    if (!before.isDirectory() || before.isSymbolicLink())
      throw new Error('Unsafe project directory')
    const canonicalDirectory = await realpath(directory)
    if (
      canonicalDirectory !== canonicalRoot &&
      !canonicalDirectory.startsWith(`${canonicalRoot}${sep}`)
    ) {
      throw new Error('Project directory escaped its root')
    }
    const children = (await readdir(directory, { withFileTypes: true })).sort((a, b) =>
      a.name.localeCompare(b.name),
    )
    for (const child of children) {
      entries += 1
      if (entries > MAX_ENTRIES) throw new Error('Project entry limit exceeded')
      const path = join(directory, child.name)
      const portablePath = relative(canonicalRoot, path).split(sep).join('/')
      const stats = await lstat(path)
      if (stats.isSymbolicLink()) throw new Error('Symbolic links are not allowed')
      if (stats.isDirectory()) {
        hash.update(`d\0${portablePath}\0`)
        await walk(path)
        continue
      }
      if (!stats.isFile()) throw new Error('Only regular files and directories are allowed')
      if (stats.size > MAX_FILE_BYTES) throw new Error('Project file limit exceeded')
      totalBytes += stats.size
      if (totalBytes > MAX_TOTAL_BYTES) throw new Error('Project size limit exceeded')

      hash.update(`f\0${portablePath}\0${stats.size}\0`)
      const handle = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW)
      try {
        const opened = await handle.stat()
        if (!sameFile(stats, opened)) throw new Error('Project file changed before read')
        const buffer = Buffer.allocUnsafe(64 * 1024)
        while (true) {
          const { bytesRead } = await handle.read(buffer, 0, buffer.length, null)
          if (!bytesRead) break
          hash.update(buffer.subarray(0, bytesRead))
        }
        if (!sameFile(opened, await handle.stat()))
          throw new Error('Project file changed during read')
      } finally {
        await handle.close()
      }
    }
    const after = await lstat(directory)
    if (!sameFile(before, after)) throw new Error('Project directory changed during read')
  }

  await walk(canonicalRoot)
  return hash.digest('hex')
}

function sameFile(
  left: Awaited<ReturnType<typeof lstat>>,
  right: Awaited<ReturnType<typeof lstat>>,
) {
  return (
    left.dev === right.dev &&
    left.ino === right.ino &&
    left.size === right.size &&
    left.mtimeMs === right.mtimeMs &&
    left.ctimeMs === right.ctimeMs &&
    left.isDirectory() === right.isDirectory()
  )
}
