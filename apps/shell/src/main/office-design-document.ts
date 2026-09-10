export interface OfficeDesignDocument {
  documentId: string
  markdown: string
  revision: string
}

export type OfficeDesignDocumentHandler = (request: {
  sessionId: string
  host: string
  body: unknown
  signal: AbortSignal
}) => Promise<OfficeDesignDocument>

export function createOfficeDesignDocuments(options: {
  directory: string
  openFile(path: string): Promise<void>
}): OfficeDesignDocumentHandler {
  const files = new Map<string, { path: string; baseRevision: string }>()
  return async ({ sessionId, host, body, signal }) => {
    const current = () => {
      if (signal.aborted) throw new Error('cancelled')
    }
    current()
    if (
      host !== 'PowerPoint' ||
      !sessionId ||
      typeof body !== 'object' ||
      !body ||
      Array.isArray(body)
    )
      throw new Error('design_document_invalid')
    const input = body as Record<string, unknown>
    if (
      !['open', 'read'].includes(String(input.action)) ||
      typeof input.documentId !== 'string' ||
      !/^[A-Za-z0-9_-]{8,128}$/.test(input.documentId) ||
      Object.keys(input).sort().join(',') !==
        (input.action === 'open' ? 'action,documentId,markdown' : 'action,documentId') ||
      (input.action === 'open' && !validMarkdown(input.markdown))
    )
      throw new Error('design_document_invalid')
    const key = `${digest(sessionId)}:${input.documentId}`
    let file = files.get(key)
    if (input.action === 'open') {
      const markdown = input.markdown as string
      if (file && file.baseRevision !== digest(markdown))
        throw new Error('design_document_conflict')
      if (!file) {
        if (files.size >= 128) throw new Error('design_document_limit')
        await mkdir(options.directory, { recursive: true, mode: 0o700 })
        current()
        const directory = await mkdtemp(join(await realpath(options.directory), 'design-'))
        const path = join(directory, 'DESIGN.md')
        const handle = await open(path, 'wx', 0o600)
        try {
          await handle.writeFile(markdown, 'utf8')
        } finally {
          await handle.close()
        }
        file = { path, baseRevision: digest(markdown) }
        files.set(key, file)
      }
    }
    if (!file) throw new Error('design_document_missing')
    current()
    let markdown: string
    try {
      if (
        (await realpath(dirname(file.path))) !== dirname(file.path) ||
        (await lstat(file.path)).isSymbolicLink()
      )
        throw new Error('unsafe_file')
      const handle = await open(file.path, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0))
      try {
        const stat = await handle.stat()
        if (!stat.isFile() || stat.size > MAX_BYTES) throw new Error('invalid_file')
        const buffer = Buffer.alloc(MAX_BYTES + 1)
        const { bytesRead } = await handle.read(buffer, 0, buffer.length, 0)
        markdown = buffer.subarray(0, bytesRead).toString('utf8')
        if (!validMarkdown(markdown) || bytesRead > MAX_BYTES) throw new Error('invalid_file')
      } finally {
        await handle.close()
      }
    } catch {
      throw new Error('design_document_unavailable')
    }
    current()
    if (input.action === 'open') await options.openFile(file.path)
    current()
    return { documentId: input.documentId, markdown, revision: digest(markdown) }
  }
}
import { createHash } from 'node:crypto'
import { constants } from 'node:fs'
import { lstat, mkdir, mkdtemp, open, realpath } from 'node:fs/promises'
import { dirname, join } from 'node:path'

const MAX_BYTES = 96 * 1024
const digest = (text: string) => createHash('sha256').update(text).digest('hex')
const validMarkdown = (value: unknown): value is string =>
  typeof value === 'string' &&
  Boolean(value.trim()) &&
  !value.includes('\0') &&
  Buffer.byteLength(value, 'utf8') <= MAX_BYTES
