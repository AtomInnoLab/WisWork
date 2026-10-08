import { createHash, randomUUID } from 'node:crypto'
import { constants } from 'node:fs'
import { lstat, open, rename, rm } from 'node:fs/promises'
import { join } from 'node:path'
import { WISWORK_MESSAGES_URL } from '@wiswork/ai-provider'

const MAX_RECORDS = 2_000
const MAX_FILE_BYTES = 1024 * 1024
const MAX_REQUEST_BYTES = 1024 * 1024
const FILENAME = 'office-outbound-audit.json'

export interface OfficeOutboundRecord {
  id: string
  createdAtMs: number
  updatedAtMs: number
  destination: string
  requestBytes: number
  requestSha256: string
  state: 'attempted' | 'response_received' | 'failed' | 'aborted' | 'auth_required'
  status?: number
}

function validRecord(value: unknown): value is OfficeOutboundRecord {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const record = value as Record<string, unknown>
  return (
    Object.keys(record).every((key) =>
      [
        'id',
        'createdAtMs',
        'updatedAtMs',
        'destination',
        'requestBytes',
        'requestSha256',
        'state',
        'status',
      ].includes(key),
    ) &&
    typeof record.id === 'string' &&
    /^[a-f0-9-]{36}$/.test(record.id) &&
    Number.isSafeInteger(record.createdAtMs) &&
    (record.createdAtMs as number) >= 0 &&
    Number.isSafeInteger(record.updatedAtMs) &&
    (record.updatedAtMs as number) >= (record.createdAtMs as number) &&
    record.destination === WISWORK_MESSAGES_URL &&
    Number.isSafeInteger(record.requestBytes) &&
    (record.requestBytes as number) > 0 &&
    (record.requestBytes as number) <= MAX_REQUEST_BYTES &&
    typeof record.requestSha256 === 'string' &&
    /^[a-f0-9]{64}$/.test(record.requestSha256) &&
    ['attempted', 'response_received', 'failed', 'aborted', 'auth_required'].includes(
      record.state as string,
    ) &&
    (record.status === undefined ||
      (record.state === 'response_received' &&
        Number.isSafeInteger(record.status) &&
        (record.status as number) >= 100 &&
        (record.status as number) <= 599))
  )
}

export function createOfficeOutboundAudit(options: { userDataPath: string; now?: () => number }) {
  const path = join(options.userDataPath, FILENAME)
  const now = options.now ?? Date.now
  let queue = Promise.resolve()
  const locked = <T>(action: () => Promise<T>): Promise<T> => {
    const result = queue.then(action)
    queue = result.then(
      () => undefined,
      () => undefined,
    )
    return result
  }
  const read = async (): Promise<OfficeOutboundRecord[]> => {
    const root = await lstat(options.userDataPath)
    if (!root.isDirectory() || root.isSymbolicLink()) throw new Error('office_audit_invalid_state')
    let handle
    try {
      handle = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW)
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') return []
      throw error
    }
    try {
      const stat = await handle.stat()
      if (!stat.isFile() || stat.size > MAX_FILE_BYTES)
        throw new Error('office_audit_invalid_state')
      let parsed: { version?: unknown; records?: unknown }
      try {
        parsed = JSON.parse((await handle.readFile()).toString()) as typeof parsed
      } catch {
        throw new Error('office_audit_invalid_state')
      }
      if (
        !parsed ||
        parsed.version !== 1 ||
        !Array.isArray(parsed.records) ||
        parsed.records.length > MAX_RECORDS ||
        !parsed.records.every(validRecord) ||
        new Set(parsed.records.map((record) => record.id)).size !== parsed.records.length
      )
        throw new Error('office_audit_invalid_state')
      return parsed.records
    } finally {
      await handle.close()
    }
  }
  const write = async (records: OfficeOutboundRecord[]): Promise<void> => {
    const value = JSON.stringify({ version: 1, records })
    if (Buffer.byteLength(value) > MAX_FILE_BYTES) throw new Error('office_audit_capacity')
    const temp = `${path}.${randomUUID()}.tmp`
    const handle = await open(
      temp,
      constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | constants.O_NOFOLLOW,
      0o600,
    )
    try {
      await handle.writeFile(value)
      await handle.sync()
    } finally {
      await handle.close()
    }
    try {
      await rename(temp, path)
      if (process.platform !== 'win32') {
        const dir = await open(options.userDataPath, constants.O_RDONLY | constants.O_NOFOLLOW)
        try {
          await dir.sync()
        } finally {
          await dir.close()
        }
      }
    } finally {
      await rm(temp, { force: true })
    }
  }
  return {
    list: () => locked(read),
    begin: (body: string, destination: string): Promise<string> =>
      locked(async () => {
        if (
          destination !== WISWORK_MESSAGES_URL ||
          Buffer.byteLength(body) < 1 ||
          Buffer.byteLength(body) > MAX_REQUEST_BYTES
        )
          throw new Error('office_audit_invalid_request')
        const records = await read()
        if (records.length >= MAX_RECORDS) {
          const index = records.findIndex((record) => record.state !== 'attempted')
          if (index < 0) throw new Error('office_audit_capacity')
          records.splice(index, 1)
        }
        const id = randomUUID()
        const time = now()
        records.push({
          id,
          createdAtMs: time,
          updatedAtMs: time,
          destination,
          requestBytes: Buffer.byteLength(body),
          requestSha256: createHash('sha256').update(body).digest('hex'),
          state: 'attempted',
        })
        await write(records)
        return id
      }),
    finish: (
      id: string,
      state: Exclude<OfficeOutboundRecord['state'], 'attempted'>,
      status?: number,
    ): Promise<void> =>
      locked(async () => {
        const records = await read()
        const index = records.findIndex((record) => record.id === id)
        if (index < 0 || records[index]!.state !== 'attempted')
          throw new Error('office_audit_invalid_state')
        if (
          (state === 'response_received' &&
            (!Number.isSafeInteger(status) || status! < 100 || status! > 599)) ||
          (state !== 'response_received' && status !== undefined)
        )
          throw new Error('office_audit_invalid_request')
        records[index] = {
          ...records[index]!,
          state,
          updatedAtMs: now(),
          ...(status ? { status } : {}),
        }
        await write(records)
      }),
  }
}
