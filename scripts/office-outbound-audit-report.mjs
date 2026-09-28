import { constants } from 'node:fs'
import { lstat, open } from 'node:fs/promises'
import { join, resolve } from 'node:path'

const root = process.argv[2]
if (!root || process.argv.length !== 3) {
  process.stderr.write('Usage: node scripts/office-outbound-audit-report.mjs <PC userDataPath>\n')
  process.exit(2)
}

try {
  const directory = resolve(root)
  const info = await lstat(directory)
  if (!info.isDirectory() || info.isSymbolicLink()) throw new Error('invalid_audit_directory')
  const handle = await open(
    join(directory, 'office-outbound-audit.json'),
    constants.O_RDONLY | constants.O_NOFOLLOW,
  )
  let raw
  try {
    const stat = await handle.stat()
    if (!stat.isFile() || stat.size > 1024 * 1024) throw new Error('invalid_audit_file')
    raw = await handle.readFile()
  } finally {
    await handle.close()
  }
  const parsed = JSON.parse(raw.toString())
  if (parsed.version !== 1 || !Array.isArray(parsed.records) || parsed.records.length > 2_000)
    throw new Error('invalid_audit_file')
  const counts = {
    attempted: 0,
    response_received: 0,
    failed: 0,
    aborted: 0,
    auth_required: 0,
  }
  const records = parsed.records.map((item) => {
    if (
      !item ||
      typeof item.id !== 'string' ||
      !Number.isSafeInteger(item.createdAtMs) ||
      !Number.isSafeInteger(item.updatedAtMs) ||
      typeof item.destination !== 'string' ||
      !Number.isSafeInteger(item.requestBytes) ||
      !/^[a-f0-9]{64}$/.test(item.requestSha256) ||
      !Object.hasOwn(counts, item.state)
    )
      throw new Error('invalid_audit_file')
    counts[item.state] += 1
    return {
      id: item.id,
      createdAtMs: item.createdAtMs,
      updatedAtMs: item.updatedAtMs,
      destination: item.destination,
      requestBytes: item.requestBytes,
      requestSha256: item.requestSha256,
      state: item.state,
      ...(item.state === 'response_received' && Number.isSafeInteger(item.status)
        ? { status: item.status }
        : {}),
    }
  })
  process.stdout.write(
    JSON.stringify({ total: records.length, counts, records: records.slice(-50) }, null, 2) + '\n',
  )
} catch {
  process.stderr.write('office_outbound_audit_unavailable\n')
  process.exit(1)
}
