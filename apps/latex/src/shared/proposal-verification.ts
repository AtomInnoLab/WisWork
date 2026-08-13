export const MAX_PROPOSAL_DIAGNOSTICS = 100
export const MAX_FORMAL_COMPILE_DIAGNOSTICS = 1_000
export const MAX_PROPOSAL_DIAGNOSTIC_PATH_BYTES = 1_024
export const MAX_PROPOSAL_DIAGNOSTIC_MESSAGE_BYTES = 4_096
export const MAX_PROPOSAL_DIAGNOSTIC_POSITION = 10_000_000

export interface NormalizedProposalDiagnostic {
  path: string | null
  line: number | null
  column: number | null
  severity: 'error' | 'warning'
  message: string
}

const encoder = new TextEncoder()

export function truncateUtf8(value: string, maxBytes: number): string {
  if (encoder.encode(value).byteLength <= maxBytes) return value
  let result = ''
  let bytes = 0
  for (const character of value) {
    const size = encoder.encode(character).byteLength
    if (bytes + size > maxBytes) break
    result += character
    bytes += size
  }
  return result
}

function position(value: unknown): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value)) return null
  return Math.min(MAX_PROPOSAL_DIAGNOSTIC_POSITION, Math.max(1, Math.trunc(value)))
}

export function normalizeProposalDiagnostics(
  values: readonly unknown[],
  maxDiagnostics = MAX_PROPOSAL_DIAGNOSTICS,
): NormalizedProposalDiagnostic[] {
  if (
    !Number.isSafeInteger(maxDiagnostics) ||
    maxDiagnostics < 1 ||
    maxDiagnostics > MAX_FORMAL_COMPILE_DIAGNOSTICS
  ) {
    throw new Error('Invalid diagnostic count limit')
  }
  const result: NormalizedProposalDiagnostic[] = []
  for (const value of values) {
    if (result.length >= maxDiagnostics) break
    if (!value || typeof value !== 'object' || Array.isArray(value)) continue
    const record = value as Record<string, unknown>
    if (record.severity !== 'error' && record.severity !== 'warning') continue
    if (typeof record.message !== 'string') continue
    const message = truncateUtf8(
      record.message.replaceAll('\0', ''),
      MAX_PROPOSAL_DIAGNOSTIC_MESSAGE_BYTES,
    )
    if (!message) continue
    let path: string | null = null
    if (typeof record.path === 'string' && record.path && !record.path.includes('\0')) {
      path = truncateUtf8(record.path, MAX_PROPOSAL_DIAGNOSTIC_PATH_BYTES) || null
    } else if (record.path !== null) {
      continue
    }
    result.push({
      path,
      line: position(record.line),
      column: position(record.column),
      severity: record.severity,
      message,
    })
  }
  return result
}
