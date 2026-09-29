export interface PresentationPreferenceSource {
  documentId: string
  projectId: string
  changeId: string
}
export interface SavedPresentationPreference {
  projectId: string
  changeId: string
  text: string
  reuse?: {
    version: 1
    source: PresentationPreferenceSource
    sourceTextDigest: string
    approvedAt: string
    approvalId: string
  }
}
const fail = (): never => {
  throw Error('invalid_request')
}
const object = (value: unknown): Record<string, unknown> => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return fail()
  return value as Record<string, unknown>
}
const keys = (value: Record<string, unknown>, expected: string[]) =>
  Object.keys(value).sort().join(',') === expected.sort().join(',')
const id = (value: unknown, max: number) =>
  typeof value === 'string' && new RegExp(`^[A-Za-z0-9_-]{1,${max}}$`).test(value)
export function parsePresentationPreferenceSource(value: unknown): PresentationPreferenceSource {
  const source = object(value)
  if (
    !keys(source, ['documentId', 'projectId', 'changeId']) ||
    typeof source.documentId !== 'string' ||
    !source.documentId ||
    source.documentId.length > 2048 ||
    !id(source.projectId, 80) ||
    !id(source.changeId, 128)
  )
    return fail()
  return structuredClone(source) as unknown as PresentationPreferenceSource
}
export function parseSavedPresentationPreference(value: unknown): SavedPresentationPreference {
  const preference = object(value)
  if (
    !keys(preference, [
      'projectId',
      'changeId',
      'text',
      ...(Object.hasOwn(preference, 'reuse') ? ['reuse'] : []),
    ]) ||
    !id(preference.projectId, 80) ||
    !id(preference.changeId, 128) ||
    typeof preference.text !== 'string' ||
    !preference.text.trim() ||
    preference.text.length > 240 ||
    Array.from(preference.text).some(
      (c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159),
    )
  )
    return fail()
  if (Object.hasOwn(preference, 'reuse')) {
    const reuse = object(preference.reuse)
    if (
      !keys(reuse, ['version', 'source', 'sourceTextDigest', 'approvedAt', 'approvalId']) ||
      reuse.version !== 1 ||
      typeof reuse.sourceTextDigest !== 'string' ||
      !/^[a-f0-9]{64}$/.test(reuse.sourceTextDigest) ||
      typeof reuse.approvalId !== 'string' ||
      !/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/.test(reuse.approvalId) ||
      typeof reuse.approvedAt !== 'string' ||
      !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(reuse.approvedAt) ||
      !Number.isFinite(Date.parse(reuse.approvedAt)) ||
      new Date(reuse.approvedAt).toISOString() !== reuse.approvedAt
    )
      return fail()
    parsePresentationPreferenceSource(reuse.source)
  }
  return structuredClone(preference) as unknown as SavedPresentationPreference
}
/** Deterministic identity for one exact, explicitly selected source snapshot and target scope. */
export async function presentationPreferenceReuseId(
  targetDocumentId: string,
  targetProjectId: string,
  source: PresentationPreferenceSource,
  expectedTextDigest: string,
): Promise<string> {
  parsePresentationPreferenceSource({
    documentId: targetDocumentId,
    projectId: targetProjectId,
    changeId: 'validation',
  })
  const parsed = parsePresentationPreferenceSource(source)
  if (!/^[a-f0-9]{64}$/.test(expectedTextDigest)) return fail()
  const input = JSON.stringify([
    targetDocumentId,
    targetProjectId,
    parsed.documentId,
    parsed.projectId,
    parsed.changeId,
    expectedTextDigest,
  ])
  const bytes = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(input))
  return (
    'reuse_' + Array.from(new Uint8Array(bytes), (n) => n.toString(16).padStart(2, '0')).join('')
  )
}
