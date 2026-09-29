export interface PresentationPlanAcceptance {
  decisionId: string
  planRevision: number
  planDigest: string
  styleDigest: string
  acceptedAt: string
}
export interface PresentationPlanAcceptanceLedger {
  version: 1
  projectId: string
  documentId: string
  records: PresentationPlanAcceptance[]
}
export function parsePresentationPlanAcceptances(value: unknown): PresentationPlanAcceptanceLedger {
  const ledger = value as PresentationPlanAcceptanceLedger
  const id = (value: unknown) => typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
  const hash = (value: unknown) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
  if (
    !ledger ||
    typeof ledger !== 'object' ||
    Array.isArray(ledger) ||
    Object.keys(ledger).sort().join(',') !== 'documentId,projectId,records,version' ||
    ledger.version !== 1 ||
    !id(ledger.projectId) ||
    typeof ledger.documentId !== 'string' ||
    !ledger.documentId.trim() ||
    ledger.documentId.length > 2048 ||
    !Array.isArray(ledger.records) ||
    ledger.records.length > 64 ||
    new Set(ledger.records.map((record) => record?.decisionId)).size !== ledger.records.length ||
    ledger.records.some(
      (record, index) =>
        !record ||
        Object.keys(record).sort().join(',') !==
          'acceptedAt,decisionId,planDigest,planRevision,styleDigest' ||
        !id(record.decisionId) ||
        !Number.isSafeInteger(record.planRevision) ||
        record.planRevision < 1 ||
        !hash(record.planDigest) ||
        !hash(record.styleDigest) ||
        typeof record.acceptedAt !== 'string' ||
        !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(record.acceptedAt) ||
        !Number.isFinite(Date.parse(record.acceptedAt)) ||
        new Date(record.acceptedAt).toISOString() !== record.acceptedAt ||
        (index > 0 &&
          (record.acceptedAt < ledger.records[index - 1]!.acceptedAt ||
            record.planRevision < ledger.records[index - 1]!.planRevision)),
    )
  )
    throw new Error('invalid_state')
  return structuredClone(ledger)
}
