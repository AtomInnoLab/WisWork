export interface PresentationAssetEventInput {
  type: 'asset.fetching' | 'asset.ready' | 'asset.rejected'
  pageId: string
  assetId: string
  attempt: number
  error?: 'asset_unavailable' | 'output_too_large' | 'aborted'
}
export interface PresentationAssetLedger {
  version: 1
  scope: 'production_asset_resolution'
  projectId: string
  documentId: string
  requestId: string
  inputDigest: string
  planDigest: string
  revision: number
  events: (PresentationAssetEventInput & { sequence: number; createdAt: string })[]
}
export function parsePresentationAssetLedger(value: unknown): PresentationAssetLedger {
  const ledger = value as PresentationAssetLedger
  const id = (value: unknown) => typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
  const hash = (value: unknown) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
  const fail = () => {
    throw new Error('invalid_state')
  }
  if (
    !ledger ||
    typeof ledger !== 'object' ||
    Array.isArray(ledger) ||
    Object.keys(ledger).sort().join(',') !==
      'documentId,events,inputDigest,planDigest,projectId,requestId,revision,scope,version' ||
    ledger.version !== 1 ||
    ledger.scope !== 'production_asset_resolution' ||
    !id(ledger.projectId) ||
    !id(ledger.requestId) ||
    typeof ledger.documentId !== 'string' ||
    !ledger.documentId.trim() ||
    ledger.documentId.length > 2048 ||
    !hash(ledger.inputDigest) ||
    !hash(ledger.planDigest) ||
    !Number.isSafeInteger(ledger.revision) ||
    ledger.revision < 0 ||
    !Array.isArray(ledger.events) ||
    ledger.events.length !== Math.min(ledger.revision, 128)
  )
    fail()
  const states = new Map<string, string>()
  const attempts = new Map<string, number>()
  let lastTime = ''
  for (const [index, event] of ledger.events.entries()) {
    if (
      !event ||
      typeof event !== 'object' ||
      Array.isArray(event) ||
      Object.keys(event).sort().join(',') !==
        (event.type === 'asset.rejected'
          ? 'assetId,attempt,createdAt,error,pageId,sequence,type'
          : 'assetId,attempt,createdAt,pageId,sequence,type') ||
      !['asset.fetching', 'asset.ready', 'asset.rejected'].includes(event.type) ||
      !id(event.pageId) ||
      !id(event.assetId) ||
      !Number.isSafeInteger(event.attempt) ||
      event.attempt < 1 ||
      event.sequence !== ledger.revision - ledger.events.length + index + 1 ||
      typeof event.createdAt !== 'string' ||
      !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(event.createdAt) ||
      !Number.isFinite(Date.parse(event.createdAt)) ||
      new Date(event.createdAt).toISOString() !== event.createdAt ||
      event.createdAt < lastTime ||
      (event.type === 'asset.rejected' &&
        !['asset_unavailable', 'output_too_large', 'aborted'].includes(event.error!))
    )
      fail()
    const asset = JSON.stringify([event.pageId, event.assetId])
    if (event.attempt < (attempts.get(asset) ?? 0)) fail()
    attempts.set(asset, event.attempt)
    const key = JSON.stringify([event.pageId, event.assetId, event.attempt]),
      previous = states.get(key)
    if (
      event.type === 'asset.fetching'
        ? previous !== undefined
        : previous !== 'asset.fetching' && !(previous === undefined && ledger.revision > 128)
    )
      fail()
    states.set(key, event.type)
    lastTime = event.createdAt
  }
  return structuredClone(ledger)
}
