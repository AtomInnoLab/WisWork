import {
  fingerprintSemanticValue,
  type PresentationGeometry,
  type PresentationOperation,
  type PresentationReceipt,
} from '@wiswork/presentation-ops'
import type { SlidesApi } from '../../shared/ipc'
import type { TextFamilyExecutionResult } from './presentation-text-transactions'

export interface GeometryFamilyTransactionRequest {
  transactionId: string
  slideIndex: number
  operations: readonly { sourceId: string; geometry: PresentationGeometry }[]
}

export async function geometryToolTransactionId(call: {
  name: string
  input: Record<string, unknown>
  invocationId?: string
}): Promise<string> {
  if (!call.invocationId) throw new TypeError('Tool invocation identity is required')
  const digest = await fingerprintSemanticValue({
    invocationId: call.invocationId,
    name: call.name,
    input: call.input,
  })
  return `slides-geometry-${digest.slice('sha256:'.length)}`
}

const unchanged = (transactionId: string, operationCount: number): PresentationReceipt => ({
  status: 'unchanged',
  transactionId,
  code: 'write_not_applied',
  operationCount,
})

export async function executePreparedGeometryFamilyTransaction(
  api: Pick<
    SlidesApi,
    'preparePresentationTarget' | 'executePresentationTransaction' | 'cancelPresentationTransaction'
  >,
  request: GeometryFamilyTransactionRequest,
  signal?: AbortSignal,
  refresh?: () => Promise<boolean>,
): Promise<TextFamilyExecutionResult> {
  signal?.throwIfAborted()
  if (request.operations.length === 0 || request.operations.length > 50) {
    return {
      receipt: unchanged(request.transactionId, Math.min(request.operations.length, 50)),
      authoritativeState: 'fresh',
    }
  }
  const prepared = new Map<
    string,
    Extract<Awaited<ReturnType<SlidesApi['preparePresentationTarget']>>, { status: 'prepared' }>
  >()
  let expectedDeckRevision: string | undefined
  for (const item of request.operations) {
    signal?.throwIfAborted()
    if (prepared.has(item.sourceId)) continue
    let result: Awaited<ReturnType<SlidesApi['preparePresentationTarget']>>
    try {
      result = await api.preparePresentationTarget({
        transactionId: request.transactionId,
        slideIndex: request.slideIndex,
        sourceId: item.sourceId,
      })
    } catch {
      return {
        receipt: unchanged(request.transactionId, request.operations.length),
        authoritativeState: 'fresh',
      }
    }
    if (result.status === 'busy') {
      return {
        receipt: unchanged(request.transactionId, request.operations.length),
        authoritativeState: 'fresh',
      }
    }
    if (result.status === 'conflict') {
      return {
        receipt: { status: 'conflict', transactionId: request.transactionId, code: result.code },
        authoritativeState: 'fresh',
      }
    }
    if (expectedDeckRevision && result.expectedDeckRevision !== expectedDeckRevision) {
      return {
        receipt: {
          status: 'conflict',
          transactionId: request.transactionId,
          code: 'target_stale',
        },
        authoritativeState: 'fresh',
      }
    }
    expectedDeckRevision = result.expectedDeckRevision
    prepared.set(item.sourceId, result)
  }

  const operations: PresentationOperation[] = request.operations.map((item, index) => ({
    kind: 'set_geometry',
    clientId: `geometry-${index + 1}`,
    target: prepared.get(item.sourceId)!.target,
    geometry: item.geometry,
  }))
  const cancel = () => void api.cancelPresentationTransaction(request.transactionId)
  signal?.addEventListener('abort', cancel, { once: true })
  try {
    let receipt: PresentationReceipt
    try {
      receipt = await api.executePresentationTransaction({
        transactionId: request.transactionId,
        expectedDeckRevision: expectedDeckRevision!,
        operations,
        mode: 'atomic',
      })
    } catch {
      signal?.throwIfAborted()
      receipt = {
        status: 'uncertain',
        transactionId: request.transactionId,
        code: 'write_state_uncertain',
      }
    }
    if (receipt.status !== 'applied' || !refresh) return { receipt, authoritativeState: 'fresh' }
    try {
      return { receipt, authoritativeState: (await refresh()) ? 'fresh' : 'reload_required' }
    } catch {
      return { receipt, authoritativeState: 'reload_required' }
    }
  } finally {
    signal?.removeEventListener('abort', cancel)
  }
}

/** Convert viewport pixels to PowerPoint points without rounding away deterministic EMUs. */
export function geometryPxToPoints(value: number, viewportScale: number): number {
  if (!Number.isFinite(value) || !Number.isFinite(viewportScale) || viewportScale <= 0)
    throw new TypeError('Invalid presentation geometry scale')
  return (value * 0.75) / viewportScale
}

/** PowerPoint rotations are clockwise and semantically periodic. */
export function normalizePresentationRotation(value: number): number {
  if (!Number.isFinite(value)) throw new TypeError('Invalid presentation rotation')
  const normalized = value % 360
  return normalized < 0 ? normalized + 360 : normalized
}
