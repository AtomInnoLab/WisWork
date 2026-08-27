import {
  fingerprintSemanticValue,
  PRESENTATION_OPS_LIMITS,
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
  const geometryIsBounded = ({
    geometry,
  }: GeometryFamilyTransactionRequest['operations'][number]) =>
    [geometry.x, geometry.y, geometry.width, geometry.height].every(Number.isFinite) &&
    Math.abs(geometry.x) <= PRESENTATION_OPS_LIMITS.maxCoordinateMagnitude &&
    Math.abs(geometry.y) <= PRESENTATION_OPS_LIMITS.maxCoordinateMagnitude &&
    geometry.width > 0 &&
    geometry.width <= PRESENTATION_OPS_LIMITS.maxCoordinateMagnitude &&
    geometry.height > 0 &&
    geometry.height <= PRESENTATION_OPS_LIMITS.maxCoordinateMagnitude &&
    (geometry.rotation === undefined ||
      (Number.isFinite(geometry.rotation) && Math.abs(geometry.rotation) <= 360_000))
  if (
    request.operations.length === 0 ||
    request.operations.length > PRESENTATION_OPS_LIMITS.maxOperations ||
    !request.operations.every(geometryIsBounded)
  ) {
    return {
      receipt: unchanged(
        request.transactionId,
        Math.min(request.operations.length, PRESENTATION_OPS_LIMITS.maxOperations),
      ),
      authoritativeState: 'fresh',
    }
  }
  const prepared = new Map<
    string,
    Extract<Awaited<ReturnType<SlidesApi['preparePresentationTarget']>>, { status: 'prepared' }>
  >()
  const cancelPrepared = async (): Promise<void> => {
    try {
      await api.cancelPresentationTransaction(request.transactionId)
    } catch {
      // Cancellation is best-effort; the original AbortError remains authoritative.
    }
  }
  const throwIfAbortedAfterPrepare = async (): Promise<void> => {
    if (!signal?.aborted) return
    if (prepared.size) await cancelPrepared()
    signal.throwIfAborted()
  }
  let expectedDeckRevision: string | undefined
  for (const item of request.operations) {
    await throwIfAbortedAfterPrepare()
    if (prepared.has(item.sourceId)) continue
    let result: Awaited<ReturnType<SlidesApi['preparePresentationTarget']>>
    try {
      result = await api.preparePresentationTarget({
        transactionId: request.transactionId,
        slideIndex: request.slideIndex,
        sourceId: item.sourceId,
      })
    } catch {
      if (signal?.aborted) {
        await cancelPrepared()
        signal.throwIfAborted()
      }
      return {
        receipt: unchanged(request.transactionId, request.operations.length),
        authoritativeState: 'fresh',
      }
    }
    if (signal?.aborted) {
      await cancelPrepared()
      signal.throwIfAborted()
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
  const cancel = () => {
    try {
      void api.cancelPresentationTransaction(request.transactionId).catch(() => false)
    } catch {
      // The AbortError still propagates from the guarded dispatch boundary.
    }
  }
  if (signal?.aborted) {
    await cancelPrepared()
    signal.throwIfAborted()
  }
  signal?.addEventListener('abort', cancel, { once: true })
  try {
    if (signal?.aborted) {
      await cancelPrepared()
      signal.throwIfAborted()
    }
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
