import {
  fingerprintSemanticValue,
  parsePresentationTransaction,
  PRESENTATION_OPS_LIMITS,
  type PresentationGeometry,
  type PresentationElementTarget,
  type PresentationFill,
  type PresentationStroke,
  type PresentationTextParagraph,
  type PresentationOperation,
  type PresentationReceipt,
} from '@wiswork/presentation-ops'
import type { SlidesApi } from '../../shared/ipc'
import type { TextFamilyExecutionResult } from './presentation-text-transactions'

export interface GeometryFamilyTransactionRequest {
  transactionId: string
  slideIndex: number
  operations: readonly CanonicalElementOperation[]
}

export type CanonicalElementOperation =
  | ({ kind?: 'set_geometry'; geometry: PresentationGeometry } & CanonicalTargetReference)
  | ({ kind: 'set_fill'; fill: PresentationFill } & CanonicalTargetReference)
  | ({ kind: 'set_stroke'; stroke: PresentationStroke | null } & CanonicalTargetReference)
  | ({
      kind: 'set_text'
      paragraphs: readonly PresentationTextParagraph[]
    } & CanonicalTargetReference)
  | {
      kind: 'add_text_box'
      clientId: string
      text: string
      geometry: PresentationGeometry
    }
  | ({ kind: 'delete_element' } & CanonicalTargetReference)

type CanonicalTargetReference =
  { sourceId: string; createdByClientId?: never } | { createdByClientId: string; sourceId?: never }

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
  const compileOperation = (
    item: CanonicalElementOperation,
    clientId: string,
    target: PresentationElementTarget,
  ): PresentationOperation => {
    if (item.kind === 'add_text_box')
      return {
        kind: 'add_text_box',
        clientId,
        slideId:
          'slideId' in target
            ? target.slideId
            : (() => {
                throw new TypeError('Insertion requires a slide target')
              })(),
        text: item.text,
        geometry: item.geometry,
      }
    if ('geometry' in item)
      return {
        kind: 'set_geometry',
        clientId,
        target,
        geometry: item.geometry,
      }
    if (item.kind === 'set_fill') return { kind: 'set_fill', clientId, target, fill: item.fill }
    if (item.kind === 'set_stroke')
      return { kind: 'set_stroke', clientId, target, stroke: item.stroke }
    if (item.kind === 'delete_element') return { kind: 'delete_element', clientId, target }
    return { kind: 'set_text', clientId, target, paragraphs: item.paragraphs }
  }
  const geometryIsBounded = (item: GeometryFamilyTransactionRequest['operations'][number]) => {
    if (!('geometry' in item)) return true
    const { geometry } = item
    return (
      [geometry.x, geometry.y, geometry.width, geometry.height].every(Number.isFinite) &&
      Math.abs(geometry.x) <= PRESENTATION_OPS_LIMITS.maxCoordinateMagnitude &&
      Math.abs(geometry.y) <= PRESENTATION_OPS_LIMITS.maxCoordinateMagnitude &&
      geometry.width > 0 &&
      geometry.width <= PRESENTATION_OPS_LIMITS.maxCoordinateMagnitude &&
      geometry.height > 0 &&
      geometry.height <= PRESENTATION_OPS_LIMITS.maxCoordinateMagnitude &&
      (geometry.rotation === undefined ||
        (Number.isFinite(geometry.rotation) && Math.abs(geometry.rotation) <= 360_000))
    )
  }
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
  const internalClientIds = request.operations.map((_item, index) => `op-${index + 1}`)
  const generatedClientIds = new Map<string, string>()
  const preflightOperations: PresentationOperation[] = []
  try {
    request.operations.forEach((item, index) => {
      let target: PresentationElementTarget
      if (item.kind === 'add_text_box') {
        if (generatedClientIds.has(item.clientId))
          throw new TypeError('Duplicate generated target alias')
        generatedClientIds.set(item.clientId, internalClientIds[index]!)
        target = { slideId: 'slide-preflight' }
      } else if ('createdByClientId' in item && typeof item.createdByClientId === 'string') {
        const generatedClientId = generatedClientIds.get(item.createdByClientId)
        if (!generatedClientId) throw new TypeError('Generated target alias must refer backward')
        target = { createdByClientId: generatedClientId }
      } else {
        target = {
          slideId: 'slide-preflight',
          elementId: 'element-preflight',
          expectedFingerprint: `sha256:${'0'.repeat(64)}`,
        }
      }
      preflightOperations.push(compileOperation(item, internalClientIds[index]!, target))
    })
    parsePresentationTransaction({
      transactionId: request.transactionId,
      expectedDeckRevision: `sha256:${'0'.repeat(64)}`,
      operations: preflightOperations,
      mode: 'atomic',
    })
  } catch {
    return {
      receipt: unchanged(request.transactionId, request.operations.length),
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
    if ('createdByClientId' in item && typeof item.createdByClientId === 'string') continue
    const sourceId = item.kind === 'add_text_box' ? undefined : item.sourceId
    const preparationKey = sourceId ?? '__slide_container__'
    if (prepared.has(preparationKey)) continue
    let result: Awaited<ReturnType<SlidesApi['preparePresentationTarget']>>
    try {
      result = await api.preparePresentationTarget({
        transactionId: request.transactionId,
        slideIndex: request.slideIndex,
        ...(sourceId === undefined ? {} : { sourceId }),
      })
    } catch {
      if (signal?.aborted) {
        await cancelPrepared()
        signal.throwIfAborted()
      }
      if (prepared.size) await cancelPrepared()
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
      if (prepared.size) await cancelPrepared()
      return {
        receipt: unchanged(request.transactionId, request.operations.length),
        authoritativeState: 'fresh',
      }
    }
    if (result.status === 'conflict') {
      if (prepared.size) await cancelPrepared()
      return {
        receipt: { status: 'conflict', transactionId: request.transactionId, code: result.code },
        authoritativeState: 'fresh',
      }
    }
    if (expectedDeckRevision && result.expectedDeckRevision !== expectedDeckRevision) {
      if (prepared.size) await cancelPrepared()
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
    prepared.set(preparationKey, result)
  }

  const operations: PresentationOperation[] = request.operations.map((item, index) => {
    if ('createdByClientId' in item && typeof item.createdByClientId === 'string') {
      const generatedClientId = generatedClientIds.get(item.createdByClientId)!
      return compileOperation(item, internalClientIds[index]!, {
        createdByClientId: generatedClientId,
      })
    }
    const key = item.kind === 'add_text_box' ? '__slide_container__' : item.sourceId!
    return compileOperation(item, internalClientIds[index]!, prepared.get(key)!.target)
  })
  try {
    parsePresentationTransaction({
      transactionId: request.transactionId,
      expectedDeckRevision: expectedDeckRevision!,
      operations,
      mode: 'atomic',
    })
  } catch {
    await cancelPrepared()
    return {
      receipt: unchanged(request.transactionId, request.operations.length),
      authoritativeState: 'fresh',
    }
  }
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
    if (receipt.status !== 'applied' && receipt.status !== 'uncertain')
      return { receipt, authoritativeState: 'fresh' }
    if (!refresh)
      return {
        receipt,
        authoritativeState: receipt.status === 'uncertain' ? 'reload_required' : 'fresh',
      }
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
