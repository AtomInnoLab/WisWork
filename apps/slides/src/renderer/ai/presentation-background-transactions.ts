import {
  fingerprintSemanticValue,
  parsePresentationTransaction,
  PRESENTATION_OPS_LIMITS,
  type PresentationOperation,
  type PresentationReceipt,
} from '@wiswork/presentation-ops'
import type { SlidesApi } from '../../shared/ipc'
import type { TextFamilyExecutionResult } from './presentation-text-transactions'

export interface BackgroundFamilyTransactionRequest {
  transactionId: string
  backgrounds: readonly { slideIndex: number; color: string }[]
}

export async function backgroundToolTransactionId(call: {
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
  return `slides-background-${digest.slice('sha256:'.length)}`
}

const unchanged = (transactionId: string, operationCount: number): TextFamilyExecutionResult => ({
  receipt: { status: 'unchanged', transactionId, code: 'write_not_applied', operationCount },
  authoritativeState: 'fresh',
})

export async function executePreparedBackgroundFamilyTransaction(
  api: Pick<
    SlidesApi,
    'preparePresentationTarget' | 'executePresentationTransaction' | 'cancelPresentationTransaction'
  >,
  request: BackgroundFamilyTransactionRequest,
  signal?: AbortSignal,
  refresh?: () => Promise<boolean>,
): Promise<TextFamilyExecutionResult> {
  signal?.throwIfAborted()
  if (
    request.backgrounds.length === 0 ||
    request.backgrounds.length > PRESENTATION_OPS_LIMITS.maxOperations ||
    request.backgrounds.some(
      ({ slideIndex, color }) =>
        !Number.isInteger(slideIndex) || slideIndex < 0 || !/^#[0-9A-Fa-f]{6}$/.test(color),
    )
  )
    return unchanged(
      request.transactionId,
      Math.min(request.backgrounds.length, PRESENTATION_OPS_LIMITS.maxOperations),
    )

  const prepared: Array<
    Extract<Awaited<ReturnType<SlidesApi['preparePresentationTarget']>>, { status: 'prepared' }>
  > = []
  const cancel = async () => {
    try {
      await api.cancelPresentationTransaction(request.transactionId)
    } catch {
      // Best effort only; the original outcome remains authoritative.
    }
  }
  let expectedDeckRevision: string | undefined
  for (const background of request.backgrounds) {
    signal?.throwIfAborted()
    let result: Awaited<ReturnType<SlidesApi['preparePresentationTarget']>>
    try {
      result = await api.preparePresentationTarget({
        transactionId: request.transactionId,
        slideIndex: background.slideIndex,
      })
    } catch {
      if (prepared.length) await cancel()
      signal?.throwIfAborted()
      return unchanged(request.transactionId, request.backgrounds.length)
    }
    if (result.status === 'busy') {
      if (prepared.length) await cancel()
      return unchanged(request.transactionId, request.backgrounds.length)
    }
    if (result.status === 'conflict') {
      if (prepared.length) await cancel()
      return {
        receipt: { status: 'conflict', transactionId: request.transactionId, code: result.code },
        authoritativeState: 'fresh',
      }
    }
    if (expectedDeckRevision && expectedDeckRevision !== result.expectedDeckRevision) {
      await cancel()
      return {
        receipt: { status: 'conflict', transactionId: request.transactionId, code: 'target_stale' },
        authoritativeState: 'fresh',
      }
    }
    expectedDeckRevision = result.expectedDeckRevision
    prepared.push(result)
  }

  const operations: PresentationOperation[] = request.backgrounds.map((background, index) => ({
    kind: 'set_slide_background',
    clientId: `background-${index + 1}`,
    target: prepared[index]!.target,
    color: background.color.toUpperCase(),
  }))
  const transaction = parsePresentationTransaction({
    transactionId: request.transactionId,
    expectedDeckRevision,
    operations,
    mode: 'atomic',
  })
  const abort = () => void api.cancelPresentationTransaction(request.transactionId)
  signal?.addEventListener('abort', abort, { once: true })
  try {
    let receipt: PresentationReceipt
    try {
      receipt = await api.executePresentationTransaction(transaction)
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
    signal?.removeEventListener('abort', abort)
  }
}
