import {
  fingerprintSemanticValue,
  type PresentationOperation,
  type PresentationReceipt,
} from '@wiswork/presentation-ops'
import type { EditParagraph, SlidesApi } from '../../shared/ipc'
import {
  assertOperationsWithinSelectionScope,
  SelectionScopeConflict,
  type SelectionScope,
} from './edit-queue'

export type TextFamilyOperation =
  | { kind: 'set_text'; paragraphs: readonly EditParagraph[] }
  | { kind: 'set_speaker_notes'; notes: string }

export interface TextFamilyTransactionRequest {
  transactionId: string
  slideIndex: number
  sourceId?: string
  operation: TextFamilyOperation
}

export interface TextFamilyExecutionResult {
  receipt: PresentationReceipt
  authoritativeState: 'fresh' | 'reload_required'
}

const preparationCache = new Map<
  string,
  {
    slideIndex: number
    sourceId?: string
    preparation: Extract<
      Awaited<ReturnType<SlidesApi['preparePresentationTarget']>>,
      { status: 'prepared' }
    >
  }
>()
const MAX_PREPARATION_CACHE = 64

function cachePreparation(
  transactionId: string,
  request: TextFamilyTransactionRequest,
  preparation: Extract<
    Awaited<ReturnType<SlidesApi['preparePresentationTarget']>>,
    { status: 'prepared' }
  >,
): void {
  preparationCache.delete(transactionId)
  preparationCache.set(transactionId, {
    slideIndex: request.slideIndex,
    ...(request.sourceId === undefined ? {} : { sourceId: request.sourceId }),
    preparation,
  })
  while (preparationCache.size > MAX_PREPARATION_CACHE) {
    const oldest = preparationCache.keys().next().value as string | undefined
    if (oldest === undefined) break
    preparationCache.delete(oldest)
  }
}

export async function textToolTransactionId(call: {
  id: string
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
  return `slides-text-${digest.slice('sha256:'.length)}`
}

export function textFamilyReceiptOutcome(receipt: PresentationReceipt): {
  ok: boolean
  mutated: boolean
  detail?: string
} {
  if (receipt.status === 'applied') return { ok: true, mutated: true }
  if (receipt.status === 'unchanged') {
    return receipt.code === 'operation_noop'
      ? { ok: true, mutated: false }
      : { ok: false, mutated: false, detail: receipt.code }
  }
  if (receipt.status === 'conflict') return { ok: false, mutated: false, detail: receipt.code }
  return { ok: false, mutated: true, detail: receipt.code }
}

export async function executePreparedTextFamilyTransaction(
  api: Pick<
    SlidesApi,
    'preparePresentationTarget' | 'executePresentationTransaction' | 'cancelPresentationTransaction'
  >,
  request: TextFamilyTransactionRequest,
  signal?: AbortSignal,
  refresh?: () => Promise<boolean>,
  scope?: SelectionScope,
  onDispatch?: () => void,
): Promise<TextFamilyExecutionResult> {
  signal?.throwIfAborted()
  const cached = preparationCache.get(request.transactionId)
  if (
    cached &&
    (cached.slideIndex !== request.slideIndex || cached.sourceId !== request.sourceId)
  ) {
    return {
      receipt: {
        status: 'conflict',
        transactionId: request.transactionId,
        code: 'target_stale',
      },
      authoritativeState: 'fresh',
    }
  }
  let preparation: Awaited<ReturnType<SlidesApi['preparePresentationTarget']>>
  if (cached) {
    preparationCache.delete(request.transactionId)
    preparationCache.set(request.transactionId, cached)
    preparation = cached.preparation
  } else {
    try {
      preparation = await api.preparePresentationTarget({
        transactionId: request.transactionId,
        slideIndex: request.slideIndex,
        ...(request.sourceId === undefined ? {} : { sourceId: request.sourceId }),
      })
    } catch {
      return {
        receipt: {
          status: 'unchanged',
          transactionId: request.transactionId,
          code: 'write_not_applied',
          operationCount: 1,
        },
        authoritativeState: 'fresh',
      }
    }
  }
  if (preparation.status === 'busy') {
    return {
      receipt: {
        status: 'unchanged',
        transactionId: request.transactionId,
        code: 'write_not_applied',
        operationCount: 1,
      },
      authoritativeState: 'fresh',
    }
  }
  if (preparation.status === 'conflict') {
    return {
      receipt: {
        status: 'conflict',
        transactionId: request.transactionId,
        code: preparation.code,
      },
      authoritativeState: 'fresh',
    }
  }
  if (!cached) cachePreparation(request.transactionId, request, preparation)
  signal?.throwIfAborted()
  const operation: PresentationOperation =
    request.operation.kind === 'set_text'
      ? {
          kind: 'set_text',
          clientId: 'text-1',
          target: preparation.target,
          paragraphs: request.operation.paragraphs.map((paragraph) => ({
            runs: paragraph.runs.map((run) => ({
              text: run.text,
              ...(run.bold === undefined ? {} : { bold: run.bold }),
              ...(run.italic === undefined ? {} : { italic: run.italic }),
              ...(run.underline === undefined ? {} : { underline: run.underline }),
              ...(run.fontSize === undefined ? {} : { fontSize: run.fontSize }),
              ...(run.fontFamily === undefined ? {} : { fontFamily: run.fontFamily }),
              ...(run.color === undefined ? {} : { color: run.color }),
            })),
            ...(paragraph.align === 'left' ||
            paragraph.align === 'center' ||
            paragraph.align === 'right'
              ? { align: paragraph.align }
              : {}),
          })),
        }
      : {
          kind: 'set_speaker_notes',
          clientId: 'notes-1',
          target: preparation.target,
          notes: request.operation.notes,
        }
  const transaction = {
    transactionId: request.transactionId,
    expectedDeckRevision: preparation.expectedDeckRevision,
    operations: [operation],
    mode: 'atomic' as const,
  }
  if (scope) {
    try {
      assertOperationsWithinSelectionScope(scope, [operation])
    } catch (error) {
      if (!(error instanceof SelectionScopeConflict)) throw error
      await api.cancelPresentationTransaction(request.transactionId).catch(() => false)
      return {
        receipt: { status: 'conflict', transactionId: request.transactionId, code: 'target_stale' },
        authoritativeState: 'fresh',
      }
    }
  }
  const cancel = () => void api.cancelPresentationTransaction(request.transactionId)
  signal?.addEventListener('abort', cancel, { once: true })
  try {
    try {
      onDispatch?.()
      const receipt = scope
        ? await api.executePresentationTransaction(transaction, scope)
        : await api.executePresentationTransaction(transaction)
      if (receipt.status !== 'applied' || !refresh) return { receipt, authoritativeState: 'fresh' }
      try {
        return {
          receipt,
          authoritativeState: (await refresh()) ? 'fresh' : 'reload_required',
        }
      } catch {
        return { receipt, authoritativeState: 'reload_required' }
      }
    } catch {
      signal?.throwIfAborted()
      return {
        receipt: {
          status: 'unchanged',
          transactionId: request.transactionId,
          code: 'write_not_applied',
          operationCount: 1,
        },
        authoritativeState: 'fresh',
      }
    }
  } finally {
    signal?.removeEventListener('abort', cancel)
  }
}
