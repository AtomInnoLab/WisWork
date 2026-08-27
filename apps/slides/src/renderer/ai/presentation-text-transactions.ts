import {
  fingerprintSemanticValue,
  type PresentationOperation,
  type PresentationReceipt,
} from '@wiswork/presentation-ops'
import type { EditParagraph, SlidesApi } from '../../shared/ipc'

export type TextFamilyOperation =
  | { kind: 'set_text'; paragraphs: readonly EditParagraph[] }
  | { kind: 'set_speaker_notes'; notes: string }

export interface TextFamilyTransactionRequest {
  transactionId: string
  slideIndex: number
  sourceId?: string
  operation: TextFamilyOperation
}

export async function textToolTransactionId(call: {
  id: string
  name: string
  input: Record<string, unknown>
}): Promise<string> {
  const digest = await fingerprintSemanticValue({ id: call.id, name: call.name, input: call.input })
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
): Promise<PresentationReceipt> {
  signal?.throwIfAborted()
  let preparation: Awaited<ReturnType<SlidesApi['preparePresentationTarget']>>
  try {
    preparation = await api.preparePresentationTarget({
      transactionId: request.transactionId,
      slideIndex: request.slideIndex,
      ...(request.sourceId === undefined ? {} : { sourceId: request.sourceId }),
    })
  } catch {
    return {
      status: 'unchanged',
      transactionId: request.transactionId,
      code: 'write_not_applied',
      operationCount: 1,
    }
  }
  if (preparation.status === 'busy') {
    return {
      status: 'unchanged',
      transactionId: request.transactionId,
      code: 'write_not_applied',
      operationCount: 1,
    }
  }
  if (preparation.status === 'conflict') {
    return {
      status: 'conflict',
      transactionId: request.transactionId,
      code: preparation.code,
    }
  }
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
  const cancel = () => void api.cancelPresentationTransaction(request.transactionId)
  signal?.addEventListener('abort', cancel, { once: true })
  try {
    try {
      return await api.executePresentationTransaction(transaction)
    } catch {
      signal?.throwIfAborted()
      return {
        status: 'unchanged',
        transactionId: request.transactionId,
        code: 'write_not_applied',
        operationCount: 1,
      }
    }
  } finally {
    signal?.removeEventListener('abort', cancel)
  }
}
