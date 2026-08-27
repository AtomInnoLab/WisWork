import type { PresentationReceipt, PresentationTransaction } from '@wiswork/presentation-ops'
import type { PresentationPlan } from './planner'

export function conflictReceipt(
  transaction: PresentationTransaction,
  conflict: Extract<PresentationPlan, { status: 'conflict' }>,
): PresentationReceipt {
  return {
    status: 'conflict',
    transactionId: transaction.transactionId,
    code: conflict.code,
    ...(conflict.operationIndex === undefined ? {} : { operationIndex: conflict.operationIndex }),
    ...(conflict.targetId === undefined ? {} : { targetId: conflict.targetId }),
  }
}

export function unchangedReceipt(
  transaction: PresentationTransaction,
  code: 'operation_noop' | 'write_not_applied',
): PresentationReceipt {
  return {
    status: 'unchanged',
    transactionId: transaction.transactionId,
    code,
    operationCount: transaction.operations.length,
  }
}

export function uncertainReceipt(
  transaction: PresentationTransaction,
  operationIndex?: number,
): PresentationReceipt {
  return {
    status: 'uncertain',
    transactionId: transaction.transactionId,
    code: 'write_state_uncertain',
    ...(operationIndex === undefined ? {} : { operationIndex }),
  }
}
