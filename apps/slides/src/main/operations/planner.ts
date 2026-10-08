import type { PresentationOperation } from '@wiswork/presentation-ops'
import { assertRegisteredPresentationOperation } from './registry'

export interface PlannedPresentationOperation {
  index: number
  operation: PresentationOperation
  createdId?: string
}

export type PresentationPlan =
  | {
      status: 'planned'
      operations: readonly PlannedPresentationOperation[]
      noOp: boolean
    }
  | {
      status: 'conflict'
      code: 'target_stale' | 'target_missing' | 'target_ambiguous'
      operationIndex?: number
      targetId?: string
    }

export function validatePlannedOperations(
  operations: readonly PlannedPresentationOperation[],
): void {
  for (const [position, planned] of operations.entries()) {
    if (planned.index !== position) throw new TypeError('Presentation plan is not sequential')
    assertRegisteredPresentationOperation(planned.operation)
    if (planned.operation.kind === 'add_text_box' && !planned.createdId) {
      throw new TypeError('Presentation insertion has no allocated durable id')
    }
    if (planned.operation.kind !== 'add_text_box' && planned.createdId !== undefined) {
      throw new TypeError('Only presentation insertions may allocate durable ids')
    }
  }
}
