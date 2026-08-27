export type PresentationElementType = 'text' | 'shape' | 'image' | 'table' | 'chart' | 'group'

export interface PresentationTarget {
  slideId: string
  elementId?: string
  expectedType?: PresentationElementType
  expectedFingerprint?: string
}

export interface PresentationGeometry {
  x: number
  y: number
  width: number
  height: number
  rotation?: number
}

export type PresentationFill =
  { kind: 'none' } | { kind: 'solid'; color: string; transparency?: number }

export interface PresentationStroke {
  color: string
  width: number
  dash?: 'solid' | 'dash' | 'dot' | 'dash_dot'
}

interface OperationBase {
  clientId: string
}

export type PresentationOperation =
  | (OperationBase & { kind: 'set_text'; target: PresentationTarget; text: string })
  | (OperationBase & {
      kind: 'set_geometry'
      target: PresentationTarget
      geometry: PresentationGeometry
    })
  | (OperationBase & { kind: 'set_fill'; target: PresentationTarget; fill: PresentationFill })
  | (OperationBase & {
      kind: 'set_stroke'
      target: PresentationTarget
      stroke: PresentationStroke
    })
  | (OperationBase & {
      kind: 'add_text_box'
      slideId: string
      text: string
      geometry: PresentationGeometry
    })
  | (OperationBase & { kind: 'delete_element'; target: PresentationTarget })
  | (OperationBase & {
      kind: 'set_speaker_notes'
      target: PresentationTarget
      notes: string
    })

export interface PresentationTransaction {
  transactionId: string
  expectedDeckRevision: string
  operations: readonly PresentationOperation[]
  mode: 'atomic'
}

export type PresentationReceipt =
  | {
      status: 'applied'
      transactionId: string
      resultingDeckRevision: string
      operationCount: number
      createdIds?: readonly string[]
    }
  | {
      status: 'unchanged'
      transactionId: string
      code: 'operation_noop' | 'write_not_applied'
      operationCount: number
    }
  | {
      status: 'conflict'
      transactionId: string
      code: 'target_stale' | 'target_missing' | 'target_ambiguous'
      operationIndex?: number
      targetId?: string
    }
  | {
      status: 'uncertain'
      transactionId: string
      code: 'write_state_uncertain'
      operationIndex?: number
    }
