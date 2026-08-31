export type PresentationElementType = 'text' | 'shape' | 'image' | 'table' | 'chart' | 'group'

export interface PresentationTarget {
  slideId: string
  elementId?: string
  expectedType?: PresentationElementType
  expectedFingerprint?: string
}

/** A target created by an earlier insertion in the same atomic transaction. */
export interface PresentationGeneratedTarget {
  createdByClientId: string
}

export type PresentationElementTarget = PresentationTarget | PresentationGeneratedTarget

export interface PresentationGeometry {
  /** Geometry uses PowerPoint points; rotation uses degrees. */
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

export interface PresentationTextRun {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontSize?: number
  fontFamily?: string
  color?: string
}

export interface PresentationTextParagraph {
  runs: readonly PresentationTextRun[]
  align?: 'left' | 'center' | 'right'
}

export type PresentationTextReplacement =
  | { text: string; paragraphs?: never }
  | { paragraphs: readonly PresentationTextParagraph[]; text?: never }

export type PresentationOperation =
  | (OperationBase & {
      kind: 'set_text'
      target: PresentationElementTarget
    } & PresentationTextReplacement)
  | (OperationBase & {
      kind: 'set_geometry'
      target: PresentationElementTarget
      geometry: PresentationGeometry
    })
  | (OperationBase & {
      kind: 'set_fill'
      target: PresentationElementTarget
      fill: PresentationFill
    })
  | (OperationBase & {
      kind: 'set_stroke'
      target: PresentationElementTarget
      stroke: PresentationStroke | null
    })
  | (OperationBase & {
      kind: 'add_text_box'
      slideId: string
      text: string
      geometry: PresentationGeometry
    })
  | (OperationBase & { kind: 'delete_element'; target: PresentationElementTarget })
  | (OperationBase & {
      kind: 'set_speaker_notes'
      target: PresentationTarget
      notes: string
    })
  | (OperationBase & {
      kind: 'set_slide_background'
      target: PresentationTarget
      color: string
    })

export interface PresentationTransaction {
  transactionId: string
  expectedDeckRevision: string
  operations: readonly PresentationOperation[]
  mode: 'atomic'
}

export interface PresentationCreatedTarget {
  clientId: string
  elementId: string
}

export type PresentationReceipt =
  | {
      status: 'applied'
      transactionId: string
      resultingDeckRevision: string
      operationCount: number
      /** Opaque authority tokens actually mutated by the committed transaction. */
      mutatedTargets?: readonly string[]
      createdIds?: readonly string[]
      createdTargets?: readonly PresentationCreatedTarget[]
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

export type PresentationQualityCode =
  | 'element_off_slide'
  | 'text_overflow_horizontal'
  | 'text_overflow_vertical'
  | 'element_collision'
  | 'tiny_text'
  | 'empty_placeholder'
  | 'low_contrast'
  | 'visual_quality'
  | 'quality_capacity_exceeded'

export type PresentationQualitySeverity = 'warning' | 'important' | 'critical'

export interface PresentationQualityFinding {
  code: PresentationQualityCode
  severity: PresentationQualitySeverity
  slideId: string
  elementId?: string
  relatedElementId?: string
  /** Numeric, geometry-only evidence. Raw slide content is never part of this contract. */
  evidence: Readonly<Record<string, number>>
}

export type PresentationQualityReceipt =
  | {
      qualityRunId: string
      transactionId: string
      slideId: string
      source: 'deterministic' | 'visual'
      status: 'available'
      findings: readonly PresentationQualityFinding[]
      truncated: boolean
    }
  | {
      qualityRunId: string
      transactionId: string
      slideId: string
      source: 'visual'
      status: 'unavailable' | 'cancelled'
      code:
        | 'screenshot_unavailable'
        | 'transport_unavailable'
        | 'cancelled'
        | 'stale_session'
        | 'visual_capacity_exceeded'
    }
