import type { PresentationOperation } from '@wiswork/presentation-ops'

export const PRESENTATION_OPERATION_KINDS = Object.freeze([
  'set_text',
  'set_geometry',
  'set_fill',
  'set_stroke',
  'add_text_box',
  'delete_element',
  'set_speaker_notes',
] as const satisfies readonly PresentationOperation['kind'][])

const operationKinds = new Set<string>(PRESENTATION_OPERATION_KINDS)

export function assertRegisteredPresentationOperation(operation: PresentationOperation): void {
  if (!operationKinds.has(operation.kind)) {
    throw new TypeError('Unsupported presentation operation')
  }
}
