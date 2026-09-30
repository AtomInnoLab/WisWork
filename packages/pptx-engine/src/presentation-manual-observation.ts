import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'

export const MAX_PRESENTATION_MANUAL_OBSERVATION_BYTES = 64 * 1024
export const MAX_PRESENTATION_MANUAL_OBSERVATIONS_BYTES = 1024 * 1024
export const MAX_PRESENTATION_MANUAL_OBSERVATIONS = 32
export interface PresentationManualObservationShape {
  id: string
  name: string
  type: 'TextBox' | 'GeometricShape' | 'Placeholder'
  left: number
  top: number
  width: number
  height: number
  rotation?: number
  text?: string
  font?: {
    name: string | null
    size: number | null
    color: string | null
    bold?: boolean | null
    italic?: boolean | null
    underline?: string | null
  }
}
export interface PresentationManualObservation {
  version: 1
  source: 'host_difference_unattributed'
  observationId: string
  documentId: string
  projectId: string
  slideId: string
  shapeId: string
  before: { capturedAt: string; shape: PresentationManualObservationShape; digest: string }
  after?: { capturedAt: string; shape: PresentationManualObservationShape; digest: string }
  atomicSnapshot: false
  coverage: 'text_geometry_aggregate_font'
}
const fail = (): never => {
  throw new Error('invalid_request')
}
const object = (v: unknown): Record<string, unknown> =>
  !v || typeof v !== 'object' || Array.isArray(v) ? fail() : (v as Record<string, unknown>)
const keys = (v: Record<string, unknown>, required: string[], optional: string[] = []) =>
  required.every((k) => Object.hasOwn(v, k)) &&
  Object.keys(v).every((k) => required.includes(k) || optional.includes(k))
const text = (v: unknown, max: number, empty = false): v is string =>
  typeof v === 'string' &&
  (empty || v.length > 0) &&
  v.length <= max &&
  !Array.from(v).some(
    (c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159),
  )
const id = (v: unknown, max: number) =>
  typeof v === 'string' && new RegExp(`^[A-Za-z0-9_-]{1,${max}}$`).test(v)
const hash = (v: unknown) => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v)
const iso = (v: unknown) =>
  typeof v === 'string' &&
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(v) &&
  Number.isFinite(Date.parse(v)) &&
  new Date(v).toISOString() === v
export function parsePresentationManualObservationShape(
  value: unknown,
): PresentationManualObservationShape {
  const v = object(value)
  if (
    !keys(
      v,
      ['id', 'name', 'type', 'left', 'top', 'width', 'height'],
      ['rotation', 'text', 'font'],
    ) ||
    !text(v.id, 256) ||
    !text(v.name, 256, true) ||
    !['TextBox', 'GeometricShape', 'Placeholder'].includes(v.type as string) ||
    (v.type === 'Placeholder' && !Object.hasOwn(v, 'text')) ||
    ['left', 'top', 'width', 'height'].some(
      (k) => typeof v[k] !== 'number' || !Number.isFinite(v[k]) || Math.abs(v[k] as number) > 1e6,
    ) ||
    (v.width as number) <= 0 ||
    (v.height as number) <= 0 ||
    (Object.hasOwn(v, 'rotation') &&
      (typeof v.rotation !== 'number' ||
        !Number.isFinite(v.rotation) ||
        Math.abs(v.rotation) > 360)) ||
    (Object.hasOwn(v, 'text') && (typeof v.text !== 'string' || v.text.length > 32000))
  )
    return fail()
  if (Object.hasOwn(v, 'font')) {
    const f = object(v.font)
    if (
      !keys(f, ['name', 'size', 'color'], ['bold', 'italic', 'underline']) ||
      (f.name !== null && !text(f.name, 256, true)) ||
      (f.color !== null && !text(f.color, 128, true)) ||
      (f.size !== null &&
        (typeof f.size !== 'number' || !Number.isFinite(f.size) || f.size <= 0 || f.size > 1000)) ||
      ['bold', 'italic'].some(
        (k) => Object.hasOwn(f, k) && f[k] !== null && typeof f[k] !== 'boolean',
      ) ||
      (Object.hasOwn(f, 'underline') && f.underline !== null && !text(f.underline, 128, true))
    )
      return fail()
  }
  return structuredClone(v) as unknown as PresentationManualObservationShape
}
export function parsePresentationManualObservation(value: unknown): PresentationManualObservation {
  const v = object(value)
  if (
    !keys(
      v,
      [
        'version',
        'source',
        'observationId',
        'documentId',
        'projectId',
        'slideId',
        'shapeId',
        'before',
        'atomicSnapshot',
        'coverage',
      ],
      ['after'],
    ) ||
    v.version !== 1 ||
    v.source !== 'host_difference_unattributed' ||
    !id(v.observationId, 80) ||
    !text(v.documentId, 2048) ||
    !id(v.projectId, 80) ||
    !text(v.slideId, 256) ||
    !text(v.shapeId, 256) ||
    v.atomicSnapshot !== false ||
    v.coverage !== 'text_geometry_aggregate_font' ||
    new TextEncoder().encode(JSON.stringify(v)).byteLength >
      MAX_PRESENTATION_MANUAL_OBSERVATION_BYTES
  )
    return fail()
  const snapshot = (value: unknown) => {
    const s = object(value)
    if (!keys(s, ['capturedAt', 'shape', 'digest']) || !iso(s.capturedAt) || !hash(s.digest))
      return fail()
    const shape = parsePresentationManualObservationShape(s.shape)
    if (shape.id !== v.shapeId) return fail()
    return { capturedAt: s.capturedAt as string, shape, digest: s.digest as string }
  }
  const before = snapshot(v.before),
    after = Object.hasOwn(v, 'after') ? snapshot(v.after) : undefined
  if (after && (after.capturedAt < before.capturedAt || after.shape.type !== before.shape.type))
    return fail()
  return {
    ...(structuredClone(v) as unknown as PresentationManualObservation),
    before,
    ...(after ? { after } : {}),
  }
}
export async function presentationManualObservationDigest(
  shape: PresentationManualObservationShape,
): Promise<string> {
  const bytes = await crypto.subtle.digest(
    'SHA-256',
    new TextEncoder().encode(
      canonicalPresentationValue(parsePresentationManualObservationShape(shape)),
    ),
  )
  return Array.from(new Uint8Array(bytes), (n) => n.toString(16).padStart(2, '0')).join('')
}
