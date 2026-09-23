/** Internal browser-safe schema primitives shared by presentation contracts. */
export type Schema = {
  type?: string
  properties?: Record<string, Schema>
  required?: string[]
  additionalProperties?: false
  items?: Schema
  minItems?: number
  maxItems?: number
  minLength?: number
  maxLength?: number
  minimum?: number
  maximum?: number
  pattern?: string
  enum?: readonly unknown[]
  anyOf?: Schema[]
}
export const text = (maxLength: number, minLength = 0): Schema => ({
  type: 'string',
  minLength,
  maxLength,
})
export const number = (minimum: number, maximum: number): Schema => ({
  type: 'number',
  minimum,
  maximum,
})
export const choice = (...values: string[]): Schema => ({ type: 'string', enum: values })
export const array = (items: Schema, maxItems: number, minItems = 0): Schema => ({
  type: 'array',
  items,
  minItems,
  maxItems,
})
export const object = (
  properties: Record<string, Schema>,
  required = Object.keys(properties),
): Schema => ({ type: 'object', properties, required, additionalProperties: false })
export const id: Schema = { ...text(80, 1), pattern: '^[A-Za-z0-9_-]+$' }
export const color: Schema = { ...text(6, 6), pattern: '^[A-Fa-f0-9]{6}$' }
export function valid(value: unknown, schema: Schema): boolean {
  if (schema.anyOf) return schema.anyOf.some((candidate) => valid(value, candidate))
  if (schema.enum && !schema.enum.includes(value)) return false
  if (schema.type === 'string')
    return (
      typeof value === 'string' &&
      value.length >= (schema.minLength ?? 0) &&
      value.length <= (schema.maxLength ?? Infinity) &&
      // eslint-disable-next-line no-control-regex -- Reject characters forbidden by XML 1.0.
      !/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value) &&
      (!schema.pattern || new RegExp(schema.pattern).test(value))
    )
  if (schema.type === 'number')
    return (
      typeof value === 'number' &&
      Number.isFinite(value) &&
      value >= (schema.minimum ?? -Infinity) &&
      value <= (schema.maximum ?? Infinity)
    )
  if (schema.type === 'boolean') return typeof value === 'boolean'
  if (schema.type === 'array')
    return (
      Array.isArray(value) &&
      value.length >= (schema.minItems ?? 0) &&
      value.length <= (schema.maxItems ?? Infinity) &&
      value.every((item) => valid(item, schema.items!))
    )
  if (schema.type === 'object') {
    if (!value || typeof value !== 'object' || Array.isArray(value)) return false
    const obj = value as Record<string, unknown>
    return (
      (schema.required ?? []).every((key) => Object.hasOwn(obj, key)) &&
      Object.keys(obj).every(
        (key) =>
          Object.hasOwn(schema.properties!, key) && valid(obj[key], schema.properties![key]!),
      )
    )
  }
  return false
}
