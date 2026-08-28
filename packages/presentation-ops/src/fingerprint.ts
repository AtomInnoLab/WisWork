import { readStrictArray } from './strict-array'

const unsafeKeys = new Set(['__proto__', 'prototype', 'constructor'])
const MAX_DEPTH = 32
const MAX_NODES = 10_000
const MAX_STRING_LENGTH = 12_000
const MAX_COLLECTION_LENGTH = 1_000

export const canonicalizeSemanticValue = (value: unknown): string => {
  const active = new Set<object>()
  let nodes = 0

  const visit = (item: unknown, depth: number): string => {
    nodes += 1
    if (nodes > MAX_NODES || depth > MAX_DEPTH)
      throw new TypeError('Semantic fingerprint input exceeds bounds')
    if (item === null) return 'null'
    if (typeof item === 'boolean') return item ? 'true' : 'false'
    if (typeof item === 'number') {
      if (!Number.isFinite(item)) throw new TypeError('Semantic fingerprint numbers must be finite')
      return Object.is(item, -0) ? '0' : JSON.stringify(item)
    }
    if (typeof item === 'string') {
      if (item.length > MAX_STRING_LENGTH)
        throw new TypeError('Semantic fingerprint string exceeds bounds')
      return JSON.stringify(item.normalize('NFC'))
    }
    if (typeof item !== 'object' || item === undefined)
      throw new TypeError('Unsupported semantic fingerprint value')

    const object = item as object
    if (active.has(object)) throw new TypeError('Cyclic semantic fingerprint value')
    active.add(object)
    try {
      if (Array.isArray(item)) {
        const values = readStrictArray(item, 'Semantic fingerprint array', {
          maxLength: MAX_COLLECTION_LENGTH,
        })
        const children: string[] = []
        for (let index = 0; index < values.length; index += 1) {
          children.push(visit(values[index], depth + 1))
        }
        return `[${children.join(',')}]`
      }
      const prototype = Object.getPrototypeOf(item)
      if (prototype !== Object.prototype && prototype !== null)
        throw new TypeError('Semantic fingerprint values must be plain objects')
      const record = item as Record<string, unknown>
      const ownKeys = Reflect.ownKeys(record)
      if (ownKeys.some((key) => typeof key !== 'string'))
        throw new TypeError('Semantic fingerprint values cannot contain symbol fields')
      const keys = (ownKeys as string[]).sort()
      if (keys.length > MAX_COLLECTION_LENGTH)
        throw new TypeError('Semantic fingerprint object exceeds bounds')
      for (const key of keys) {
        if (unsafeKeys.has(key)) throw new TypeError(`Unsafe semantic fingerprint key ${key}`)
        if (key.length > MAX_STRING_LENGTH)
          throw new TypeError('Semantic fingerprint key exceeds bounds')
        const descriptor = Object.getOwnPropertyDescriptor(record, key)
        if (
          descriptor === undefined ||
          descriptor.get !== undefined ||
          descriptor.set !== undefined
        )
          throw new TypeError('Semantic fingerprint values cannot contain accessor fields')
        if (!descriptor.enumerable)
          throw new TypeError('Semantic fingerprint values cannot contain hidden fields')
      }
      return `{${keys.map((key) => `${JSON.stringify(key.normalize('NFC'))}:${visit(record[key], depth + 1)}`).join(',')}}`
    } finally {
      active.delete(object)
    }
  }

  return visit(value, 0)
}

export const fingerprintSemanticValue = async (value: unknown): Promise<string> => {
  const bytes = new TextEncoder().encode(canonicalizeSemanticValue(value))
  const digest = await globalThis.crypto.subtle.digest('SHA-256', bytes)
  const hex = Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join(
    '',
  )
  return `sha256:${hex}`
}
