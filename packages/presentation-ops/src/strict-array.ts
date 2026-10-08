const unsafeKeys = new Set(['__proto__', 'prototype', 'constructor'])
const arrayIndexPattern = /^(0|[1-9][0-9]*)$/

export const readStrictArray = (
  value: unknown,
  name: string,
  bounds: { minLength?: number; maxLength: number },
): unknown[] => {
  if (!Array.isArray(value) || Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError(`${name} must be a standard array`)
  }

  const lengthDescriptor = Object.getOwnPropertyDescriptor(value, 'length')
  if (
    lengthDescriptor === undefined ||
    lengthDescriptor.get !== undefined ||
    lengthDescriptor.set !== undefined ||
    typeof lengthDescriptor.value !== 'number'
  ) {
    throw new TypeError(`${name} has an invalid length`)
  }
  const length = lengthDescriptor.value
  const minLength = bounds.minLength ?? 0
  if (!Number.isSafeInteger(length) || length < minLength || length > bounds.maxLength) {
    throw new TypeError(`${name} length is out of bounds`)
  }

  const descriptors: PropertyDescriptor[] = []
  for (const key of Reflect.ownKeys(value)) {
    if (key === 'length') continue
    if (typeof key !== 'string') throw new TypeError(`${name} contains an unknown symbol field`)
    if (unsafeKeys.has(key)) throw new TypeError(`${name} contains unsafe field ${key}`)
    if (!arrayIndexPattern.test(key)) throw new TypeError(`${name} contains extra field ${key}`)
    const index = Number(key)
    if (!Number.isSafeInteger(index) || index >= length || String(index) !== key) {
      throw new TypeError(`${name} contains an invalid index`)
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key)
    if (descriptor === undefined || descriptor.get !== undefined || descriptor.set !== undefined) {
      throw new TypeError(`${name} contains an accessor field`)
    }
    if (!descriptor.enumerable) throw new TypeError(`${name} contains a hidden field`)
    descriptors[index] = descriptor
  }

  const result: unknown[] = []
  for (let index = 0; index < length; index += 1) {
    const descriptor = descriptors[index]
    if (descriptor === undefined) throw new TypeError(`${name} must not contain array holes`)
    result.push(descriptor.value)
  }
  return result
}
