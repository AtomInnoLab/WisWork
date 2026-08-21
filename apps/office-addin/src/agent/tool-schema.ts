export type FieldParser<T> = (value: unknown) => T

export function stringField(options: {
  minLength?: number
  maxLength: number
}): FieldParser<string> {
  return (value) => {
    if (
      typeof value !== 'string' ||
      value.length < (options.minLength ?? 0) ||
      value.length > options.maxLength
    )
      throw new Error('invalid_tool_input')
    return value
  }
}

export function optionalField<T>(parser: FieldParser<T>): FieldParser<T | undefined> {
  return (value) => (value === undefined ? undefined : parser(value))
}

export function integerField(options: { min: number; max: number }): FieldParser<number> {
  return (value) => {
    if (
      !Number.isInteger(value) ||
      (value as number) < options.min ||
      (value as number) > options.max
    ) {
      throw new Error('invalid_tool_input')
    }
    return value as number
  }
}

export function exactObject<T extends Record<string, FieldParser<unknown>>>(fields: T) {
  return (input: unknown): { [K in keyof T]: ReturnType<T[K]> } => {
    if (!input || typeof input !== 'object' || Array.isArray(input))
      throw new Error('invalid_tool_input')
    const source = input as Record<string, unknown>
    if (Object.keys(source).some((key) => !Object.hasOwn(fields, key)))
      throw new Error('invalid_tool_input')
    const result: Record<string, unknown> = {}
    for (const [key, parser] of Object.entries(fields)) result[key] = parser(source[key])
    return result as { [K in keyof T]: ReturnType<T[K]> }
  }
}
