export const MAX_DECLARATIVE_PROGRAM_BYTES = 32 * 1024
export const MAX_DECLARATIVE_OPERATIONS = 32

export interface DeclarativeProgram<T> {
  version: 1
  operations: T[]
}

export function parseDeclarativeProgram<T>(
  source: string,
  parseOperation: (value: unknown) => T,
): DeclarativeProgram<T> {
  if (new TextEncoder().encode(source).byteLength > MAX_DECLARATIVE_PROGRAM_BYTES)
    throw new Error('invalid_tool_input')
  let value: unknown
  try {
    value = JSON.parse(source)
  } catch {
    throw new Error('invalid_tool_input')
  }
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('invalid_tool_input')
  const record = value as Record<string, unknown>
  if (
    Object.keys(record).some((key) => !['version', 'operations'].includes(key)) ||
    record.version !== 1 ||
    !Array.isArray(record.operations) ||
    record.operations.length < 1 ||
    record.operations.length > MAX_DECLARATIVE_OPERATIONS
  )
    throw new Error('invalid_tool_input')
  return { version: 1, operations: record.operations.map(parseOperation) }
}
