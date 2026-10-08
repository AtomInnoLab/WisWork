/** Office hosts return paragraph and soft-break separators differently. Preserve all other text. */
export function equivalentPowerPointText(actual: string, expected: string): boolean {
  const normalize = (value: string) => value.replace(/\r\n|\r|\v/g, '\n')
  return normalize(actual) === normalize(expected)
}
