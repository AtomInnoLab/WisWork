/** Exact calendar dates can be ordered; period labels remain incomparable. */
function calendarDate(value: string): string | undefined {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return undefined
  const year = Number(value.slice(0, 4))
  const month = Number(value.slice(5, 7))
  const day = Number(value.slice(8, 10))
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0)
  const days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
  return year >= 1 && month >= 1 && month <= 12 && day >= 1 && day <= days[month - 1]!
    ? value
    : undefined
}

export function sourceAsOfFinding(
  claimAsOf: string | undefined,
  sourceAsOf: string | undefined,
): 'source_as_of_missing' | 'source_as_of_earlier' | 'source_as_of_differs' | undefined {
  const claim = claimAsOf?.replace(/\s+/g, ' ').trim()
  if (!claim) return undefined
  const source = sourceAsOf?.replace(/\s+/g, ' ').trim()
  if (!source) return 'source_as_of_missing'
  if (source === claim) return undefined
  const claimDate = calendarDate(claim)
  const sourceDate = calendarDate(source)
  return claimDate && sourceDate && sourceDate < claimDate
    ? 'source_as_of_earlier'
    : 'source_as_of_differs'
}
