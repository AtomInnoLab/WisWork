import type { PowerPointMasterOperation } from './browser-powerpoint-adapter.js'

export interface PowerPointStyleDependencies {
  slides: Array<{ slideId: string; masterId: string; layoutId: string }>
}

export function parsePowerPointStyleDependencies(value: unknown): PowerPointStyleDependencies {
  const fail = (): never => {
    throw new Error('office_read_failed')
  }
  const record = (value: unknown, keys: string[]): Record<string, unknown> => {
    if (!value || typeof value !== 'object' || Array.isArray(value)) return fail()
    const item = value as Record<string, unknown>
    if (Object.keys(item).length !== keys.length || !keys.every((key) => Object.hasOwn(item, key)))
      return fail()
    return item
  }
  const input = record(value, ['slides'])
  // Bound the complete serialized graph instead of silently limiting its page count.
  if (
    !Array.isArray(input.slides) ||
    new TextEncoder().encode(JSON.stringify(input)).byteLength > 8 * 1024 * 1024
  )
    return fail()
  const ids = new Set<string>()
  const slides = Array.from(input.slides, (raw) => {
    const item = record(raw, ['slideId', 'masterId', 'layoutId'])
    const id = (value: unknown): string => {
      if (
        typeof value !== 'string' ||
        !value.trim() ||
        value.length > 256 ||
        [...value].some(
          (char) =>
            char.charCodeAt(0) < 32 || (char.charCodeAt(0) >= 127 && char.charCodeAt(0) <= 159),
        )
      )
        return fail()
      return value
    }
    const slide = {
      slideId: id(item.slideId),
      masterId: id(item.masterId),
      layoutId: id(item.layoutId),
    }
    if (ids.has(slide.slideId)) return fail()
    ids.add(slide.slideId)
    return slide
  })
  slides.sort((a, b) => (a.slideId < b.slideId ? -1 : a.slideId > b.slideId ? 1 : 0))
  return { slides }
}

export function affectedStyleSlideIds(
  snapshot: PowerPointStyleDependencies,
  operations: readonly PowerPointMasterOperation[],
): string[] {
  return parsePowerPointStyleDependencies(snapshot)
    .slides.filter((slide) =>
      operations.some(
        (operation) =>
          slide.masterId === operation.master_id &&
          (operation.op !== 'set_layout_background_following' ||
            slide.layoutId === operation.layout_id),
      ),
    )
    .map((slide) => slide.slideId)
}
