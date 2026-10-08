import type { Slide, SlideElement } from '@wiswork/pptx-engine'

/** Read-only, recursively bounded runtime-id to durable creation-id projection for QC. */
export function buildQualityIdentityMap(slide: Slide): {
  slideId: string
  elementIds: Record<string, string>
  truncated: boolean
} {
  const elementIds: Record<string, string> = Object.create(null)
  let visited = 0
  let truncated = false
  const visit = (elements: SlideElement[]) => {
    for (const element of elements) {
      if (visited >= 500) {
        truncated = true
        return
      }
      visited += 1
      if (element.creationId) elementIds[element.id] = element.creationId
      if (element.type === 'group') visit(element.children)
      if (truncated) return
    }
  }
  visit(slide.elements)
  return { slideId: slide.durableId, elementIds, truncated }
}
