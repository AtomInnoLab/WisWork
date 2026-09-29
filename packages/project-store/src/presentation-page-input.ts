import { canonicalPresentationValue as canonical } from './presentation-canonical.js'
import type { PresentationPlanBinding } from './presentation-store.js'

type Row = Record<string, unknown>
const object = (value: unknown): value is Row =>
  !!value && typeof value === 'object' && !Array.isArray(value)
const rows = (value: unknown): value is Row[] => Array.isArray(value) && value.every(object)
const ids = (value: unknown): value is string[] =>
  Array.isArray(value) && value.every((id) => typeof id === 'string')

/** Compare the full input of this page and its declared predecessors.
 * Incomplete legacy shapes cannot establish equivalence and are never reused.
 */
export function presentationPageInput(
  deck: unknown,
  binding: PresentationPlanBinding,
  pageId: string,
): string | undefined {
  const plan = binding.plan
  if (!object(deck) || !object(plan)) return
  const { slides, claims, assets, ...sharedDeck } = deck
  const {
    slides: tasks,
    claims: plannedClaims,
    sources,
    parallelism: _parallelism,
    ...sharedPlan
  } = plan
  if (
    !rows(slides) ||
    !rows(claims) ||
    !rows(assets) ||
    !rows(tasks) ||
    !rows(plannedClaims) ||
    !rows(sources)
  )
    return
  if (slides.length > 32 || tasks.length > 32) return
  const active = new Set<string>()
  const inputs = new Map<string, unknown>()
  const visit = (id: string): boolean => {
    if (active.has(id)) return false
    if (inputs.has(id)) return true
    const slide = slides.find((item) => item.id === id)
    const task = tasks.find((item) => item.id === id)
    if (!slide || !task || !ids(slide.claimIds) || !ids(task.claimIds) || !rows(slide.elements))
      return false
    const dependencies = task.dependsOn ?? []
    if (!ids(dependencies)) return false
    active.add(id)
    for (const dependency of dependencies) if (!visit(dependency)) return false
    const pageClaims = slide.claimIds.map((claimId) => claims.find((claim) => claim.id === claimId))
    const pagePlannedClaims = task.claimIds.map((claimId) =>
      plannedClaims.find((claim) => claim.id === claimId),
    )
    if (
      pageClaims.some((claim) => !claim) ||
      pagePlannedClaims.some((claim) => !claim || !ids(claim.sourceIds))
    )
      return false
    const sourceIds = new Set(pagePlannedClaims.flatMap((claim) => claim!.sourceIds as string[]))
    const pageSources = [...sourceIds]
      .sort()
      .map((sourceId) => sources.find((source) => source.id === sourceId))
    const assetIds = new Set(
      slide.elements
        .filter((element) => element.kind === 'image')
        .map((element) => element.assetId),
    )
    const pageAssets = [...assetIds]
      .sort()
      .map((assetId) => assets.find((asset) => asset.id === assetId))
    if (pageSources.some((source) => !source) || pageAssets.some((asset) => !asset)) return false
    inputs.set(id, {
      slide,
      task,
      claims: pageClaims,
      plannedClaims: pagePlannedClaims,
      sources: pageSources,
      assets: pageAssets,
      firstPage: slides[0]?.id === id,
    })
    active.delete(id)
    return true
  }
  if (!visit(pageId)) return
  return canonical({
    sharedDeck,
    sharedPlan,
    pages: [...inputs].sort(([a], [b]) => a.localeCompare(b)),
  })
}
