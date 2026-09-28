import { parsePresentationDeck, type PresentationDeck } from './presentation'
import {
  assertDeckMatchesPresentationPlan,
  parsePresentationPlan,
  type PresentationPlan,
} from './presentation-plan'
import { array, choice, id, object, valid } from './presentation-schema'
import { sourceAsOfFinding } from './presentation-source-time'

export interface PresentationPageContentCheck {
  version: 1
  pageId: string
  claimIds: string[]
  findings: {
    code:
      | 'claim_text_not_found'
      | 'source_excerpt_missing'
      | 'source_locator_missing'
      | 'quote_not_in_excerpt'
      | 'source_as_of_missing'
      | 'source_as_of_earlier'
      | 'source_as_of_differs'
      | 'calculation_not_reproduced'
    claimId: string
    sourceId?: string
  }[]
  checks: {
    content: 'needs_review'
    sources: 'not_verified'
    calculations: 'not_verified'
    timeliness: 'not_verified'
    host: 'not_checked'
  }
}

const reportSchema = object({
  version: { type: 'number', enum: [1] },
  pageId: id,
  claimIds: array(id, 32),
  findings: array(
    {
      anyOf: [
        object({ code: choice('claim_text_not_found', 'calculation_not_reproduced'), claimId: id }),
        object({
          code: choice(
            'source_excerpt_missing',
            'source_locator_missing',
            'quote_not_in_excerpt',
            'source_as_of_missing',
            'source_as_of_earlier',
            'source_as_of_differs',
          ),
          claimId: id,
          sourceId: id,
        }),
      ],
    },
    352,
  ),
  checks: object({
    content: choice('needs_review'),
    sources: choice('not_verified'),
    calculations: choice('not_verified'),
    timeliness: choice('not_verified'),
    host: choice('not_checked'),
  }),
})

export function parsePresentationPageContentCheck(value: unknown): PresentationPageContentCheck {
  const reject = (reason: string): never => {
    throw new Error(`presentation_content_check_invalid:${reason}`)
  }
  if (!valid(value, reportSchema)) reject('schema')
  const report = value as PresentationPageContentCheck
  const claimIds = new Set(report.claimIds)
  if (claimIds.size !== report.claimIds.length) reject('duplicate_claim')
  const findings = new Set<string>()
  for (const finding of report.findings) {
    if (!claimIds.has(finding.claimId)) reject('claim_reference')
    const key = JSON.stringify([finding.code, finding.claimId, finding.sourceId])
    if (findings.has(key)) reject('duplicate_finding')
    findings.add(key)
  }
  return structuredClone(report)
}

const normalize = (text: string): string => text.replace(/\s+/g, ' ').trim()

/** Literal coverage and evidence completeness only. Never executes formulas or fetches sources. */
export function checkPresentationPageContent(
  inputPlan: PresentationPlan,
  inputDeck: PresentationDeck,
  pageId: string,
): PresentationPageContentCheck {
  const plan = parsePresentationPlan(inputPlan)
  const deck = parsePresentationDeck(inputDeck)
  assertDeckMatchesPresentationPlan(deck, plan)
  const slide = deck.slides.find((page) => page.id === pageId)
  if (!slide) throw new Error('not_found')
  // Metadata titles and notes are not rendered body content; never join separate visible strings.
  const visible = slide.elements
    .flatMap((element): string[] => {
      if (element.kind === 'text') return [element.text]
      if (element.kind === 'table') return element.rows.flat()
      if (element.kind === 'chart') {
        // Match compiler visibility: pie has no axis/legend/category labels;
        // bar/line series names appear only in the multi-series legend.
        if (element.chartType === 'pie') return []
        return [
          ...element.categories,
          ...(element.series.length > 1 ? element.series.map((series) => series.name) : []),
        ]
      }
      return []
    })
    .map(normalize)
  const claimIds = [...(slide.claimIds ?? [])]
  const claims = new Map(plan.claims.map((claim) => [claim.id, claim]))
  const sources = new Map(plan.sources.map((source) => [source.id, source]))
  const findings: PresentationPageContentCheck['findings'] = []
  for (const claimId of claimIds) {
    const claim = claims.get(claimId)!
    const statement = normalize(claim.statement)
    if (!statement || !visible.some((text) => text.includes(statement)))
      findings.push({ code: 'claim_text_not_found', claimId })
    for (const sourceId of claim.sourceIds) {
      const source = sources.get(sourceId)!
      const excerpt = normalize(source.excerpt)
      if (!excerpt) findings.push({ code: 'source_excerpt_missing', claimId, sourceId })
      if (!source.locator?.trim())
        findings.push({ code: 'source_locator_missing', claimId, sourceId })
      const asOfFinding = sourceAsOfFinding(claim.asOf, source.asOf)
      if (asOfFinding) findings.push({ code: asOfFinding, claimId, sourceId })
      // Missing excerpts are already flagged; quote comparison needs supplied text.
      if (claim.type === 'quote' && excerpt && (!statement || !excerpt.includes(statement)))
        findings.push({ code: 'quote_not_in_excerpt', claimId, sourceId })
    }
    if (claim.type === 'calculation') findings.push({ code: 'calculation_not_reproduced', claimId })
  }
  return parsePresentationPageContentCheck({
    version: 1,
    pageId,
    claimIds,
    findings,
    checks: {
      content: 'needs_review',
      sources: 'not_verified',
      calculations: 'not_verified',
      timeliness: 'not_verified',
      host: 'not_checked',
    },
  })
}
