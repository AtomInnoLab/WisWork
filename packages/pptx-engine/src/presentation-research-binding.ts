import {
  parsePresentationResearchRecord,
  type PresentationResearchRecord,
} from '@wiswork/project-store/presentation-research'
import { canonicalPresentationValue as canonical } from '@wiswork/project-store/presentation-canonical'
import { parsePresentationPlan, type PresentationPlan } from './presentation-plan'

export interface PresentationResearchBindingFinding {
  code:
    'unmapped_claim' | 'omitted_conflict_partner' | 'unselected_source_ref' | 'source_unavailable'
  claimId: string
  researchClaimId?: string
  relatedResearchClaimId?: string
  /** Original research source identity (including sources omitted from the plan). */
  sourceId?: string
}
function invalid(): never {
  throw new Error('research_binding_invalid')
}
const fields = (value: object, keys: string[]) =>
  Object.fromEntries(keys.map((key) => [key, (value as Record<string, unknown>)[key]]))

export function assertPresentationResearchBinding(
  planValue: PresentationPlan,
  recordValue: PresentationResearchRecord,
): void {
  try {
    const plan = parsePresentationPlan(planValue)
    if (!plan.research) return
    const record = parsePresentationResearchRecord(recordValue)
    const binding = plan.research
    if (
      record.state !== 'completed' ||
      record.projectId !== plan.projectId ||
      record.id !== binding.ledgerId ||
      record.sequence !== binding.sequence ||
      record.draftDigest !== binding.draftDigest
    )
      invalid()
    const sourceMappings = new Map(
      binding.sources.map((mapping) => [mapping.sourceId, mapping.researchSourceId]),
    )
    for (const mapping of binding.sources) {
      const source = plan.sources.find((item) => item.id === mapping.sourceId)!
      const original = record.draft.sources.find((item) => item.id === mapping.researchSourceId)
      if (
        !original ||
        canonical(fields(source, ['uri', 'snapshotAttachmentId', 'excerpt', 'locator', 'asOf'])) !==
          canonical(fields(original, ['uri', 'snapshotAttachmentId', 'excerpt', 'locator', 'asOf']))
      )
        invalid()
    }
    for (const mapping of binding.claims) {
      const claim = plan.claims.find((item) => item.id === mapping.claimId)!
      const original = record.draft.facts.find((item) => item.claimId === mapping.researchClaimId)
      if (
        !original ||
        canonical(
          fields(claim, ['statement', 'type', 'asOf', 'jurisdiction', 'professionalContext']),
        ) !==
          canonical(
            fields(original, ['statement', 'type', 'asOf', 'jurisdiction', 'professionalContext']),
          )
      )
        invalid()
      const calculation = (value: typeof claim | typeof original) =>
        value.calculation
          ? fields(value.calculation, ['formula', 'inputs', 'unit', 'currency'])
          : undefined
      if (
        canonical(calculation(claim)) !== canonical(calculation(original)) ||
        claim.sourceIds.some(
          (id) => !sourceMappings.has(id) || !original.sourceRefs.includes(sourceMappings.get(id)!),
        )
      )
        invalid()
    }
  } catch (cause) {
    throw new Error('research_binding_invalid', { cause })
  }
}

export function presentationResearchBindingFindings(
  plan: PresentationPlan,
  recordValue: PresentationResearchRecord,
): PresentationResearchBindingFinding[] {
  assertPresentationResearchBinding(plan, recordValue)
  if (!plan.research) return []
  const record = parsePresentationResearchRecord(recordValue)
  return plan.claims.flatMap((claim) =>
    presentationResearchClaimBindingFindings(claim, plan.research!, record),
  )
}

/** Callers validate their full plan or evidence context before deriving claim findings. */
export function presentationResearchClaimBindingFindings(
  claim: PresentationPlan['claims'][number],
  binding: NonNullable<PresentationPlan['research']>,
  record: PresentationResearchRecord,
): PresentationResearchBindingFinding[] {
  const findings: PresentationResearchBindingFinding[] = []
  const mappedResearchClaims = new Set(binding.claims.map((mapping) => mapping.researchClaimId))
  const sourceMappings = new Map(
    binding.sources.map((mapping) => [mapping.sourceId, mapping.researchSourceId]),
  )
  {
    const mapping = binding.claims.find((item) => item.claimId === claim.id)
    if (!mapping) {
      findings.push({ code: 'unmapped_claim', claimId: claim.id })
      return findings
    }
    const original = record.draft.facts.find((item) => item.claimId === mapping.researchClaimId)!
    const base = { claimId: claim.id, researchClaimId: original.claimId }
    for (const conflict of original.conflictsWith) {
      if (!mappedResearchClaims.has(conflict))
        findings.push({
          code: 'omitted_conflict_partner',
          ...base,
          relatedResearchClaimId: conflict,
        })
    }
    // Some archives record only one direction of a conflict; retain the omitted partner either way.
    for (const other of record.draft.facts) {
      if (
        other.conflictsWith.includes(original.claimId) &&
        !original.conflictsWith.includes(other.claimId) &&
        !mappedResearchClaims.has(other.claimId)
      )
        findings.push({
          code: 'omitted_conflict_partner',
          ...base,
          relatedResearchClaimId: other.claimId,
        })
    }
    const selected = new Set(claim.sourceIds.map((id) => sourceMappings.get(id)!))
    for (const sourceId of original.sourceRefs) {
      if (!selected.has(sourceId))
        findings.push({ code: 'unselected_source_ref', ...base, sourceId })
      if (record.sources?.find((item) => item.sourceId === sourceId)?.status !== 'found')
        findings.push({ code: 'source_unavailable', ...base, sourceId })
    }
  }
  return findings
}
