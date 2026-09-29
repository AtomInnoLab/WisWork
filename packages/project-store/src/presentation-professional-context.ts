/** Declared professional context; completeness never authenticates facts or applicability. */
type Common = { limitations?: string }
export type PresentationProfessionalContext =
  | (Common & {
      domain: 'science'
      materialKind?: 'paper' | 'dataset' | 'standard' | 'institution'
      publicationId?: string
      version?: string
      sample?: string
      method?: string
      statisticalBasis?: string
    })
  | (Common & {
      domain: 'law'
      materialKind?: 'statute' | 'case' | 'regulation' | 'contract'
      jurisdiction?: string
      effectLevel?: string
      effectiveFrom?: string
      effectiveUntil?: string
      applicabilityDate?: string
      caseNumber?: string
      originalLocation?: string
    })
  | (Common & {
      domain: 'finance'
      materialKind?: 'disclosure' | 'financial_statement' | 'ir' | 'market_data'
      reportingPeriod?: string
      asOf?: string
      currency?: string
      unit?: string
      accountingBasis?: string
      formula?: string
    })
const fields = {
  science: ['publicationId', 'version', 'sample', 'method', 'statisticalBasis', 'limitations'],
  law: [
    'jurisdiction',
    'effectLevel',
    'effectiveFrom',
    'effectiveUntil',
    'applicabilityDate',
    'caseNumber',
    'originalLocation',
    'limitations',
  ],
  finance: [
    'reportingPeriod',
    'asOf',
    'currency',
    'unit',
    'accountingBasis',
    'formula',
    'limitations',
  ],
}
const kinds = {
  science: ['paper', 'dataset', 'standard', 'institution'],
  law: ['statute', 'case', 'regulation', 'contract'],
  finance: ['disclosure', 'financial_statement', 'ir', 'market_data'],
}
const dates = new Set(['effectiveFrom', 'effectiveUntil', 'applicabilityDate', 'asOf'])
export const PROFESSIONAL_CONTEXT_SCHEMA = {
  anyOf: Object.entries(fields).map(([domain, keys]) => ({
    type: 'object',
    required: ['domain'],
    additionalProperties: false,
    properties: {
      domain: { type: 'string', enum: [domain] },
      materialKind: { type: 'string', enum: kinds[domain as keyof typeof kinds] },
      ...Object.fromEntries(
        keys.map((key) => [
          key,
          {
            type: 'string',
            minLength: 1,
            maxLength: dates.has(key) ? 10 : 800,
            ...(dates.has(key) ? { pattern: '^[0-9]{4}-[0-9]{2}-[0-9]{2}$' } : {}),
          },
        ]),
      ),
    },
  })),
}
function invalid(): never {
  throw new Error('professional_context_invalid')
}
export function parsePresentationProfessionalContext(
  value: unknown,
): PresentationProfessionalContext {
  if (!value || typeof value !== 'object' || Array.isArray(value)) invalid()
  const v = value as Record<string, unknown>,
    domain = v.domain
  if (typeof domain !== 'string' || !Object.hasOwn(fields, domain)) invalid()
  const allowed = fields[domain as keyof typeof fields]
  for (const [key, item] of Object.entries(v)) {
    if (key === 'domain') continue
    if (key === 'materialKind') {
      if (typeof item !== 'string' || !kinds[domain as keyof typeof kinds].includes(item)) invalid()
      continue
    }
    if (!allowed.includes(key) || typeof item !== 'string' || !item.trim() || item.length > 800)
      invalid()
    // eslint-disable-next-line no-control-regex
    if (/[^\u0009\u000a\u000d\u0020-\ud7ff\ue000-\ufffd\u{10000}-\u{10ffff}]/u.test(item)) invalid()
    if (
      dates.has(key) &&
      (!/^\d{4}-\d{2}-\d{2}$/.test(item) ||
        !Number.isFinite(Date.parse(item)) ||
        new Date(item).toISOString().slice(0, 10) !== item)
    )
      invalid()
  }
  if (
    domain === 'law' &&
    typeof v.effectiveFrom === 'string' &&
    typeof v.effectiveUntil === 'string' &&
    v.effectiveUntil < v.effectiveFrom
  )
    invalid()
  if (new TextEncoder().encode(JSON.stringify(value)).length > 16 * 1024) invalid()
  return structuredClone(value) as PresentationProfessionalContext
}
export function presentationProfessionalContextMissingFields(
  value: PresentationProfessionalContext,
  claimType?: string,
): string[] {
  const c = parsePresentationProfessionalContext(value)
  const required =
    c.domain === 'science'
      ? ['materialKind', ...fields.science]
      : c.domain === 'law'
        ? [
            'materialKind',
            'jurisdiction',
            'effectLevel',
            'applicabilityDate',
            'originalLocation',
            'limitations',
            ...(c.materialKind !== 'contract' ? ['effectiveFrom'] : []),
            ...(c.materialKind === 'case' ? ['caseNumber'] : []),
          ]
        : [
            'materialKind',
            'reportingPeriod',
            'asOf',
            'currency',
            'unit',
            'accountingBasis',
            'limitations',
            ...(claimType === 'calculation' ? ['formula'] : []),
          ]
  return required.filter((key) => !Object.hasOwn(c, key))
}
