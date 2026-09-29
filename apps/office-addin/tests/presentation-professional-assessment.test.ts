import { expect, it, vi } from 'vitest'
import { createPresentationProductionSkill } from '../src/skills/powerpoint/presentation-production.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'

const context = {
  domain: 'science',
  materialKind: 'paper',
  publicationId: 'doi:example',
  sample: 'adults',
  method: 'observational',
  statisticalBasis: 'association',
  limitations: 'no causality',
}
const input = {
  project_id: 'p',
  request_id: 'r',
  page_id: 'one',
  claim_id: 'c',
  source_id: 's',
  offset: 2,
  max_chars: 8,
}
const read = { id: 'read', name: 'read_presentation_claim_evidence', input }
const assessment = () => ({
  scope: 'This read window only',
  authority: { outcome: 'uncertain', sourceTier: 'unverified', reason: 'not authenticated' },
  timeliness: { outcome: 'uncertain', referenceDate: '2026-09-29', reason: 'not authenticated' },
  basis: [{ offset: 4, text: 'fact' }],
  professional: {
    context: structuredClone(context),
    checks: [
      { aspect: 'conclusion_scope', outcome: 'consistent', reason: 'limited to association' },
      {
        aspect: 'qualifications',
        outcome: 'uncertain',
        reason: 'limitations require further reading',
      },
    ],
  },
})
function fixture(professional = true, responseChange?: (v: Record<string, unknown>) => void) {
  const evidence = {
    version: 1,
    projectId: 'p',
    requestId: 'r',
    planRevision: 1,
    inputDigest: 'a'.repeat(64),
    planDigest: 'b'.repeat(64),
    pageId: 'one',
    claimId: 'c',
    statement: 'fact',
    ...(professional
      ? {
          documentId: 'doc',
          claim: {
            id: 'c',
            statement: 'fact',
            type: 'fact',
            sourceIds: ['s'],
            confidence: 'low',
            reviewStatus: 'needs_review',
            professionalContext: context,
          },
        }
      : {}),
    source: { id: 's', uri: `attachment:${'c'.repeat(64)}`, excerpt: 'fact', locator: 'section 1' },
    attachment: {
      id: 'c'.repeat(64),
      name: 'source.txt',
      offset: 2,
      totalChars: 10,
      text: 'a fact z',
      offsetUnit: 'utf16_code_unit',
    },
    excerptMatch: { status: 'found', offset: 4 },
    checks: {
      support: 'not_verified',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
      host: 'not_checked',
    },
  }
  const request = vi.fn(async (body: unknown) => {
    const b = body as Record<string, unknown>
    if (b.operation === 'production_claim_evidence') return new Response(JSON.stringify(evidence))
    const response: Record<string, unknown> = {
      version: 1,
      projectId: 'p',
      requestId: 'r',
      reviewId: 'review',
      planRevision: 1,
      inputDigest: evidence.inputDigest,
      planDigest: evidence.planDigest,
      pageId: 'one',
      claimId: 'c',
      sourceId: 's',
      attachmentId: evidence.attachment.id,
      offset: 2,
      maxChars: 8,
      evidenceDigest: b.evidenceDigest,
      outcome: b.outcome,
      notes: b.notes,
      ...(b.sourceAssessment ? { sourceAssessment: structuredClone(b.sourceAssessment) } : {}),
      reviewer: 'agent',
      createdAt: '2026-09-29T00:00:00.000Z',
      checks: {
        support: 'agent_reviewed',
        sourceAuthority: 'not_verified',
        timeliness: 'not_verified',
        host: 'not_checked',
      },
    }
    responseChange?.(response)
    return new Response(JSON.stringify(response))
  })
  const skill = createPresentationProductionSkill({
    vfs: new InMemoryVfs(),
    request,
    documentId: async () => 'doc',
    available: () => true,
    lastProject: () => 'p',
    rememberProject: async () => {},
    readReceipt: () => undefined,
  })
  const review = (sourceAssessment?: unknown) =>
    skill.executeTool({
      id: 'review',
      name: 'record_presentation_claim_review',
      input: {
        ...input,
        review_id: 'review',
        outcome: 'supported',
        notes: 'Scoped judgment only',
        ...(sourceAssessment ? { source_assessment: sourceAssessment } : {}),
      },
    })
  return { skill, request, review }
}

it('requires actual evidence reading and preserves full professional assessment without certification', async () => {
  const f = fixture()
  expect(await f.review(assessment())).toMatchObject({
    isError: true,
    output: 'presentation_evidence_read_required',
  })
  expect(f.request).not.toHaveBeenCalled()
  const readResult = await f.skill.executeTool(read)
  expect(readResult.isError, readResult.output).not.toBe(true)
  const result = await f.review(assessment())
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({
    sourceAssessment: assessment(),
    checks: { sourceAuthority: 'not_verified', timeliness: 'not_verified' },
  })
})

it('rejects wrong domain, omitted or changed declared context before sending the review', async () => {
  for (const changed of [
    { domain: 'law' },
    { domain: 'science' },
    { ...context, sample: 'children' },
  ]) {
    const f = fixture()
    expect((await f.skill.executeTool(read)).isError).not.toBe(true)
    const a = assessment()
    Object.assign(a.professional, { context: changed })
    expect(await f.review(a)).toMatchObject({ isError: true, output: 'source_assessment_invalid' })
    expect(f.request).toHaveBeenCalledTimes(1)
  }
})

it('rejects professional assessment without frozen context and keeps legacy absent shape', async () => {
  const f = fixture(false)
  expect((await f.skill.executeTool(read)).isError).not.toBe(true)
  expect(await f.review(assessment())).toMatchObject({
    isError: true,
    output: 'source_assessment_invalid',
  })
  const result = await f.review()
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).not.toHaveProperty('sourceAssessment')
})

it('requires non-uncertain professional basis and exact literal basis in the actual window', async () => {
  for (const basis of [[], [{ offset: 3, text: 'fact' }], [{ offset: 20, text: 'fact' }]]) {
    const f = fixture()
    expect((await f.skill.executeTool(read)).isError).not.toBe(true)
    const a = assessment()
    a.basis = basis
    expect(await f.review(a)).toMatchObject({ isError: true, output: 'source_assessment_invalid' })
    expect(f.request).toHaveBeenCalledTimes(1)
  }
})

it('rejects PC omission or replacement of exact nested professional assessment', async () => {
  for (const change of ['omit', 'context', 'reason']) {
    const f = fixture(true, (v) => {
      const a = v.sourceAssessment as ReturnType<typeof assessment>
      if (change === 'omit') delete (a as Partial<typeof a>).professional
      if (change === 'context') a.professional.context.sample = 'children'
      if (change === 'reason') a.professional.checks[0]!.reason = 'different historical opinion'
    })
    expect((await f.skill.executeTool(read)).isError).not.toBe(true)
    expect(await f.review(assessment())).toMatchObject({
      isError: true,
      output: 'presentation_response_invalid',
    })
  }
})
