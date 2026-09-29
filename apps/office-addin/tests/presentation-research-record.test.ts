import { expect, it } from 'vitest'
import { verifyPresentationResearchRecord } from '../src/skills/powerpoint/presentation-research'
import { researchRecord } from './presentation-research-fixture'

it('verifies the full draft digest without aliasing the original record', async () => {
  const input = researchRecord()
  const value = await verifyPresentationResearchRecord(input)
  expect(value).toEqual(input)
  value.draft.scope = 'caller change'
  expect(input.draft.scope).toBe('销售趋势研究')
})
it('rejects a shape-valid record whose original draft has been changed', async () => {
  const input = researchRecord()
  input.draft.scope = 'changed original'
  await expect(verifyPresentationResearchRecord(input)).rejects.toThrow(
    'presentation_response_invalid',
  )
})
it('uses a stable safe error for malformed terminal states', async () => {
  const input = researchRecord()
  input.state = 'running'
  await expect(verifyPresentationResearchRecord(input)).rejects.toThrow(
    'presentation_response_invalid',
  )
})
it('preserves an aborted record and its partial evidence without certifying support', async () => {
  const input = researchRecord()
  input.state = 'failed'
  input.error = 'aborted'
  expect(await verifyPresentationResearchRecord(input)).toEqual(input)
  expect(input.checks.support).toBe('not_verified')
})
