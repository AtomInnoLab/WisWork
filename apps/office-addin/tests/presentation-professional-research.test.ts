import { expect, it } from 'vitest'
import { presentationResearchMarkdown } from '../src/skills/powerpoint/presentation-research'
import { researchRecord } from './presentation-research-fixture'
it('exports complete declared professional context with missing fields and no authenticity claim', () => {
  const record = researchRecord()
  record.draft.facts[0]!.professionalContext = {
    domain: 'law',
    materialKind: 'case',
    jurisdiction: '测试法域',
    effectiveFrom: '2020-01-01',
    effectiveUntil: '2026-01-01',
    applicabilityDate: '2026-09-29',
    originalLocation: '第<3>段',
    caseNumber: '案号示例',
    limitations: '需专家复核',
  }
  const md = presentationResearchMarkdown(record)
  for (const value of [
    '专业上下文',
    'caseNumber',
    '案号示例',
    'effectiveUntil',
    '2026-01-01',
    '第&lt;3&gt;段',
    'effectLevel',
    '不代表专业认证',
  ])
    expect(md).toContain(value)
})
it('leaves old Markdown without context details', () =>
  expect(presentationResearchMarkdown(researchRecord())).not.toContain('专业上下文'))
