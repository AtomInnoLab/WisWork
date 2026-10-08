import { expect, it } from 'vitest'
import { presentationMutationScope } from '../src/skills/powerpoint/presentation-mutation-scope'
function proposal(operation: string, ids = ['old-host']) {
  return {
    operation,
    toolName: operation,
    impact: { host: 'powerpoint', count: 1, targets: ids },
    preview: { qaScope: { basis: 'package_xml_savepoint', hostSlideIds: ids } },
  } as any
}
it.each([
  'edit_slide_xml',
  'edit_slide_chart',
  'resume_package_xml_change',
  'undo_package_xml_change',
  'discard_package_xml_change',
])('limits %s QA invalidation to proven actual XML savepoint identities', (operation) => {
  expect(presentationMutationScope(proposal(operation))).toEqual(['old-host'])
  expect(presentationMutationScope(proposal(operation, ['old-host', 'new-host']))).toEqual([
    'old-host',
    'new-host',
  ])
})
it('keeps unknown XML labels and forged or mismatched identity metadata at full scope', () => {
  for (const p of [
    proposal('execute_office_js'),
    proposal('edit_slide_master_xml'),
    { ...proposal('edit_slide_xml'), toolName: 'different' },
    { ...proposal('edit_slide_xml'), preview: {} },
    {
      ...proposal('edit_slide_xml'),
      preview: { qaScope: { basis: 'other', hostSlideIds: ['old-host'] } },
    },
    {
      ...proposal('edit_slide_xml'),
      preview: { qaScope: { basis: 'package_xml_savepoint', hostSlideIds: ['different'] } },
    },
    proposal('edit_slide_xml', ['bad\n']),
  ])
    expect(presentationMutationScope(p)).toBeUndefined()
})
