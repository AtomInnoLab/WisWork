import { expect, it } from 'vitest'
import { presentationMutationScope } from '../src/skills/powerpoint/presentation-mutation-scope.js'
const ids = Array.from({ length: 600 }, (_, i) => `actual-${i}`)
const proposal = (operation = 'edit_slide_master_xml') =>
  ({
    operation,
    toolName: operation,
    impact: { host: 'powerpoint', targets: ids, count: ids.length },
    preview: { qaScope: { basis: 'master_xml_savepoint', hostSlideIds: ids } },
  }) as any
it.each([
  'edit_slide_master_xml',
  'resume_master_xml_change',
  'undo_master_xml_change',
  'discard_master_xml_change',
])('uses the complete proven actual scope for %s', (operation) => {
  expect(presentationMutationScope(proposal(operation))).toEqual(ids)
})
it('does not narrow mismatched or generic master XML metadata', () => {
  for (const p of [
    proposal('execute_office_js'),
    { ...proposal(), toolName: 'different' },
    { ...proposal(), impact: { ...proposal().impact, count: 1 } },
    { ...proposal(), preview: { qaScope: { basis: 'package_xml_savepoint', hostSlideIds: ids } } },
    {
      ...proposal(),
      preview: { qaScope: { basis: 'master_xml_savepoint', hostSlideIds: ids.slice(0, 20) } },
    },
    {
      ...proposal(),
      impact: { host: 'powerpoint', targets: ['bad\n'], count: 1 },
      preview: { qaScope: { basis: 'master_xml_savepoint', hostSlideIds: ['bad\n'] } },
    },
  ])
    expect(presentationMutationScope(p)).toBeUndefined()
})
