import { expect, it } from 'vitest'

it('retains generic activity outside presentation context', () => {
  for (const name of ['execute_office_js', 'web_search', 'web_fetch', 'image_search']) {
    expect(acpToolActivity(name, 'running', false)).toBe('正在准备修改…')
    expect(acpToolActivity(name, 'complete', false)).toBe('已准备修改')
    expect(acpToolActivity(name, 'error', false)).toBe('准备修改未完成')
  }
})
import {
  acpPresentationStage,
  acpPresentationStageLabel,
  acpToolActivity,
} from '../src/acp-events.js'

it('covers finite dynamically registered lifecycle actions', () => {
  for (const action of [
    'stage',
    'reapply',
    'inspect',
    'reconcile',
    'resume',
    'commit',
    'discard',
    'undo',
    'release',
  ])
    expect(acpPresentationStage(`${action}_existing_presentation_page_change`)).toBe('editing')
  for (const action of ['capture', 'record'])
    expect(acpPresentationStage(`${action}_existing_presentation_page_change`)).toBe('review')
  for (const action of ['inspect', 'resume', 'undo', 'release', 'reapply'])
    expect(acpPresentationStage(`${action}_slide_chart_values_change`)).toBe('editing')
  for (const action of ['start', 'read', 'pause', 'resume', 'cancel'])
    expect(acpPresentationStage(`${action}_presentation_production_job`)).toBe('production')
})

it('shares stage labels while reporting only tool-operation completion', () => {
  expect(acpPresentationStageLabel('planning')).toBe('演示文稿计划')
  expect(acpPresentationStageLabel('baseline')).toBe('文档基线')
  for (const tool of [
    'web_fetch',
    'save_presentation_brand_kit',
    'image_search',
    'read_presentation_baseline',
    'export_presentation_pdf',
    'save_presentation_page_backup',
  ]) {
    const stage = acpPresentationStage(tool)!
    expect(acpToolActivity(tool, 'complete')).toBe(`${acpPresentationStageLabel(stage)}操作已结束`)
    expect(acpToolActivity(tool, 'error')).toBe(`${acpPresentationStageLabel(stage)}处理未完成`)
  }
})

it.each([
  ['import_presentation_production', 'import'],
  ['capture_existing_presentation_batch_page', 'review'],
  ['undo_slide_chart_values_change', 'editing'],
  ['prepare_existing_presentation_composite_revision', 'editing'],
  ['read_presentation_baseline_chart_source', 'baseline'],
  ['web_search', 'research'],
  ['build_research_ledger', 'research'],
  ['read_research_ledger', 'research'],
  ['list_research_ledgers', 'research'],
  ['export_research_ledger', 'research'],
  ['image_search', 'assets'],
  ['save_presentation_brand_kit', 'style'],
  ['save_presentation_page_backup', 'checkpoint'],
  ['export_presentation_pdf', 'delivery'],
  ['export_current_presentation_bundle', 'delivery'],
  ['restore_presentation_delivery_bundle', 'delivery'],
])('classifies registered %s as %s', (tool, stage) => {
  expect(acpPresentationStage(tool)).toBe(stage)
})

it('does not infer a stage for unregistered presentation write names', () => {
  expect(acpPresentationStage('replace_unknown_presentation_page')).toBeUndefined()
  expect(acpPresentationStage('undo_existing_presentation_unknown_change')).toBeUndefined()
})

it('classifies persisted presentation work stages', () => {
  expect(acpPresentationStage('save_presentation_plan')).toBe('planning')
  expect(acpPresentationStage('run_presentation_production')).toBe('production')
  expect(acpPresentationStage('import_generated_presentation')).toBe('import')
  expect(acpPresentationStage('capture_presentation_page_qa')).toBe('review')
  expect(acpPresentationStage('replace_presentation_page')).toBe('editing')
  expect(acpPresentationStage('read_document')).toBeUndefined()
})

it('maps PowerPoint tools to bounded user-facing stages without claiming completion', () => {
  expect(acpToolActivity('save_presentation_plan', 'complete')).toBe('演示文稿计划操作已结束')
  expect(acpToolActivity('run_presentation_production', 'running')).toBe('正在处理逐页制作…')
  expect(acpToolActivity('import_generated_presentation', 'error')).toBe('页面导入处理未完成')
  expect(acpToolActivity('capture_presentation_page_qa', 'complete')).toBe('页面审查操作已结束')
  expect(acpToolActivity('read_presentation_delivery_report', 'complete')).toBe(
    '内容证据操作已结束',
  )
  expect(acpToolActivity('list_presentation_attachments', 'running')).toBe('正在处理参考资料…')
  expect(acpToolActivity('restore_presentation_project', 'complete')).toBe('项目恢复操作已结束')
  expect(acpToolActivity('read_document', 'running')).toBe('正在读取内容…')
  expect(acpToolActivity('write_document', 'error')).toBe('准备修改未完成')
})

it.each([
  ['begin_presentation_edit_observation', 'baseline'],
  ['complete_presentation_edit_observation', 'baseline'],
  ['read_presentation_edit_observation', 'baseline'],
  ['list_presentation_edit_observations', 'baseline'],
  ['delete_presentation_edit_observation', 'baseline'],
  ['save_presentation_observed_preference', 'planning'],
] as const)(
  'maps %s to its user-visible phase without claiming host edits or QA',
  (tool, stage) => {
    expect(acpPresentationStage(tool)).toBe(stage)
    expect(acpToolActivity(tool, 'complete')).toBe(`${acpPresentationStageLabel(stage)}操作已结束`)
    expect(acpPresentationStage(`${tool}_unknown`)).toBeUndefined()
  },
)

it.each([
  ['inspect_native_modify_batch', 'editing'],
  ['resume_native_modify_batch', 'editing'],
  ['finalize_native_modify_restore', 'editing'],
  ['capture_native_modify_page', 'review'],
  ['record_native_modify_page_review', 'review'],
])('classifies the durable generic edit tool %s without claiming host QA', (tool, stage) => {
  expect(acpPresentationStage(tool)).toBe(stage)
  expect(acpToolActivity(tool, 'complete')).toBe(
    `${acpPresentationStageLabel(stage as 'editing' | 'review')}操作已结束`,
  )
})
it('does not infer unregistered generic edit action suffixes', () => {
  expect(acpPresentationStage('approve_native_modify_batch')).toBeUndefined()
  expect(acpPresentationStage('capture_native_modify_page_untrusted')).toBeUndefined()
})
