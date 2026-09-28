import { expect, it } from 'vitest'
import { acpPresentationStage, acpToolActivity } from '../src/acp-events.js'

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
