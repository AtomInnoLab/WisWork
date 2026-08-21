import { describe, expect, it } from 'vitest'
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'

const inventories = {
  word: [
    'bash',
    'execute_office_js',
    'get_document_structure',
    'get_document_text',
    'get_ooxml',
    'read',
    'screenshot_document',
  ],
  excel: [
    'bash',
    'clear_cell_range',
    'copy_to',
    'eval_officejs',
    'get_all_objects',
    'get_cell_ranges',
    'get_range_as_csv',
    'modify_object',
    'modify_sheet_structure',
    'modify_workbook_structure',
    'read',
    'resize_range',
    'screenshot_range',
    'search_data',
    'set_cell_range',
  ],
  powerpoint: [
    'bash',
    'duplicate_slide',
    'edit_slide_chart',
    'edit_slide_master',
    'edit_slide_text',
    'edit_slide_xml',
    'execute_office_js',
    'list_slide_shapes',
    'read',
    'read_slide_text',
    'screenshot_slide',
    'verify_slides',
  ],
} as const

describe('host runtime composition', () => {
  it.each(Object.entries(inventories))(
    'composes shared tools with only the %s host skill',
    (host, expected) => {
      const runtime = createOfficeHostRuntime(host as keyof typeof inventories)
      expect(runtime.skill.tools.map((tool) => tool.name).sort()).toEqual(expected)
      expect(runtime.vfs).toBeDefined()
      expect(runtime.skills.list()).toEqual([])
    },
  )

  it('fails closed for an unsupported host', () => {
    expect(() => createOfficeHostRuntime('unknown')).toThrow('office_host_unsupported')
  })

  it('clears disposable VFS, installed skills, and proposals on dispose', () => {
    const runtime = createOfficeHostRuntime('word')
    runtime.vfs.writeFile('/home/user/private.txt', 'secret')
    runtime.skills.install('---\nname: demo\ndescription: demo skill\n---\nInstructions')
    runtime.dispose()
    expect(runtime.vfs.list('/home/user')).toEqual([])
    expect(runtime.vfs.list('/home/skills')).toEqual([])
    expect(runtime.skills.list()).toEqual([])
    expect(runtime.proposals.pending()).toBeUndefined()
  })

  it('keeps the legacy selection skill as an explicit rollback gate', () => {
    const runtime = createOfficeHostRuntime('excel', { enableHostSkills: false })
    expect(runtime.skill.tools.map((tool) => tool.name).sort()).toEqual([
      'propose_append_text',
      'propose_replace_selection',
      'read_selection',
    ])
  })
})
