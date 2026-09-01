import { describe, expect, it, vi } from 'vitest'
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'
import type { StructuredProposalController } from '../src/agent/proposal-controller.js'
import type { ElevatedOfficeAdapter } from '../src/skills/shared/elevated-office-program.js'

const inventories = {
  word: [
    'bash',
    'execute_office_js',
    'get_document_structure',
    'get_document_text',
    'get_ooxml',
    'read',
    'screenshot_document',
    'write_document',
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
    'edit_slide_master_xml',
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
  it('advertises the distinct raw tool only when an Enhanced adapter is supplied', () => {
    const standard = createOfficeHostRuntime('word')
    expect(standard.skill.tools.map((tool) => tool.name)).not.toContain('propose_raw_office_edit')
    const adapter: ElevatedOfficeAdapter = {
      host: 'word',
      captureAuthority: () => ({
        activeMode: 'enhanced',
        signedIn: true,
        paired: true,
        hostEnabled: true,
        rawOfficeEnabled: true,
        rawOfficeJsEnabled: true,
        rawOfficeOoxmlEnabled: true,
        documentId: 'doc_AAAAAAAAAAAAAAAA',
        sessionId: 'ses_AAAAAAAAAAAAAAAA',
        generation: 1,
        revision: 'rev_AAAAAAAAAAAAAAAA',
      }),
      snapshot: async () => ({ id: 'history_AAAAAAAAAAAAAAAA' }),
      validateSnapshot: async () => true,
      execute: async () => undefined,
      readback: async () => ({ verified: true }),
      rollback: async () => undefined,
    }
    const names = createOfficeHostRuntime('word', {
      elevatedOfficeAdapter: adapter,
    }).skill.tools.map((tool) => tool.name)
    expect(names).toContain('execute_office_js')
    expect(names).toContain('propose_raw_office_edit')
    standard.enableElevatedOffice(adapter.captureAuthority)
    expect(standard.skill.tools.map((tool) => tool.name)).toContain('propose_raw_office_edit')
    standard.disableElevatedOffice()
    expect(standard.skill.tools.map((tool) => tool.name)).not.toContain('propose_raw_office_edit')
  })
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

  it('omits unsupported master package editing from the PowerPoint Mac inventory', () => {
    const runtime = createOfficeHostRuntime('powerpoint', { platform: 'Mac' })
    expect(runtime.skill.tools.map((tool) => tool.name)).not.toContain('edit_slide_master')
    expect(runtime.skill.tools.map((tool) => tool.name)).not.toContain('inspect_slide_masters')
    expect(runtime.skill.tools.map((tool) => tool.name)).not.toContain('edit_slide_master_xml')
    expect(runtime.skill.tools.map((tool) => tool.name)).toContain('edit_slide_xml')
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

  it('inherits quarantine across session clearing and releases it only on document disposal', () => {
    const runtime = createOfficeHostRuntime('word')
    const proposals = runtime.proposals as StructuredProposalController
    proposals.quarantine({ sessionId: 'session-a', generation: 1 })

    runtime.clearSession()
    expect(proposals.isQuarantined()).toBe(true)
    runtime.disableElevatedOffice()
    expect(proposals.isQuarantined()).toBe(true)
    runtime.dispose()
    expect(proposals.isQuarantined()).toBe(false)
  })

  it('ignores an upload that settles after the runtime is disposed', async () => {
    const runtime = createOfficeHostRuntime('word')
    let resolve!: (value: ArrayBuffer) => void
    const upload = runtime.uploadFile('late.txt', new Promise((next) => (resolve = next)))
    runtime.dispose()
    resolve(new TextEncoder().encode('secret').buffer as ArrayBuffer)
    await expect(upload).rejects.toThrow('upload_cancelled')
    expect(runtime.vfs.list('/home/user')).toEqual([])
  })

  it.each(['clearSession', 'dispose'] as const)(
    '%s cancels active package parsing and rejects late installation',
    async (action) => {
      let reject!: (error: Error) => void
      const cancelAll = vi.fn(() => reject(new Error('upload_cancelled')))
      const packageRuntime = {
        parse: vi.fn(() => new Promise<never>((_resolve, next) => (reject = next))),
        cancelAll,
      }
      const runtime = createOfficeHostRuntime('word', { packageRuntime })
      const pending = runtime.installSkillPackage(Promise.resolve(new ArrayBuffer(1)))
      await Promise.resolve()
      runtime[action]()
      await expect(pending).rejects.toThrow('upload_cancelled')
      expect(cancelAll).toHaveBeenCalledOnce()
      expect(runtime.skills.list()).toEqual([])
    },
  )

  it('installs a bounded SKILL.md and exposes it dynamically to Agent context', async () => {
    const runtime = createOfficeHostRuntime('word')
    await runtime.installSkill(
      Promise.resolve('---\nname: writer\ndescription: Helps edit prose\n---\nBe concise.'),
    )
    expect(runtime.skills.list()).toMatchObject([{ name: 'writer' }])
    expect(runtime.skill.buildContext?.()).toContain(
      'writer: Helps edit prose (/home/skills/writer/SKILL.md)',
    )
  })

  it('keeps the legacy selection skill as an explicit rollback gate', () => {
    const runtime = createOfficeHostRuntime('excel', { enableHostSkills: false })
    expect(runtime.skill.tools.map((tool) => tool.name).sort()).toEqual([
      'propose_append_text',
      'propose_replace_selection',
      'read_selection',
    ])
  })

  it('independently disables conversion, package, and import/media capability families', async () => {
    const runtime = createOfficeHostRuntime('excel', {
      enableConversions: false,
      enableSkillPackages: false,
      enableImportMedia: false,
    })
    expect(runtime.skill.tools.map((tool) => tool.name)).not.toContain('csv-to-sheet')
    const bash = await runtime.skill.executeTool(
      {
        id: 'disabled-conversion',
        name: 'bash',
        input: { command: 'pdf-to-text /home/user/a.pdf' },
      },
      new AbortController().signal,
    )
    expect(bash.output).toBe('sandbox_denied')
    await expect(runtime.installSkillPackage(Promise.resolve(new ArrayBuffer(1)))).rejects.toThrow(
      'office_capability_disabled',
    )
    await expect(
      runtime.installSkill(Promise.resolve('---\nname: disabled\ndescription: disabled\n---\nNo.')),
    ).rejects.toThrow('office_capability_disabled')
  })
})
