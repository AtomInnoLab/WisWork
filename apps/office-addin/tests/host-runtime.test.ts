import { describe, expect, it, vi } from 'vitest'
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

describe('presentation capability composition', () => {
  it('exposes generation only after negotiation and removes it after disconnect', async () => {
    let connected = false
    const runtime = createOfficeHostRuntime('powerpoint', {
      presentation: {
        available: () => connected,
        request: vi.fn(),
        documentId: async () => 'document-1',
        lastProject: () => undefined,
        rememberProject: async () => undefined,
      },
    })
    expect(runtime.skill.tools.map((tool) => tool.name)).not.toContain(
      'compile_deck_with_pptxgenjs',
    )
    connected = true
    expect(runtime.skill.tools.map((tool) => tool.name)).toContain('compile_deck_with_pptxgenjs')
    expect(runtime.skill.tools.map((tool) => tool.name)).toContain('restore_presentation_project')
    connected = false
    expect(runtime.skill.tools.map((tool) => tool.name)).not.toContain(
      'compile_deck_with_pptxgenjs',
    )
    expect(
      await runtime.skill.executeTool({
        id: 'stale',
        name: 'compile_deck_with_pptxgenjs',
        input: {},
      }),
    ).toMatchObject({ isError: true, output: 'presentation_unavailable' })
    runtime.dispose()
  })
})

describe('presentation project runtime lifecycle', () => {
  it('exposes recovery controls and invalidates pending status on clear', async () => {
    let finish!: (response: Response) => void
    const request = vi.fn(
      () =>
        new Promise<Response>((resolve) => {
          finish = resolve
        }),
    )
    const runtime = createOfficeHostRuntime('powerpoint', {
      presentation: {
        available: () => true,
        request,
        documentId: async () => 'document-1',
        lastProject: () => 'project-1',
        rememberProject: async () => undefined,
      },
    })
    expect(runtime.presentation).toBeDefined()
    expect(runtime.skill.tools.map((tool) => tool.name)).toContain('resume_presentation_project')
    const pending = runtime.presentation!.refresh()
    await vi.waitFor(() => expect(request).toHaveBeenCalled())
    runtime.clearSession()
    finish(new Response('{}'))
    await pending
    expect(runtime.presentation!.snapshot()).toEqual({ phase: 'idle' })
    runtime.dispose()
  })
})

it('composes saved planning tools and clears their asynchronous state with the session', async () => {
  let finish!: (response: Response) => void
  const request = vi.fn(
    () =>
      new Promise<Response>((resolve) => {
        finish = resolve
      }),
  )
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      request,
      documentId: async () => 'doc-1',
      lastProject: () => 'project-1',
      rememberProject: async () => {},
    },
  })
  expect(runtime.skill.tools.map((tool) => tool.name)).toContain('save_presentation_plan')
  expect(runtime.skill.systemPrompt).toContain('needs_review')
  const pending = runtime.skill.executeTool({
    id: 'read',
    name: 'read_presentation_plan',
    input: {},
  })
  await vi.waitFor(() => expect(request).toHaveBeenCalled())
  runtime.clearSession()
  finish(new Response('{}'))
  expect(await pending).toMatchObject({ isError: true, output: 'cancelled' })
  expect(runtime.vfs.list('/home/user')).toEqual([])
  runtime.dispose()
})
it('routes supported attachments to PC and cancels outstanding reads on disposal', async () => {
  const request = vi.fn(
    async (_body: unknown, _signal?: AbortSignal) =>
      new Response(JSON.stringify({ attachments: [] })),
  )
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      attachmentsAvailable: () => true,
      request,
      documentId: async () => 'doc',
      lastProject: () => undefined,
      rememberProject: async () => {},
    },
  })
  expect(runtime.durableAttachmentsAvailable?.()).toBe(true)
  expect(runtime.skill.tools.map((t) => t.name)).toContain('list_presentation_attachments')
  expect(
    await runtime.skill.executeTool({
      id: 'list',
      name: 'list_presentation_attachments',
      input: {},
    }),
  ).toMatchObject({ output: '{"attachments":[]}', mutated: false })
  await runtime.uploadFile('image.png', Promise.resolve(new Uint8Array([1]).buffer))
  expect(request).toHaveBeenCalledTimes(1)
  request.mockImplementation(async (_body, signal) => {
    runtime.dispose()
    expect(signal?.aborted).toBe(true)
    return new Response('{"attachments":[]}')
  })
  expect(
    await runtime.skill.executeTool({
      id: 'list',
      name: 'list_presentation_attachments',
      input: {},
    }),
  ).toMatchObject({ isError: true, output: 'upload_cancelled' })
  expect(runtime.vfs.list('/home/user')).toEqual([])
  await expect(
    runtime.uploadFile('notes.txt', Promise.resolve(new ArrayBuffer(0))),
  ).rejects.toThrow('upload_cancelled')
})
it('keeps older presentation-only PCs on local attachment behavior', async () => {
  const request = vi.fn()
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      request,
      documentId: async () => 'doc',
      lastProject: () => undefined,
      rememberProject: async () => {},
    },
  })
  expect(runtime.durableAttachmentsAvailable?.()).toBe(false)
  expect(runtime.skill.tools.map((t) => t.name)).not.toContain('list_presentation_attachments')
  expect(runtime.skill.tools.map((t) => t.name)).toContain('compile_deck_with_pptxgenjs')
  await runtime.uploadFile('source.txt', Promise.resolve(new Uint8Array([65]).buffer))
  expect(request).not.toHaveBeenCalled()
  expect(runtime.vfs.list('/home/user')).toEqual(['/home/user/source.txt'])
  runtime.dispose()
})
it('retains local image uploads on PCs with document attachments but no asset capability', async () => {
  const request = vi.fn()
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      attachmentsAvailable: () => true,
      request,
      documentId: async () => 'doc',
      lastProject: () => undefined,
      rememberProject: async () => {},
    },
  })
  expect(runtime.durableImagesAvailable?.()).toBe(false)
  await runtime.uploadFile('photo.png', Promise.resolve(new Uint8Array([65]).buffer))
  expect(request).not.toHaveBeenCalled()
  expect(runtime.vfs.list('/home/user')).toContain('/home/user/photo.png')
  runtime.dispose()
})

it('gates QA by host support and refreshes saved QA after project restoration', async () => {
  let supported = true
  vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => supported } } })
  try {
    const runtime = createOfficeHostRuntime('powerpoint', {
      presentation: {
        available: () => true,
        request: vi.fn(async () => new Response('{}')),
        documentId: async () => 'doc',
        lastProject: () => undefined,
        rememberProject: async () => {},
        readReceipt: () => undefined,
        readQa: () => undefined,
        writeQa: async () => {},
      },
    })
    expect(runtime.skill.tools.map((t) => t.name)).toContain('capture_presentation_page_qa')
    expect(runtime.skill.tools.map((t) => t.name)).toContain('read_presentation_page')
    expect(runtime.skill.tools.map((t) => t.name)).toContain('edit_presentation_page_text')
    expect(runtime.skill.tools.map((t) => t.name)).toContain('read_presentation_page_geometry')
    expect(runtime.skill.tools.map((t) => t.name)).toContain('edit_presentation_page_geometry')
    supported = false
    expect(runtime.skill.tools.map((t) => t.name)).not.toContain('capture_presentation_page_qa')
    expect(runtime.skill.tools.map((t) => t.name)).not.toContain('edit_presentation_page_text')
    expect(runtime.skill.tools.map((t) => t.name)).not.toContain('edit_presentation_page_geometry')
    supported = true
    const listener = vi.fn()
    runtime.qa!.subscribe(listener)
    await runtime.skill.executeTool({
      id: 'restore',
      name: 'restore_presentation_project',
      input: {},
    })
    expect(listener).toHaveBeenCalledOnce()
    expect(runtime.qa!.revision()).toBe(1)
    runtime.clearSession()
    expect(runtime.qa!.read()).toBeUndefined()
    expect(listener).toHaveBeenCalledTimes(2)
  } finally {
    vi.unstubAllGlobals()
  }
})

it('registers and routes background production controls only through the PC job capability', async () => {
  const request = vi.fn(async () => new Response(JSON.stringify({ error: 'not_found' })))
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      documentId: async () => 'doc',
      lastProject: () => 'p',
      rememberProject: async () => {},
      request,
    },
  })
  try {
    for (const operation of ['start', 'status', 'pause', 'resume', 'cancel']) {
      const name = `${operation === 'status' ? 'read' : operation}_presentation_production_job`
      expect(runtime.skill.tools.map((tool) => tool.name)).toContain(name)
      const result = await runtime.skill.executeTool({
        id: operation,
        name,
        input: { project_id: 'p', request_id: 'r' },
      })
      expect(result.isError).toBe(true)
      expect(request).toHaveBeenLastCalledWith(
        {
          operation: `production_job_${operation}`,
          documentId: 'doc',
          projectId: 'p',
          requestId: 'r',
        },
        undefined,
      )
    }
    expect(runtime.vfs.list('/home/user')).toEqual([])
  } finally {
    runtime.dispose()
  }
})

it('routes evidence delivery and issue actions through their dedicated PC operations', async () => {
  const request = vi.fn(async () => new Response(JSON.stringify({ error: 'not_found' })))
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      documentId: async () => 'doc',
      lastProject: () => 'p',
      rememberProject: async () => {},
      request,
    },
  })
  const action = {
    actionId: 'a1',
    issueId: 'issue-1',
    issueDigest: 'a'.repeat(64),
    state: 'explained',
    note: 'Scoped explanation, not factual acceptance.',
  }
  try {
    for (const [name, operation, extra, body] of [
      ['read_presentation_delivery_report', 'production_delivery_report', {}, {}],
      ['export_presentation_delivery_report', 'production_delivery_report', {}, {}],
      [
        'record_presentation_issue_action',
        'production_record_issue_action',
        { expected_revision: 0, action },
        { expectedRevision: 0, action },
      ],
    ] as const) {
      expect(runtime.skill.tools.map((tool) => tool.name)).toContain(name)
      const result = await runtime.skill.executeTool({
        id: name,
        name,
        input: { project_id: 'p', request_id: 'r', ...extra },
      })
      expect(result.isError).toBe(true)
      expect(request).toHaveBeenLastCalledWith(
        { operation, documentId: 'doc', projectId: 'p', requestId: 'r', ...body },
        undefined,
      )
    }
    expect(runtime.vfs.list('/home/user')).toEqual([])
  } finally {
    runtime.dispose()
  }
})

it('exposes durable text recovery and a changes workbench through the runtime', async () => {
  vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } })
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      documentId: async () => 'doc',
      request: async () => new Response('{}'),
      lastProject: () => undefined,
      rememberProject: async () => {},
      readReceipt: () => undefined,
      readTextChange: () => undefined,
      writeTextChange: async () => {},
      listImageReplacements: () => [],
    },
  })
  try {
    for (const name of [
      'read_presentation_text_change',
      'inspect_presentation_text_change',
      'undo_presentation_text_change',
      'resume_presentation_text_change',
    ]) {
      expect(
        runtime.skill.tools.some((tool) => tool.name === name),
        name,
      ).toBe(true)
      expect(
        (await runtime.skill.executeTool({ id: name, name, input: { page_id: 'p1' } })).output,
      ).toBe('presentation_restore_required')
    }
    expect(runtime.changes).toBeDefined()
    await runtime.changes!.refresh()
    expect(runtime.changes!.snapshot().entries).toEqual([])
    runtime.clearSession()
    expect(runtime.changes!.snapshot().entries).toEqual([])
  } finally {
    runtime.dispose()
    vi.unstubAllGlobals()
  }
})
