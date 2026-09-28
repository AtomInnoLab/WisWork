import { describe, expect, it, vi } from 'vitest'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
import { createPresentationGenerationSkill } from '../src/skills/powerpoint/presentation-generation.js'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document.js'

const deck = {
  version: 1,
  id: 'research-1',
  title: 'Research',
  style: { fontFace: 'Arial', background: 'FFFFFF', textColor: '172B43', accentColor: '226B83' },
  assets: [],
  claims: [],
  slides: [
    {
      id: 'slide-1',
      title: 'Result',
      elements: [{ id: 'title', kind: 'text', text: 'Result', x: 1, y: 1, w: 10, h: 1 }],
    },
  ],
}
const report = {
  deckId: deck.id,
  slideCount: 1,
  elementCount: 1,
  geometry: [],
  checks: {
    structure: 'passed',
    geometry: 'passed',
    render: 'not_run',
    sources: 'not_verified',
    roundTrip: 'not_run',
  },
}
const response = () =>
  new Response(
    JSON.stringify({
      projectId: deck.id,
      requestId: 'request-1',
      status: 'compiled',
      pptxBase64: 'UEsDBAAAAAA=',
      report,
    }),
  )
function fixture() {
  const vfs = new InMemoryVfs()
  const request = vi.fn(async () => response())
  const pdfRequest = vi.fn(async () => response())
  const pdfAvailable = vi.fn(() => true)
  const productionPdfAvailable = vi.fn(() => true)
  const documentId = vi.fn(async () => 'document-1')
  const rememberProject = vi.fn(async () => undefined)
  const available = vi.fn(() => true)
  const skill = createPresentationGenerationSkill({
    vfs,
    request,
    pdfRequest,
    pdfAvailable,
    productionPdfAvailable,
    documentId,
    rememberProject,
    lastProject: () => deck.id,
    available,
  })
  return {
    skill,
    vfs,
    request,
    pdfRequest,
    pdfAvailable,
    productionPdfAvailable,
    documentId,
    rememberProject,
    available,
  }
}
const compileCall = () => ({
  id: 'call-1',
  name: 'compile_deck_with_pptxgenjs',
  input: { request_id: 'request-1', deck },
})

describe('PowerPoint presentation generation', () => {
  it('keeps distinct requests and PDF sources as separate downloadable artifacts', async () => {
    const f = fixture()
    f.request.mockResolvedValueOnce(response()).mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          projectId: deck.id,
          requestId: 'request-2',
          status: 'compiled',
          pptxBase64: btoa('PK\u0003\u0004second'),
          report,
        }),
      ),
    )
    expect((await f.skill.executeTool(compileCall())).isError).not.toBe(true)
    expect(
      (
        await f.skill.executeTool({
          ...compileCall(),
          id: 'compile-2',
          input: { request_id: 'request-2', deck },
        })
      ).isError,
    ).not.toBe(true)
    const firstPptx = '/home/user/generated/research-1/request-1/deck.pptx'
    const secondPptx = '/home/user/generated/research-1/request-2/deck.pptx'
    expect(f.vfs.readBytes(firstPptx)).not.toEqual(f.vfs.readBytes(secondPptx))
    for (const [source, content] of [
      ['compiled', '%PDF-1.4\ncompiled\n%%EOF\n'],
      ['production', '%PDF-1.4\nproduction\n%%EOF\n'],
    ] as const) {
      f.pdfRequest.mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            status: 'exported',
            source,
            projectId: deck.id,
            requestId: 'request-1',
            slideCount: 1,
            pdfBase64: btoa(content),
          }),
        ),
      )
      const result = await f.skill.executeTool({
        id: `pdf-${source}`,
        name: 'export_presentation_pdf',
        input: { project_id: deck.id, request_id: 'request-1', source },
      })
      expect(result.isError).not.toBe(true)
    }
    expect(f.vfs.readText('/home/user/generated/research-1/request-1/compiled.pdf')).toContain(
      'compiled',
    )
    expect(f.vfs.readText('/home/user/generated/research-1/request-1/production.pdf')).toContain(
      'production',
    )
  })

  it('exports a completed project PDF into Session attachments without returning binary to the model', async () => {
    const f = fixture()
    f.pdfRequest.mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          status: 'exported',
          source: 'compiled',
          projectId: deck.id,
          requestId: 'request-1',
          slideCount: 1,
          pdfBase64: btoa('%PDF-1.4\n%%EOF\n'),
        }),
      ),
    )
    const result = await f.skill.executeTool({
      id: 'pdf-call',
      name: 'export_presentation_pdf',
      input: { project_id: deck.id, request_id: 'request-1' },
    })
    expect(result.isError).not.toBe(true)
    expect(f.pdfRequest).toHaveBeenCalledWith(
      {
        operation: 'export_pdf',
        documentId: 'document-1',
        projectId: deck.id,
        requestId: 'request-1',
      },
      undefined,
    )
    expect(f.vfs.list('/home/user')).toContain(
      '/home/user/generated/research-1/request-1/compiled.pdf',
    )
    expect(result.output).not.toContain('JVBER')
  })

  it('exports the exact completed page-production request as one PDF', async () => {
    const f = fixture()
    f.pdfRequest.mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          status: 'exported',
          source: 'production',
          projectId: deck.id,
          requestId: 'request-1',
          slideCount: 8,
          pdfBase64: btoa('%PDF-1.4\n%%EOF\n'),
        }),
      ),
    )
    const result = await f.skill.executeTool({
      id: 'production-pdf-call',
      name: 'export_presentation_pdf',
      input: { project_id: deck.id, request_id: 'request-1', source: 'production' },
    })
    expect(result.isError).not.toBe(true)
    expect(f.pdfRequest).toHaveBeenCalledWith(
      {
        operation: 'export_pdf',
        documentId: 'document-1',
        projectId: deck.id,
        requestId: 'request-1',
        source: 'production',
      },
      undefined,
    )
    expect(result.output).toContain('"source":"production"')
    expect(f.vfs.list('/home/user')).toContain(
      '/home/user/generated/research-1/request-1/production.pdf',
    )
  })

  it('accepts the legacy compiled PDF receipt without a source field', async () => {
    const f = fixture()
    f.pdfRequest.mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          status: 'exported',
          projectId: deck.id,
          requestId: 'request-1',
          slideCount: 1,
          pdfBase64: btoa('%PDF-1.4\n%%EOF\n'),
        }),
      ),
    )
    const result = await f.skill.executeTool({
      id: 'legacy-pdf-call',
      name: 'export_presentation_pdf',
      input: { project_id: deck.id, request_id: 'request-1' },
    })
    expect(result.isError).not.toBe(true)
    expect(result.output).toContain('"source":"compiled"')
  })

  it('rejects production PDF when only the legacy compiled PDF capability is available', async () => {
    const f = fixture()
    f.productionPdfAvailable.mockReturnValue(false)
    const result = await f.skill.executeTool({
      id: 'production-pdf-call',
      name: 'export_presentation_pdf',
      input: { project_id: deck.id, request_id: 'request-1', source: 'production' },
    })
    expect(result).toMatchObject({ isError: true, output: 'presentation_pdf_unavailable' })
    expect(f.pdfRequest).not.toHaveBeenCalled()
  })

  it('exports production PDF when only the production PDF capability is available', async () => {
    const f = fixture()
    f.pdfAvailable.mockReturnValue(false)
    f.pdfRequest.mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          status: 'exported',
          source: 'production',
          projectId: deck.id,
          requestId: 'request-1',
          slideCount: 1,
          pdfBase64: btoa('%PDF-1.4\n%%EOF\n'),
        }),
      ),
    )
    expect(f.skill.tools.map((tool) => tool.name)).toContain('export_presentation_pdf')
    const result = await f.skill.executeTool({
      id: 'production-only-pdf-call',
      name: 'export_presentation_pdf',
      input: { project_id: deck.id, request_id: 'request-1', source: 'production' },
    })
    expect(result.isError).not.toBe(true)
  })

  it('hides PDF export on an older PC without its negotiated capability', async () => {
    const f = fixture()
    f.pdfAvailable.mockReturnValue(false)
    f.productionPdfAvailable.mockReturnValue(false)
    expect(f.skill.tools.map((tool) => tool.name)).not.toContain('export_presentation_pdf')
    expect(
      await f.skill.executeTool({
        id: 'pdf-call',
        name: 'export_presentation_pdf',
        input: { project_id: deck.id, request_id: 'request-1' },
      }),
    ).toMatchObject({ isError: true, output: 'presentation_pdf_unavailable' })
    expect(f.pdfRequest).not.toHaveBeenCalled()
  })

  it('compiles through the paired PC and delivers PPTX plus report without sending binary back to the model', async () => {
    const f = fixture()
    const result = await f.skill.executeTool(compileCall())
    expect(result.isError).not.toBe(true)
    expect(f.request).toHaveBeenCalledWith(
      expect.objectContaining({
        operation: 'compile',
        documentId: 'document-1',
        requestId: 'request-1',
        deck,
      }),
      undefined,
    )
    expect(f.vfs.list('/home/user')).toEqual([
      '/home/user/generated/research-1/request-1/deck.pptx',
      '/home/user/generated/research-1/request-1/report.json',
    ])
    expect(result.output).toContain('not_run')
    expect(result.output).not.toContain('UEsD')
    expect(f.rememberProject).toHaveBeenCalledWith(deck.id)
    expect(result.mutated).toBe(false)
  })
  it('rejects invalid SlideIR before using the PC', async () => {
    const f = fixture()
    const result = await f.skill.executeTool({
      ...compileCall(),
      input: { request_id: 'request-1', deck: { ...deck, slides: [] } },
    })
    expect(result.isError).toBe(true)
    expect(f.request).not.toHaveBeenCalled()
  })
  it('does not expose compilation tools when the capability was not negotiated', () => {
    const f = fixture()
    f.available.mockReturnValue(false)
    expect(f.skill.tools).toEqual([])
  })
  it('fails closed if the capability disappears before execution', async () => {
    const f = fixture()
    f.available.mockReturnValue(false)
    expect(await f.skill.executeTool(compileCall())).toMatchObject({
      isError: true,
      output: 'presentation_unavailable',
    })
    expect(f.request).not.toHaveBeenCalled()
  })
  it('rejects an artifact belonging to another project', async () => {
    const f = fixture()
    f.request.mockResolvedValue(
      new Response(
        JSON.stringify({
          projectId: 'other',
          requestId: 'request-1',
          status: 'compiled',
          pptxBase64: 'UEsDBAAAAAA=',
          report,
        }),
      ),
    )
    expect(await f.skill.executeTool(compileCall())).toMatchObject({ isError: true })
    expect(f.vfs.list('/home/user')).toEqual([])
  })
  it('does not publish a late artifact after the active document changed', async () => {
    const f = fixture()
    f.documentId.mockResolvedValueOnce('document-1').mockResolvedValue('document-2')
    expect(await f.skill.executeTool(compileCall())).toMatchObject({
      isError: true,
      output: 'presentation_document_changed',
    })
    expect(f.vfs.list('/home/user')).toEqual([])
  })
  it('restores the last compiled project after recreating the Taskpane runtime', async () => {
    const f = fixture()
    expect(
      await f.skill.executeTool({ id: 'restore', name: 'restore_presentation_project', input: {} }),
    ).toMatchObject({ mutated: false })
    expect(f.request).toHaveBeenCalledWith(
      { operation: 'get', documentId: 'document-1', projectId: deck.id },
      undefined,
    )
    expect(f.vfs.list('/home/user')).toContain(
      '/home/user/generated/research-1/request-1/deck.pptx',
    )
  })
  it('preserves existing downloads on cancellation', async () => {
    const f = fixture()
    const controller = new AbortController()
    f.request.mockImplementation(async () => {
      controller.abort()
      return response()
    })
    f.vfs.writeFile('/home/user/existing.pptx', new Uint8Array([1]))
    expect(await f.skill.executeTool(compileCall(), controller.signal)).toMatchObject({
      isError: true,
      output: 'cancelled',
    })
    expect(f.vfs.list('/home/user')).toEqual(['/home/user/existing.pptx'])
  })
})

describe('presentation document binding', () => {
  function setup() {
    const values = new Map<string, unknown>()
    const save = vi.fn(async () => {})
    let location = 'file:///research.pptx'
    const runtime = {
      get: (key: string) => values.get(key),
      set: (key: string, value: string) => {
        values.set(key, value)
      },
      save,
      location: () => location,
    }
    return {
      values,
      save,
      runtime,
      move: (url: string) => {
        location = url
      },
    }
  }
  it('persists a host document ID across runtimes and records a recoverable project', async () => {
    const f = setup()
    const a = createPresentationDocumentBinding(f.runtime, () => 'document-uuid')
    const id = await a.documentId()
    await a.rememberProject('project-1')
    const b = createPresentationDocumentBinding(f.runtime, () => 'should-not-be-used')
    expect(await b.documentId()).toBe(id)
    expect(b.lastProject()).toBe('project-1')
    f.move('file:///copy.pptx')
    expect(await b.documentId()).not.toBe(id)
  })
  it('persists a selected page task only for the same document and project', async () => {
    const f = setup()
    const a = createPresentationDocumentBinding(f.runtime, () => 'document-uuid')
    const documentId = await a.documentId()
    await a.rememberSelectedProduction('project-1', documentId, 'older-request')
    const b = createPresentationDocumentBinding(f.runtime)
    expect(b.selectedProduction('project-1', documentId)).toBe('older-request')
    expect(b.selectedProduction('project-2', documentId)).toBeUndefined()
    f.move('file:///copy.pptx')
    expect(b.selectedProduction('project-1', await b.documentId())).toBeUndefined()
  })
  it('does not expose an unsaved selected task after settings save fails', async () => {
    const f = setup()
    const binding = createPresentationDocumentBinding(f.runtime, () => 'document-uuid')
    const documentId = await binding.documentId()
    f.save.mockRejectedValueOnce(new Error('save_failed'))
    await expect(
      binding.rememberSelectedProduction('project-1', documentId, 'older-request'),
    ).rejects.toThrow('save_failed')
    expect(binding.selectedProduction('project-1', documentId)).toBeUndefined()
    await expect(
      binding.rememberSelectedProduction('project-1', 'wrong-document', 'older-request'),
    ).rejects.toThrow('presentation_document_changed')
  })
  it('rejects a selection write if Save As changes the document during settings save', async () => {
    const f = setup()
    const binding = createPresentationDocumentBinding(f.runtime, () => 'document-uuid')
    const documentId = await binding.documentId()
    f.save.mockImplementationOnce(async () => {
      f.move('file:///copy.pptx')
    })
    await expect(
      binding.rememberSelectedProduction('project-1', documentId, 'older-request'),
    ).rejects.toThrow('presentation_document_changed')
    expect(binding.selectedProduction('project-1', await binding.documentId())).toBeUndefined()
  })
  it('serializes simultaneous identity creation', async () => {
    const f = setup()
    const uuid = vi.fn(() => 'document-uuid')
    const binding = createPresentationDocumentBinding(f.runtime, uuid)
    expect(new Set(await Promise.all([binding.documentId(), binding.documentId()])).size).toBe(1)
    expect(uuid).toHaveBeenCalledOnce()
    expect(f.save).toHaveBeenCalledOnce()
  })
  it('fails closed after settings persistence fails instead of using an unsaved identity', async () => {
    const f = setup()
    f.save.mockRejectedValue(new Error('cannot save'))
    const binding = createPresentationDocumentBinding(f.runtime, () => 'document-uuid')
    await expect(binding.documentId()).rejects.toThrow('presentation_document_identity_unavailable')
    await expect(binding.documentId()).rejects.toThrow('presentation_document_identity_unavailable')
  })
})

describe('generation lifecycle and source preservation', () => {
  it('keeps a user attachment with the same basename as generated output', async () => {
    const f = fixture()
    const source = new Uint8Array([7, 8, 9])
    f.vfs.writeFile('/home/user/research-1.pptx', source)
    await f.skill.executeTool(compileCall())
    expect(f.vfs.readBytes('/home/user/research-1.pptx')).toEqual(source)
    expect(f.vfs.list('/home/user')).toContain(
      '/home/user/generated/research-1/request-1/deck.pptx',
    )
  })
  it('does not repopulate artifacts after logout or a new task clears a pending compile', async () => {
    const f = fixture()
    let resolve!: (value: Response) => void
    f.request.mockImplementation(
      () =>
        new Promise<Response>((done) => {
          resolve = done
        }),
    )
    const pending = f.skill.executeTool(compileCall())
    await vi.waitFor(() => expect(f.request).toHaveBeenCalled())
    f.skill.clear()
    f.vfs.clear()
    resolve(response())
    expect(await pending).toMatchObject({ isError: true, output: 'cancelled' })
    expect(f.vfs.list('/home/user')).toEqual([])
    expect(f.skill.artifact()).toBeUndefined()
  })
})

describe('generation final document-check cancellation', () => {
  it.each([2, 3])(
    'does not publish after clear while identity check %s is pending',
    async (boundary) => {
      const f = fixture()
      let calls = 0
      let resolve!: (id: string) => void
      f.documentId.mockImplementation(async () => {
        calls += 1
        return calls === boundary
          ? new Promise<string>((done) => {
              resolve = done
            })
          : 'document-1'
      })
      const pending = f.skill.executeTool(compileCall())
      await vi.waitFor(() => expect(calls).toBe(boundary))
      f.skill.clear()
      resolve('document-1')
      expect(await pending).toMatchObject({ isError: true, output: 'cancelled' })
      expect(f.vfs.list('/home/user')).toEqual([])
      if (boundary === 2) expect(f.request).not.toHaveBeenCalled()
    },
  )
})

describe('durable project recovery entry', () => {
  it('remembers the project before the first PC response can be lost', async () => {
    const f = fixture()
    f.request.mockImplementation(async () => {
      expect(f.rememberProject).toHaveBeenCalledWith(deck.id)
      throw new Error('connection_lost')
    })
    expect(await f.skill.executeTool(compileCall())).toMatchObject({ isError: true })
    expect(f.rememberProject).toHaveBeenCalledWith(deck.id)
  })
  it('resumes a saved request without sending a replacement deck', async () => {
    const f = fixture()
    expect(
      await f.skill.executeTool({
        id: 'resume',
        name: 'resume_presentation_project',
        input: { project_id: deck.id, request_id: 'request-1' },
      }),
    ).toMatchObject({ mutated: false })
    expect(f.request).toHaveBeenCalledWith(
      { operation: 'resume', documentId: 'document-1', projectId: deck.id, requestId: 'request-1' },
      undefined,
    )
    expect(f.vfs.list('/home/user')).toContain(
      '/home/user/generated/research-1/request-1/deck.pptx',
    )
  })
})

describe('failed recovery preserves the selected project', () => {
  it.each(['restore_presentation_project', 'resume_presentation_project'])(
    '%s does not replace the saved project with a missing/foreign one',
    async (name) => {
      const f = fixture()
      f.request.mockResolvedValue(new Response(JSON.stringify({ error: 'not_found' })))
      const result = await f.skill.executeTool({
        id: 'missing',
        name,
        input: {
          project_id: 'missing-project',
          ...(name === 'resume_presentation_project' ? { request_id: 'request-1' } : {}),
        },
      })
      expect(result.isError).toBe(true)
      expect(f.rememberProject).not.toHaveBeenCalled()
    },
  )
})

it('forwards a saved plan revision with a compilation request', async () => {
  const f = fixture()
  const result = await f.skill.executeTool({
    ...compileCall(),
    input: { ...compileCall().input, plan_revision: 2 },
  })
  expect(result.isError).not.toBe(true)
  expect(f.request).toHaveBeenCalledWith(
    expect.objectContaining({ operation: 'compile', planRevision: 2 }),
    undefined,
  )
})
it('requires negotiated asset references while allowing compact reference compilation on new PCs', async () => {
  const referenced = { ...deck, assets: [{ id: 'photo', attachmentId: 'a'.repeat(64) }] }
  const old = fixture()
  expect(
    await old.skill.executeTool({
      id: 'ref',
      name: 'compile_deck_with_pptxgenjs',
      input: { request_id: 'request-1', deck: referenced },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_assets_unavailable' })
  expect(old.request).not.toHaveBeenCalled()
  const request = vi.fn(async () => response()),
    assetsAvailable = vi.fn(() => true)
  const skill = createPresentationGenerationSkill({
    vfs: new InMemoryVfs(),
    available: () => true,
    assetsAvailable,
    request,
    documentId: async () => 'doc',
    lastProject: () => undefined,
    rememberProject: async () => {},
  })
  expect(
    await skill.executeTool({
      id: 'ref',
      name: 'compile_deck_with_pptxgenjs',
      input: { request_id: 'request-1', deck: referenced },
    }),
  ).not.toHaveProperty('isError', true)
  expect(request).toHaveBeenCalledWith(expect.objectContaining({ deck: referenced }), undefined)
})

it('retains immutable source-page metadata and rejects mismatched page identities', async () => {
  const f = fixture()
  const value = await response().json()
  const pages = [{ id: 'slide-1', title: 'Result', sourceSlideId: '256#' }]
  f.request.mockResolvedValueOnce(new Response(JSON.stringify({ ...value, pages })))
  expect((await f.skill.executeTool(compileCall())).isError).not.toBe(true)
  expect(f.skill.artifact()?.pages).toEqual(pages)
  expect(Object.isFrozen(f.skill.artifact()?.pages)).toBe(true)
  for (const page of [
    { ...pages[0], id: 'other' },
    { ...pages[0], sourceSlideId: '256' },
    { ...pages[0], sourceSlideId: '4294967296#' },
  ]) {
    f.request.mockResolvedValueOnce(new Response(JSON.stringify({ ...value, pages: [page] })))
    expect((await f.skill.executeTool(compileCall())).output).toBe('presentation_response_invalid')
  }
})
