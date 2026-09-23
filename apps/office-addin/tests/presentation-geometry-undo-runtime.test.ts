import { afterEach, expect, it, vi } from 'vitest'
import { createHash } from 'node:crypto'
import { createOfficeHostRuntime } from '../src/agent/host-runtime'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
import { presentationArtifactContent } from '../src/skills/powerpoint/presentation-page-delivery'

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})
it.each([false, true])(
  'persists and recovers geometry without duplicate host writes (completion save fails: %s)',
  async (failSaves) => {
    vi.stubGlobal('Office', {
      context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } },
    })
    let failState: 'applied' | 'undone' | undefined
    const settings = new Map<string, unknown>()
    const documentSettings = {
      get: (key: string) => settings.get(key),
      set: (key: string, value: string) => {
        settings.set(key, value)
      },
      save: async () => {
        if (
          failState &&
          [...settings.values()].some(
            (raw) =>
              typeof raw === 'string' &&
              raw.includes('"before":') &&
              raw.includes(`"state":"${failState}"`),
          )
        ) {
          failState = undefined
          throw new Error('save_failed')
        }
      },
      location: () => 'file',
    }
    const binding = createPresentationDocumentBinding(documentSettings, () => 'doc')
    const documentId = await binding.documentId()
    const pageBytes = 'UEsDBAAAAAA='
    const pages = [
      { id: 'first', title: 'First', sourceSlideId: '256#' },
      { id: 'second', title: 'Second', sourceSlideId: '256#' },
    ]
    const artifact = {
      documentId,
      projectId: 'project',
      requestId: 'request',
      pptxBase64: '',
      pagePptxBase64: [pageBytes, pageBytes],
      planRevision: 1,
      slideCount: 2,
      pages,
    }
    await binding.writeReceipt('production/project/request', {
      state: 'complete',
      documentId,
      slideIds: ['host1', 'host2'],
      checkpoint: {
        version: 2,
        artifactDigest: createHash('sha256')
          .update(presentationArtifactContent(artifact))
          .digest('hex'),
        sourceSlideIds: ['256#', '256#'],
        pageIds: ['first', 'second'],
        baselineSlideIds: [],
        completed: [
          { sourceSlideId: '256#', slideId: 'host1' },
          { sourceSlideId: '256#', slideId: 'host2' },
        ],
      },
    })
    const before = { left: 1, top: 2, width: 30, height: 40 },
      after = { left: 12, top: 13, width: 60, height: 70 }
    let geometry = { ...before }
    vi.spyOn(BrowserPowerPointAdapter.prototype, 'readPresentationPageGeometry').mockImplementation(
      async (slideId, shapeId) => ({ slideId, shapeId, geometry: { ...geometry } }),
    )
    const writes = vi
      .spyOn(BrowserPowerPointAdapter.prototype, 'editPresentationPageGeometry')
      .mockImplementation(async (slideId, shapeId, value, expected) => {
        expect([slideId, shapeId]).toEqual(['host2', 'shape'])
        expect(expected).toEqual(geometry)
        geometry = { ...value }
      })
    const invalidations = vi.fn(async (_ids?: readonly string[]) => {})
    const create = () =>
      createOfficeHostRuntime('powerpoint', {
        presentation: {
          ...createPresentationDocumentBinding(documentSettings, () => 'doc'),
          available: () => true,
          invalidateQa: invalidations,
          request: async (body) => {
            const operation = (body as { operation: string }).operation
            return new Response(
              JSON.stringify(
                operation === 'production_status'
                  ? {
                      projectId: 'project',
                      requestId: 'request',
                      planRevision: 1,
                      status: 'compiled',
                      compiledCount: 2,
                      total: 2,
                      pages: pages.map((p) => ({
                        id: p.id,
                        title: p.title,
                        state: 'compiled',
                        attempt: 1,
                      })),
                    }
                  : {
                      projectId: 'project',
                      requestId: 'request',
                      pageId: (body as { pageId: string }).pageId,
                      planRevision: 1,
                      status: 'compiled',
                      pptxBase64: pageBytes,
                      sourceSlideId: '256#',
                      report: { deckId: 'project', slideCount: 1 },
                    },
              ),
            )
          },
        },
      })
    let runtime = create()
    const prepare = {
      id: 'prepare',
      name: 'prepare_presentation_production_import',
      input: { project_id: 'project', request_id: 'request' },
    }
    try {
      expect((await runtime.skill.executeTool(prepare)).isError).not.toBe(true)
      const edit = await runtime.skill.executeTool({
        id: 'edit',
        name: 'edit_presentation_page_geometry',
        input: { page_id: 'second', shape_id: 'shape', geometry: after },
      })
      expect(edit.isError, edit.output).not.toBe(true)
      if (failSaves) {
        failState = 'applied'
        await expect(runtime.proposals.confirm(runtime.proposals.pending()!.id)).rejects.toThrow(
          'save_failed',
        )
        expect(binding.readGeometryChange()?.state).toBe('pending')
      } else await runtime.proposals.confirm(runtime.proposals.pending()!.id)
      expect(geometry).toEqual(after)
      runtime.dispose()
      runtime = create()
      expect((await runtime.skill.executeTool(prepare)).isError).not.toBe(true)
      const recover = async () => {
        const inspection = await runtime.skill.executeTool({
          id: 'inspect',
          name: 'inspect_presentation_geometry_change',
          input: { page_id: 'second' },
        })
        expect(inspection.isError, inspection.output).not.toBe(true)
        expect(JSON.parse(inspection.output)).toMatchObject({ status: 'already_applied' })
        const result = await runtime.skill.executeTool({
          id: 'recover',
          name: 'resume_presentation_geometry_change',
          input: { page_id: 'second' },
        })
        expect(result.isError, result.output).not.toBe(true)
        await runtime.proposals.confirm(runtime.proposals.pending()!.id)
      }
      if (failSaves) {
        await recover()
        expect(writes).toHaveBeenCalledTimes(1)
      }
      const read = await runtime.skill.executeTool({
        id: 'read',
        name: 'read_presentation_geometry_change',
        input: { page_id: 'second' },
      })
      expect(read.isError, read.output).not.toBe(true)
      expect(JSON.parse(read.output)).toMatchObject({
        historical: true,
        record: { state: 'applied', before, after },
      })
      const undoCall = {
        id: 'undo',
        name: 'undo_presentation_geometry_change',
        input: { page_id: 'second' },
      }
      const undo = await runtime.skill.executeTool(undoCall)
      expect(undo.isError, undo.output).not.toBe(true)
      if (failSaves) {
        failState = 'undone'
        await expect(runtime.proposals.confirm(runtime.proposals.pending()!.id)).rejects.toThrow(
          'save_failed',
        )
        expect(binding.readGeometryChange()?.state).toBe('undo_pending')
        runtime.dispose()
        runtime = create()
        expect((await runtime.skill.executeTool(prepare)).isError).not.toBe(true)
        await recover()
        expect(binding.readGeometryChange()?.state).toBe('undone')
      } else await runtime.proposals.confirm(runtime.proposals.pending()!.id)
      expect(geometry).toEqual(before)
      expect(writes).toHaveBeenCalledTimes(2)
      expect(invalidations.mock.calls).toEqual(
        Array.from({ length: failSaves ? 4 : 2 }, () => [['host2']]),
      )
      const repeated = await runtime.skill.executeTool(undoCall)
      expect(repeated.isError, repeated.output).not.toBe(true)
      expect(runtime.proposals.pending()).toBeUndefined()
      expect(writes).toHaveBeenCalledTimes(2)
    } finally {
      runtime.dispose()
    }
  },
)
