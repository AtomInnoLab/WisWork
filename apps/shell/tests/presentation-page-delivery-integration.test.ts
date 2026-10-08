import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationService } from '../src/main/presentation-service'
import { createPresentationGenerationSkill } from '../../office-addin/src/skills/powerpoint/presentation-generation'
import { createPresentationDeliverySkill } from '../../office-addin/src/skills/powerpoint/presentation-delivery'
import { createPresentationDocumentBinding } from '../../office-addin/src/skills/powerpoint/presentation-document'
import { InMemoryVfs } from '../../office-addin/src/skills/shared/vfs'
import { createStructuredProposalController } from '../../office-addin/src/agent/proposal-controller'

it('restores real compiled page metadata and continues only the remaining pages after a safe failure', async () => {
  const root = mkdtempSync(join(tmpdir(), 'presentation-page-recovery-'))
  try {
    let service = createPresentationService({ userDataPath: root })
    let values = new Map<string, string>()
    let persisted = new Map<string, string>()
    const binding = () =>
      createPresentationDocumentBinding(
        {
          get: (key) => values.get(key),
          set: (key, value) => {
            values.set(key, value)
          },
          save: async () => {
            persisted = new Map(values)
          },
          location: () => 'file://recovery.pptx',
        },
        () => 'doc-recovery',
      )
    const setup = () => {
      const settings = binding()
      const generation = createPresentationGenerationSkill({
        ...settings,
        vfs: new InMemoryVfs(),
        available: () => true,
        request: async (body, signal) =>
          new Response(Buffer.from(await service(body, signal ?? new AbortController().signal))),
      })
      const proposals = createStructuredProposalController()
      const delivery = createPresentationDeliverySkill({
        ...settings,
        proposals,
        adapter,
        available: () => true,
        artifact: generation.artifact,
      })
      return { settings, generation, proposals, delivery }
    }
    const ids = ['existing-page']
    const inserted: string[] = []
    let failThird = true
    const snapshot = async () => ({ slideIds: [...ids], fingerprint: JSON.stringify(ids) })
    const adapter = {
      available: () => true,
      snapshot,
      insert: async () => {
        throw new Error('unexpected_bulk_import')
      },
      insertPage: async (_base64: string, sourceSlideId: string) => {
        if (inserted.length === 2 && failThird) {
          failThird = false
          throw new Error('cancelled')
        }
        inserted.push(sourceSlideId)
        const id = `inserted-${inserted.length}`
        ids.push(id)
        return { slideIds: [id] }
      },
      verify: async (receipt: { slideIds: string[] }, before: { slideIds: string[] }) =>
        JSON.stringify([...before.slideIds, ...receipt.slideIds]) === JSON.stringify(ids),
    }
    let runtime = setup()
    const deck = benchmarkDeck()
    expect(
      (
        await runtime.generation.executeTool({
          id: 'compile',
          name: 'compile_deck_with_pptxgenjs',
          input: { request_id: 'request-pages', deck },
        })
      ).isError,
    ).not.toBe(true)
    expect(runtime.generation.artifact()?.pages).toHaveLength(8)
    const call = {
      id: 'import',
      name: 'import_generated_presentation',
      input: { project_id: deck.id },
    }
    await runtime.delivery.executeTool(call)
    await expect(runtime.proposals.confirm(runtime.proposals.pending()!.id)).rejects.toThrow(
      'cancelled',
    )
    expect(inserted).toEqual(['256#', '257#'])
    expect(
      runtime.settings.readReceipt(`${deck.id}/request-pages`)?.checkpoint?.inFlight,
    ).toBeUndefined()
    values = new Map(persisted)
    service = createPresentationService({ userDataPath: root })
    runtime = setup()
    expect(
      (
        await runtime.generation.executeTool({
          id: 'restore',
          name: 'restore_presentation_project',
          input: { project_id: deck.id },
        })
      ).isError,
    ).not.toBe(true)
    const status = await runtime.delivery.executeTool({
      id: 'status',
      name: 'read_presentation_import_status',
      input: { project_id: deck.id },
    })
    expect(JSON.parse(status.output)).toMatchObject({ completed: 2, total: 8, status: 'partial' })
    await runtime.delivery.executeTool(call)
    await runtime.proposals.confirm(runtime.proposals.pending()!.id)
    expect(inserted).toEqual(['256#', '257#', '258#', '259#', '260#', '261#', '262#', '263#'])
    expect(ids).toHaveLength(9)
    expect(runtime.settings.readReceipt(`${deck.id}/request-pages`)?.state).toBe('complete')
    expect((await runtime.delivery.executeTool(call)).output).toContain('already_imported')
    expect(inserted).toHaveLength(8)
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
