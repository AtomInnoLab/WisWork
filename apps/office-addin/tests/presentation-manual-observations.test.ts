import { expect, it, vi } from 'vitest'
import type { PresentationManualObservationShape } from '@wiswork/pptx-engine/presentation-manual-observation'
import { createPresentationManualObservationSkill } from '../src/skills/powerpoint/presentation-manual-observations'
import { createStructuredProposalController } from '../src/agent/proposal-controller'

const shape: PresentationManualObservationShape = {
  id: 'shape',
  name: 'Title',
  type: 'TextBox',
  left: 0,
  top: 0,
  width: 100,
  height: 50,
  rotation: 0,
  text: 'Before',
  font: { name: 'Arial', size: 20, color: '000000' },
}
function fixture() {
  const proposals = createStructuredProposalController()
  const adapter = {
    readContext: vi.fn(async () => ({
      slideIds: ['slide'],
      selectedSlideIds: ['slide'],
      selectedShapeIds: ['shape'],
    })),
    readPage: vi.fn(async () => ({ slideId: 'slide', shapes: [shape] })),
  }
  const request = vi.fn(async (_body: unknown, _signal?: AbortSignal) => new Response('{}'))
  let doc = 'doc'
  const skill = createPresentationManualObservationSkill({
    adapter,
    proposals,
    available: () => true,
    documentId: async () => doc,
    request,
  })
  const call = (name: string, input: Record<string, unknown>) =>
    skill.executeTool({ name, input, id: 'call' })
  return {
    skill,
    adapter,
    request,
    proposals,
    call,
    setDoc: (value: string) => {
      doc = value
    },
  }
}
it('reads a stable exact target twice before requesting durable before capture', async () => {
  const f = fixture()
  await f.call('begin_presentation_edit_observation', {
    project_id: 'p',
    slide_id: 'slide',
    shape_id: 'shape',
  })
  expect(f.adapter.readPage).toHaveBeenCalledTimes(2)
  expect(f.request).toHaveBeenCalledWith(
    expect.objectContaining({ documentId: 'doc', projectId: 'p', slideId: 'slide', shape }),
    undefined,
  )
})
it('captures text placeholders but refuses picture placeholders with unreadable content', async () => {
  const f = fixture()
  const placeholder = { ...shape, type: 'Placeholder' as const }
  f.adapter.readPage.mockResolvedValue({ slideId: 'slide', shapes: [placeholder] })
  await f.call('begin_presentation_edit_observation', {
    project_id: 'p',
    slide_id: 'slide',
    shape_id: 'shape',
  })
  expect(f.request).toHaveBeenCalledWith(
    expect.objectContaining({ shape: expect.objectContaining({ type: 'Placeholder' }) }),
    undefined,
  )
  f.request.mockClear()
  f.adapter.readPage.mockResolvedValue({
    slideId: 'slide',
    shapes: [{ ...placeholder, text: undefined }],
  })
  const picture = await f.call('begin_presentation_edit_observation', {
    project_id: 'p',
    slide_id: 'slide',
    shape_id: 'shape',
  })
  expect(picture.isError).toBe(true)
  expect(f.request).not.toHaveBeenCalled()
})
it('rejects unstable target without persisting an observation', async () => {
  const f = fixture()
  f.adapter.readPage
    .mockResolvedValueOnce({ slideId: 'slide', shapes: [shape] })
    .mockResolvedValueOnce({ slideId: 'slide', shapes: [{ ...shape, text: 'Changed' }] })
  const result = await f.call('begin_presentation_edit_observation', {
    project_id: 'p',
    slide_id: 'slide',
    shape_id: 'shape',
  })
  expect(result.isError).toBe(true)
  expect(f.request).not.toHaveBeenCalled()
})
it('clear during an awaited read prevents durable capture', async () => {
  const f = fixture()
  f.adapter.readPage.mockImplementationOnce(async () => {
    f.skill.clear()
    return { slideId: 'slide', shapes: [shape] }
  })
  await f.call('begin_presentation_edit_observation', {
    project_id: 'p',
    slide_id: 'slide',
    shape_id: 'shape',
  })
  expect(f.request).not.toHaveBeenCalled()
})

import {
  presentationManualObservationDigest,
  type PresentationManualObservation,
} from '@wiswork/pptx-engine/presentation-manual-observation'
function durableFixture() {
  const f = fixture()
  let record: PresentationManualObservation | undefined
  let current = { ...shape }
  f.adapter.readPage.mockImplementation(async () => ({ slideId: 'slide', shapes: [current] }))
  f.request.mockImplementation(async (body?: unknown) => {
    const b = body as Record<string, unknown>
    if (b.operation === 'manual_observation_begin')
      record = {
        version: 1,
        source: 'host_difference_unattributed',
        observationId: b.observationId as string,
        documentId: 'doc',
        projectId: 'p',
        slideId: 'slide',
        shapeId: 'shape',
        before: {
          capturedAt: '2026-09-29T00:00:00.000Z',
          shape: b.shape as typeof shape,
          digest: await presentationManualObservationDigest(b.shape as typeof shape),
        },
        atomicSnapshot: false,
        coverage: 'text_geometry_aggregate_font',
      }
    if (b.operation === 'manual_observation_complete')
      record = {
        ...record!,
        after: {
          capturedAt: '2026-09-29T00:01:00.000Z',
          shape: b.shape as typeof shape,
          digest: await presentationManualObservationDigest(b.shape as typeof shape),
        },
      }
    if (b.operation === 'preference_save_observation')
      return new Response(
        JSON.stringify({
          preference: {
            projectId: 'p',
            changeId: `manual_${record!.observationId}`,
            text: b.text,
            origin: {
              version: 1,
              observationId: record!.observationId,
              beforeDigest: record!.before.digest,
              afterDigest: record!.after!.digest,
            },
          },
        }),
      )
    if (b.operation === 'manual_observation_delete')
      return new Response(JSON.stringify({ deleted: true }))
    return new Response(JSON.stringify({ observation: record }))
  })
  const begin = async () => {
    const result = await f.call('begin_presentation_edit_observation', {
      project_id: 'p',
      slide_id: 'slide',
      shape_id: 'shape',
    })
    expect(result.isError, result.output).not.toBe(true)
    return JSON.parse(result.output).observation.observationId as string
  }
  const complete = async (id: string) => {
    const result = await f.call('complete_presentation_edit_observation', {
      project_id: 'p',
      observation_id: id,
    })
    expect(result.isError, result.output).not.toBe(true)
  }
  return {
    ...f,
    begin,
    complete,
    edit: (text: string) => {
      current = { ...current, text }
    },
    rotate: (rotation: number) => {
      current = { ...current, rotation }
    },
    get: () => record!,
  }
}
it('captures a rotation-only host edit in the durable observation', async () => {
  const f = durableFixture()
  const id = await f.begin()
  f.rotate(45)
  await f.complete(id)
  expect(f.get().before.shape.rotation).toBe(0)
  expect(f.get().after?.shape.rotation).toBe(45)
  expect(f.get().after?.digest).not.toBe(f.get().before.digest)
})
it('persists before, completes after clear, then only saves a preference after confirmation', async () => {
  const f = durableFixture(),
    id = await f.begin()
  f.skill.clear()
  f.edit('After')
  await f.complete(id)
  const result = await f.call('save_presentation_observed_preference', {
    project_id: 'p',
    observation_id: id,
    preference: 'Keep this style',
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(
    f.request.mock.calls.some(
      ([body]) => (body as Record<string, unknown>).operation === 'preference_save_observation',
    ),
  ).toBe(false)
  const proposal = f.proposals.pending()!
  expect(proposal.fingerprint).toMatch(/^[a-f0-9]{64}$/)
  await f.proposals.confirm(proposal.id)
  expect(f.request).toHaveBeenCalledWith(
    expect.objectContaining({
      operation: 'preference_save_observation',
      expectedBeforeDigest: f.get().before.digest,
      expectedAfterDigest: f.get().after!.digest,
    }),
    expect.any(AbortSignal),
  )
})
it('does not accept an old before snapshot when completion never reached PC', async () => {
  const f = durableFixture()
  const id = await f.begin()
  f.edit('After')
  const original = f.request.getMockImplementation()!
  f.request.mockImplementation(async (body, signal) => {
    if ((body as { operation: string }).operation === 'manual_observation_complete')
      throw new Error('request failed before commit')
    return original(body, signal)
  })
  const result = await f.call('complete_presentation_edit_observation', {
    project_id: 'p',
    observation_id: id,
  })
  expect(result.isError).toBe(true)
  expect(f.get().after).toBeUndefined()
})
it('continued host edits after proposing block approval and clear cancels its pending decision', async () => {
  const f = durableFixture(),
    id = await f.begin()
  f.edit('After')
  await f.complete(id)
  await f.call('save_presentation_observed_preference', {
    project_id: 'p',
    observation_id: id,
    preference: 'Keep',
  })
  const proposal = f.proposals.pending()!
  f.edit('Later')
  await expect(f.proposals.confirm(proposal.id)).rejects.toThrow()
  expect(
    f.request.mock.calls.some(
      ([b]) => (b as Record<string, unknown>).operation === 'preference_save_observation',
    ),
  ).toBe(false)
  f.edit('After')
  await f.call('save_presentation_observed_preference', {
    project_id: 'p',
    observation_id: id,
    preference: 'Keep',
  })
  f.skill.clear()
  expect(f.proposals.pending()).toBeUndefined()
})
it('blocks preference approval when the host shape was rotated after the observed snapshot', async () => {
  const f = durableFixture()
  const id = await f.begin()
  f.edit('After')
  await f.complete(id)
  const proposed = await f.call('save_presentation_observed_preference', {
    project_id: 'p',
    observation_id: id,
    preference: 'Keep',
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const proposal = f.proposals.pending()!
  f.rotate(45)
  await expect(f.proposals.confirm(proposal.id)).rejects.toThrow('proposal_stale')
  expect(
    f.request.mock.calls.some(
      ([body]) => (body as Record<string, unknown>).operation === 'preference_save_observation',
    ),
  ).toBe(false)
})
it('delete requires confirmation and binds the captured snapshots', async () => {
  const f = durableFixture(),
    id = await f.begin()
  await f.call('delete_presentation_edit_observation', { project_id: 'p', observation_id: id })
  expect(
    f.request.mock.calls.some(
      ([b]) => (b as Record<string, unknown>).operation === 'manual_observation_delete',
    ),
  ).toBe(false)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.request).toHaveBeenCalledWith(
    expect.objectContaining({
      operation: 'manual_observation_delete',
      expectedBeforeDigest: f.get().before.digest,
      expectedAfterDigest: null,
    }),
    expect.any(AbortSignal),
  )
})
it('unchanged observations cannot become approved preferences', async () => {
  const f = durableFixture(),
    id = await f.begin()
  await f.complete(id)
  const result = await f.call('save_presentation_observed_preference', {
    project_id: 'p',
    observation_id: id,
    preference: 'Keep',
  })
  expect(result.isError).toBe(true)
  expect(f.proposals.pending()).toBeUndefined()
})
it('large valid snapshots produce a bounded explicitly truncated proposal', async () => {
  const f = durableFixture()
  f.edit('a'.repeat(30000))
  const id = await f.begin()
  f.edit('b'.repeat(30000))
  await f.complete(id)
  const result = await f.call('save_presentation_observed_preference', {
    project_id: 'p',
    observation_id: id,
    preference: 'Keep',
  })
  expect(result.isError, result.output).not.toBe(true)
  const preview = f.proposals.pending()!.preview
  expect(preview.textPreviewTruncated).toBe(true)
  expect(new TextEncoder().encode(JSON.stringify(preview)).length).toBeLessThan(4096)
})
it.each([false, true])(
  'confirmation uses immutable input even when scope fields mutate: %s',
  async (mutateIdentity) => {
    const f = durableFixture(),
      id = await f.begin()
    f.edit('After')
    await f.complete(id)
    const input = { project_id: 'p', observation_id: id, preference: 'Original preview suggestion' }
    const call = { id: 'capture', name: 'save_presentation_observed_preference', input }
    const result = await f.skill.executeTool(call)
    expect(result.isError, result.output).not.toBe(true)
    const proposal = f.proposals.pending()!
    expect(proposal.preview.preference).toBe('Original preview suggestion')
    input.preference = 'Mutated unapproved text'
    if (mutateIdentity) {
      input.observation_id = 'other'
      input.project_id = 'other-project'
      call.name = 'delete_presentation_edit_observation'
    }
    await f.proposals.confirm(proposal.id)
    expect(f.request).toHaveBeenCalledWith(
      expect.objectContaining({
        operation: 'preference_save_observation',
        projectId: 'p',
        observationId: id,
        text: 'Original preview suggestion',
      }),
      expect.any(AbortSignal),
    )
    expect(
      f.request.mock.calls.some(
        ([body]) => (body as Record<string, unknown>).operation === 'manual_observation_delete',
      ),
    ).toBe(false)
  },
)
