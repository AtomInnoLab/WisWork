import { expect, it, vi } from 'vitest'
import type { PowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { durableCompatibility } from './powerpoint-durable-fixture'

it.each(['\r', '\v', '\r\n', 'wrong text'])(
  'verifies durable text readback with %j while retaining the original-page backup',
  async (separator) => {
    const expected = 'First\nSecond\nThird'
    let text = 'Hello'
    let changeId = ''
    const shapes = [
      { id: 'text-1', name: 'Title', type: 'TextBox', left: 10, top: 20, width: 200, height: 80 },
    ]
    const execute = vi.fn(async () => {
      expect(binding.readExistingBatch(changeId)).toMatchObject({
        version: 3,
        state: 'applying',
        inFlightIndex: 0,
        backups: [
          expect.objectContaining({ sizeBytes: expect.any(Number), sha256: expect.any(String) }),
        ],
      })
      text =
        separator === 'wrong text' ? 'First\nDifferent\nThird' : expected.replace(/\n/g, separator)
      return { createdShapeIds: [] }
    })
    const adapter = {
      listSlideShapes: async () => ({ slideId: 'slide-1', slideIndex: 0, shapes }),
      readSlideText: async () => ({
        slideId: 'slide-1',
        shapeId: 'text-1',
        text,
        paragraphs: text.split(/\r\n|\r|\v|\n/),
      }),
      executeDeclarative: execute,
    } as unknown as PowerPointAdapter
    const proposals = createStructuredProposalController()
    const fixture = await durableCompatibility(adapter, proposals)
    const binding = fixture.binding
    const result = await fixture.skill.executeTool({
      id: 'normalize',
      name: 'execute_office_js',
      input: {
        program: {
          version: 1,
          operations: [
            { op: 'set_shape_text', slide_index: 0, shape_id: 'text-1', text: expected },
          ],
        },
      },
    })
    expect(result.isError).not.toBe(true)
    changeId = JSON.parse(result.output).changeId
    expect(execute).not.toHaveBeenCalled()
    if (separator === 'wrong text') {
      await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow(
        'office_verify_failed',
      )
      expect(binding.readExistingBatch(changeId)).toMatchObject({
        state: 'applying',
        inFlightIndex: 0,
        nextIndex: 0,
      })
    } else {
      await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
      expect(binding.readExistingBatch(changeId)).toMatchObject({ state: 'applied', nextIndex: 1 })
    }
    expect(execute).toHaveBeenCalledOnce()
    expect(binding.readExistingBatch(changeId)).toMatchObject({
      backups: [expect.objectContaining({ sizeBytes: expect.any(Number) })],
    })
  },
)
