import { afterEach, expect, it, vi } from 'vitest'
import {
  BrowserPowerPointAdapter,
  MAX_POWERPOINT_TEXT,
} from '../src/skills/powerpoint/browser-powerpoint-adapter.js'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { createOfficeDiagnostics } from '../src/diagnostics/office-diagnostics.js'

afterEach(() => vi.unstubAllGlobals())

function officeFixture(options: { ignoredStyle?: boolean; mixedStyle?: boolean } = {}) {
  const characterFonts = [
    { color: '#000000', name: 'Aptos', size: 32, bold: false, italic: false },
    {
      color: '#000000',
      name: 'Aptos',
      size: options.mixedStyle ? 18 : 32,
      bold: false,
      italic: false,
    },
  ]
  const font = new Proxy({ load: vi.fn() } as Record<string, unknown>, {
    get(target, key) {
      if (key === 'load') return target.load
      const values = characterFonts.map((value) => value[key as keyof typeof value])
      return values.every((value) => value === values[0]) ? values[0] : null
    },
    set(_target, key, value) {
      if (!options.ignoredStyle)
        for (const character of characterFonts) Reflect.set(character, key, value)
      return true
    },
  })
  const range = {
    text: 'AB',
    font,
    load: vi.fn(),
    getSubstring: vi.fn((start: number, length: number) => ({
      font: length === 1 ? { ...characterFonts[start], load: vi.fn() } : font,
    })),
  }
  const title = {
    id: 'title',
    name: 'Title',
    type: 'TextBox',
    left: 10,
    top: 10,
    width: 300,
    height: 40,
    load: vi.fn(),
    textFrame: { hasText: true, load: vi.fn(), textRange: range },
  }
  const placeholder = {
    id: 'empty',
    name: 'Empty placeholder',
    type: 'Placeholder',
    left: 10,
    top: 70,
    width: 300,
    height: 40,
    load: vi.fn(),
    delete: () => {
      shapes.items = shapes.items.filter((value) => value.id !== 'empty')
    },
  }
  const shapes = {
    items: [title, placeholder],
    load: vi.fn(),
    getItem: (id: string) => shapes.items.find((value) => value.id === id),
  }
  const slide = { id: 'slide-1', shapes, load: vi.fn() }
  const context = {
    presentation: {
      slides: {
        items: [slide],
        load: vi.fn(),
        getCount: () => ({ value: 1 }),
        getItemAt: () => slide,
      },
      pageSetup: { load: vi.fn(), slideWidth: 960, slideHeight: 540 },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  }
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', platform: 'Mac', requirements: { isSetSupported: () => true } },
  })
  vi.stubGlobal('PowerPoint', {
    run: (callback: (ctx: typeof context) => unknown) => callback(context),
  })
  const adapter = new BrowserPowerPointAdapter()
  const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'test' })
  const proposals = createStructuredProposalController(diagnostics)
  const skill = createPowerPointSkill({ adapter, proposals })
  return { adapter, proposals, skill, characterFonts, range, shapes, diagnostics }
}

it('confirms style and delete on different shapes in one batch', async () => {
  const fixture = officeFixture()
  const result = await fixture.skill.executeTool({
    id: 'probe',
    name: 'execute_office_js',
    input: {
      program: {
        version: 1,
        operations: [
          {
            op: 'set_shape_text_style',
            slide_index: 0,
            shape_id: 'title',
            fontSize: 34,
            fontFamily: 'Aptos Display',
            color: '#2457A7',
            bold: true,
          },
          { op: 'delete_shape', slide_index: 0, shape_id: 'empty' },
        ],
      },
    },
  })
  expect(result.isError).not.toBe(true)
  await expect(fixture.proposals.confirm(fixture.proposals.pending()!.id)).resolves.toBeUndefined()
  expect(fixture.shapes.items.map((shape) => shape.id)).toEqual(['title'])
  expect(fixture.characterFonts.every((font) => font.size === 34 && font.bold)).toBe(true)
})

it('does not approve full-range style when only the first character matches', async () => {
  const fixture = officeFixture({ ignoredStyle: true, mixedStyle: true })
  await fixture.skill.executeTool({
    id: 'probe',
    name: 'execute_office_js',
    input: {
      program: {
        version: 1,
        operations: [
          { op: 'set_shape_text_style', slide_index: 0, shape_id: 'title', fontSize: 32 },
        ],
      },
    },
  })
  await expect(fixture.proposals.confirm(fixture.proposals.pending()!.id)).rejects.toThrow(
    'office_verify_failed',
  )
})

it.each(['bold', 'italic'] as const)(
  'preserves mixed %s as unknown instead of false',
  async (property) => {
    const fixture = officeFixture()
    fixture.characterFonts[0][property] = true
    fixture.characterFonts[1][property] = false
    const readback = await fixture.adapter.readShapeTextStyle(0, 'title')
    expect(readback[property]).toBeUndefined()
  },
)

it('fails closed instead of verifying a truncated prefix of an oversized text range', async () => {
  const fixture = officeFixture()
  fixture.range.text = 'A'.repeat(MAX_POWERPOINT_TEXT + 1)
  await expect(fixture.adapter.readShapeTextStyle(0, 'title')).rejects.toThrow('office_read_failed')
  expect(fixture.range.getSubstring).not.toHaveBeenCalled()
})

it.each([
  { property: 'color', style: { color: '#2457A7' } },
  { property: 'fontFamily', style: { fontFamily: 'Aptos Display' } },
  { property: 'fontSize', style: { fontSize: 34 } },
  { property: 'bold', style: { bold: true } },
  { property: 'italic', style: { italic: true } },
])(
  'reports the failing $property without retrying the applied batch',
  async ({ property, style }) => {
    const fixture = officeFixture({ ignoredStyle: true })
    const readback = vi.spyOn(fixture.adapter, 'readShapeTextStyle')
    const execute = vi.spyOn(fixture.adapter, 'executeDeclarative')
    await fixture.skill.executeTool({
      id: 'probe',
      name: 'execute_office_js',
      input: {
        program: {
          version: 1,
          operations: [
            { op: 'set_shape_text_style', slide_index: 0, shape_id: 'title', ...style },
            { op: 'delete_shape', slide_index: 0, shape_id: 'empty' },
          ],
        },
      },
    })
    const proposalId = fixture.proposals.pending()!.id
    const decision = fixture.proposals.waitForDecision(proposalId)
    await expect(fixture.proposals.confirm(proposalId)).rejects.toMatchObject({
      message: 'office_verify_failed',
      debugInfo: { errorLocation: `PowerPoint.operations.0.set_shape_text_style.${property}` },
    })
    expect(await decision).toMatchObject({
      status: 'failed',
      error: 'office_verify_failed',
      errorLocation: `PowerPoint.operations.0.set_shape_text_style.${property}`,
    })
    expect(readback).toHaveBeenCalledTimes(3)
    expect(execute).toHaveBeenCalledTimes(1)
    expect(fixture.shapes.items.map((shape) => shape.id)).toEqual(['title'])
    expect(fixture.diagnostics.snapshot().events.at(-1)).toMatchObject({
      phase: 'verify',
      error_code: 'office_verify_failed',
      office_error_location: `PowerPoint.operations.0.set_shape_text_style.${property}`,
    })
  },
)

it('still accepts converging style readback without repeating the mutation', async () => {
  const fixture = officeFixture()
  const execute = vi.spyOn(fixture.adapter, 'executeDeclarative')
  const readback = vi
    .spyOn(fixture.adapter, 'readShapeTextStyle')
    .mockResolvedValueOnce({ fontSize: 32 } as Awaited<
      ReturnType<typeof fixture.adapter.readShapeTextStyle>
    >)
    .mockResolvedValueOnce({ fontSize: 32 } as Awaited<
      ReturnType<typeof fixture.adapter.readShapeTextStyle>
    >)
  await fixture.skill.executeTool({
    id: 'probe',
    name: 'execute_office_js',
    input: {
      program: {
        version: 1,
        operations: [
          { op: 'set_shape_text_style', slide_index: 0, shape_id: 'title', fontSize: 34 },
        ],
      },
    },
  })
  await expect(fixture.proposals.confirm(fixture.proposals.pending()!.id)).resolves.toBeUndefined()
  expect(readback).toHaveBeenCalledTimes(3)
  expect(execute).toHaveBeenCalledTimes(1)
})
