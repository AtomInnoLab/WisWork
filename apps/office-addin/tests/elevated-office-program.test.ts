import { describe, expect, it, vi } from 'vitest'
import {
  createElevatedOfficeSkill,
  parseElevatedOfficeProgram,
  type ElevatedOfficeAdapter,
} from '../src/skills/shared/elevated-office-program.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'

const authority = () => ({
  activeMode: 'enhanced' as const,
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
})

function fixture(host: 'word' | 'excel' | 'powerpoint' = 'word') {
  let state = 'before'
  const adapter: ElevatedOfficeAdapter = {
    host,
    captureAuthority: vi.fn(authority),
    snapshot: vi.fn(async () => ({ id: 'history_AAAAAAAAAAAAAAAA', state })),
    validateSnapshot: vi.fn(async () => true),
    execute: vi.fn(async (_program, _snapshot, _signal, lifecycle) => {
      lifecycle?.markStarted()
      state = 'after'
      lifecycle?.markApplied()
    }),
    readback: vi.fn(async () => ({ verified: state === 'after', output: { state } })),
    rollback: vi.fn(async () => {
      state = 'before'
    }),
  }
  const proposals = createStructuredProposalController()
  return { adapter, proposals, skill: createElevatedOfficeSkill({ host, adapter, proposals }) }
}

describe('elevated raw Office program', () => {
  it('parses a bounded closed Word AST and rejects executable ambient authority', () => {
    expect(
      parseElevatedOfficeProgram('word', {
        version: 1,
        kind: 'office_js_ast',
        operations: [{ call: 'body.insertText', args: { location: 'end', text: 'Hello' } }],
      }),
    ).toMatchObject({ kind: 'office_js_ast' })
    for (const source of [
      'fetch("https://x")',
      'new Function("return 1")',
      'eval("1")',
      'import("x")',
      'localStorage.x',
      'navigator.clipboard',
      'setTimeout(x)',
      'window.open("x")',
    ]) {
      expect(() => parseElevatedOfficeProgram('word', source)).toThrow('raw_office_program_invalid')
    }
  })

  it('enforces OOXML bytes, nodes, depth and rejects external relationships', () => {
    expect(() =>
      parseElevatedOfficeProgram('powerpoint', {
        version: 1,
        kind: 'ooxml_patch',
        patches: [
          {
            part: 'ppt/slides/slide1.xml',
            xml: '<p:sld xmlns:p="p"><p:cSld><p:spTree/></p:cSld></p:sld>',
          },
        ],
      }),
    ).not.toThrow()
    expect(() =>
      parseElevatedOfficeProgram('powerpoint', {
        version: 1,
        kind: 'ooxml_patch',
        patches: [
          {
            part: 'ppt/slides/slide1.xml',
            xml: '<Relationship TargetMode="External" Target="https://example.com"/>',
          },
        ],
      }),
    ).toThrow('raw_office_program_invalid')
    expect(() =>
      parseElevatedOfficeProgram('word', {
        version: 1,
        kind: 'ooxml_patch',
        patches: [
          { part: 'word/document.xml', xml: `<a>${'<b>'.repeat(40)}x${'</b>'.repeat(40)}</a>` },
        ],
      }),
    ).toThrow('raw_office_program_invalid')
    for (const xml of [
      '<p:sld xmlns:p="p" xmlns:a="a"><a:hlinkClick r:id="rId1"/></p:sld>',
      '<p:sld xmlns:p="p" xmlns:a="a"><a:fld>secret</a:fld></p:sld>',
      '<p:sld xmlns:p="p" xmlns:a="a"><a:instrText>LINK</a:instrText></p:sld>',
      '<p:sld xmlns:p="p" xmlns:a="a"><a:r action="ppaction://hlinkshowjump"/></p:sld>',
      '<p:sld xmlns:p="p" xmlns:a="a"><a:blip r:embed="rId1"/></p:sld>',
    ]) {
      expect(() =>
        parseElevatedOfficeProgram('powerpoint', {
          version: 1,
          kind: 'ooxml_patch',
          patches: [{ part: 'ppt/slides/slide1.xml', xml }],
        }),
      ).toThrow('raw_office_program_invalid')
    }
    expect(() =>
      parseElevatedOfficeProgram('word', {
        version: 1,
        kind: 'ooxml_patch',
        patches: [{ part: 'word/document.xml', xml: '<a><b></a>' }],
      }),
    ).toThrow('raw_office_program_invalid')
    expect(() =>
      parseElevatedOfficeProgram('excel', {
        version: 1,
        kind: 'ooxml_patch',
        patches: [{ part: 'xl/worksheets/sheet1.xml', xml: '<worksheet/>' }],
      }),
    ).toThrow('raw_office_program_invalid')
  })

  it('enforces adapter-exact program cardinality and denies Word/Excel OOXML', () => {
    expect(() =>
      parseElevatedOfficeProgram('excel', {
        version: 1,
        kind: 'office_js_ast',
        operations: [
          { call: 'range.clear', args: { sheetId: 1, range: 'A1', clearType: 'contents' } },
          { call: 'range.clear', args: { sheetId: 1, range: 'A2', clearType: 'contents' } },
        ],
      }),
    ).toThrow('raw_office_program_invalid')
    for (const host of ['word', 'excel'] as const)
      expect(() =>
        parseElevatedOfficeProgram(host, {
          version: 1,
          kind: 'ooxml_patch',
          patches: [
            {
              part: host === 'word' ? 'word/document.xml' : 'xl/worksheets/sheet1.xml',
              xml: '<x/>',
            },
          ],
        }),
      ).toThrow('raw_office_program_invalid')
    expect(() =>
      parseElevatedOfficeProgram('powerpoint', {
        version: 1,
        kind: 'ooxml_patch',
        patches: [
          { part: 'ppt/slides/slide1.xml', xml: '<p:sld xmlns:p="p"/>' },
          { part: 'ppt/slides/slide2.xml', xml: '<p:sld xmlns:p="p"/>' },
        ],
      }),
    ).toThrow('raw_office_program_invalid')
  })

  it('rejects Excel formula-like literals after BOM, Unicode whitespace, or controls', () => {
    for (const value of ['=WEBSERVICE("x")', '\uFEFF+1', '\u200B-1', '\u0009@SUM(A1)']) {
      expect(() =>
        parseElevatedOfficeProgram('excel', {
          version: 1,
          kind: 'office_js_ast',
          operations: [
            { call: 'range.setValues', args: { sheetId: 1, range: 'A1', values: [[value]] } },
          ],
        }),
      ).toThrow('raw_office_program_invalid')
    }
  })

  it.each(['word', 'excel', 'powerpoint'] as const)(
    '%s requires exactly one fresh confirmation and executes through its adapter',
    async (host) => {
      const { skill, proposals, adapter } = fixture(host)
      const result = await skill.executeTool(
        {
          id: 'call_AAAAAAAAAAAAAAAA',
          name: 'propose_raw_office_edit',
          input: {
            program:
              host === 'word'
                ? {
                    version: 1,
                    kind: 'office_js_ast',
                    operations: [
                      { call: 'body.insertText', args: { location: 'end', text: 'Hello' } },
                    ],
                  }
                : host === 'excel'
                  ? {
                      version: 1,
                      kind: 'office_js_ast',
                      operations: [
                        {
                          call: 'range.setValues',
                          args: { sheetId: 1, range: 'A1', values: [['x']] },
                        },
                      ],
                    }
                  : {
                      version: 1,
                      kind: 'office_js_ast',
                      operations: [
                        {
                          call: 'shape.setText',
                          args: { slideIndex: 0, shapeId: 's1', text: 'x' },
                        },
                      ],
                    },
          },
        },
        new AbortController().signal,
      )
      expect(JSON.parse(result.output)).toMatchObject({ status: 'awaiting_user_confirmation' })
      expect(adapter.execute).not.toHaveBeenCalled()
      const proposal = proposals.pending()!
      await proposals.confirm(proposal.id)
      await expect(proposals.waitForDecision(proposal.id)).rejects.toThrow('proposal_missing')
      expect(adapter.execute).toHaveBeenCalledTimes(1)
      await expect(proposals.confirm(proposal.id)).rejects.toThrow('proposal_missing')
    },
  )

  it('fails closed before proposal when Enhanced authority is absent or raw policy is off', async () => {
    const { skill, adapter } = fixture()
    vi.mocked(adapter.captureAuthority).mockReturnValue({ ...authority(), rawOfficeEnabled: false })
    const result = await skill.executeTool(
      {
        id: 'call_AAAAAAAAAAAAAAAA',
        name: 'propose_raw_office_edit',
        input: {
          program: {
            version: 1,
            kind: 'office_js_ast',
            operations: [{ call: 'body.insertText', args: { location: 'end', text: 'x' } }],
          },
        },
      },
      new AbortController().signal,
    )
    expect(result).toMatchObject({ isError: true, mutated: false, output: 'raw_office_denied' })
  })

  it('rechecks revision after confirmation and never writes stale proposals', async () => {
    const { skill, proposals, adapter } = fixture()
    await skill.executeTool(
      {
        id: 'call_AAAAAAAAAAAAAAAA',
        name: 'propose_raw_office_edit',
        input: {
          program: {
            version: 1,
            kind: 'office_js_ast',
            operations: [{ call: 'body.insertText', args: { location: 'end', text: 'x' } }],
          },
        },
      },
      new AbortController().signal,
    )
    vi.mocked(adapter.captureAuthority).mockReturnValue({
      ...authority(),
      revision: 'rev_BBBBBBBBBBBBBBBB',
    })
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(adapter.execute).not.toHaveBeenCalled()
  })

  it('reports applied_unverified and does not retry when readback is uncertain', async () => {
    const { skill, proposals, adapter } = fixture()
    vi.mocked(adapter.readback).mockRejectedValue(new Error('office_state_uncertain'))
    const result = await skill.executeTool(
      {
        id: 'call_AAAAAAAAAAAAAAAA',
        name: 'propose_raw_office_edit',
        input: {
          program: {
            version: 1,
            kind: 'office_js_ast',
            operations: [{ call: 'body.insertText', args: { location: 'end', text: 'x' } }],
          },
        },
      },
      new AbortController().signal,
    )
    const proposal = proposals.pending()!
    const settled = proposals.waitForDecision(proposal.id)
    await expect(proposals.confirm(proposal.id)).resolves.toBeUndefined()
    await expect(settled).resolves.toMatchObject({ status: 'applied_unverified' })
    const decision = await JSON.parse(result.output)
    expect(decision.status).toBe('awaiting_user_confirmation')
    expect(adapter.execute).toHaveBeenCalledTimes(1)
  })

  it('rolls back a deterministic failed readback using the captured history snapshot', async () => {
    const { skill, proposals, adapter } = fixture()
    vi.mocked(adapter.readback).mockResolvedValue({ verified: false })
    const result = await skill.executeTool({
      id: 'call_AAAAAAAAAAAAAAAA',
      name: 'propose_raw_office_edit',
      input: {
        program: {
          version: 1,
          kind: 'office_js_ast',
          operations: [{ call: 'body.insertText', args: { location: 'end', text: 'x' } }],
        },
      },
    })
    const id = JSON.parse(result.output).proposalId
    const settled = proposals.waitForDecision(id)
    await expect(proposals.confirm(id)).rejects.toThrow('office_write_failed')
    await expect(settled).resolves.toMatchObject({ status: 'failed', error: 'office_write_failed' })
    expect(adapter.rollback).toHaveBeenCalledTimes(1)
  })

  it('does not let automatic correction invoke raw Office', async () => {
    const { adapter, proposals } = fixture('powerpoint')
    const skill = createElevatedOfficeSkill({
      host: 'powerpoint',
      adapter,
      proposals,
      automaticCorrection: true,
    })
    const result = await skill.executeTool({
      id: 'call_AAAAAAAAAAAAAAAA',
      name: 'propose_raw_office_edit',
      input: {
        program: {
          version: 1,
          kind: 'office_js_ast',
          operations: [
            { call: 'shape.setText', args: { slideIndex: 0, shapeId: 's1', text: 'x' } },
          ],
        },
      },
    })
    expect(result).toMatchObject({ isError: true, output: 'raw_office_confirmation_required' })
    expect(proposals.pending()).toBeUndefined()
    expect(adapter.execute).not.toHaveBeenCalled()
  })

  it('reconciles after cancellation races a dispatched write instead of claiming no mutation', async () => {
    const { skill, proposals, adapter } = fixture()
    let release!: () => void
    vi.mocked(adapter.execute).mockImplementation(async () => {
      await new Promise<void>((resolve) => {
        release = resolve
      })
    })
    vi.mocked(adapter.readback).mockResolvedValue({ verified: true, output: { state: 'after' } })
    await skill.executeTool({
      id: 'call_AAAAAAAAAAAAAAAA',
      name: 'propose_raw_office_edit',
      input: {
        program: {
          version: 1,
          kind: 'office_js_ast',
          operations: [{ call: 'body.insertText', args: { location: 'end', text: 'x' } }],
        },
      },
    })
    const proposal = proposals.pending()!
    const settled = proposals.waitForDecision(proposal.id)
    const confirmation = proposals.confirm(proposal.id)
    await vi.waitFor(() => expect(adapter.execute).toHaveBeenCalledTimes(1))
    proposals.logout()
    release()
    await confirmation
    await expect(settled).resolves.toMatchObject({ status: 'confirmed' })
    expect(adapter.readback).toHaveBeenCalledTimes(1)
  })

  it('returns applied_unverified when a timed-out dispatched write may still settle later', async () => {
    vi.useFakeTimers()
    try {
      const { skill, proposals, adapter } = fixture()
      let release!: () => void
      vi.mocked(adapter.execute).mockImplementation(
        (_program, _snapshot, _signal, lifecycle) =>
          new Promise<void>((resolve) => {
            lifecycle?.markStarted()
            release = resolve
          }),
      )
      vi.mocked(adapter.readback).mockResolvedValue({ verified: false })
      await skill.executeTool({
        id: 'call_AAAAAAAAAAAAAAAA',
        name: 'propose_raw_office_edit',
        input: {
          program: {
            version: 1,
            kind: 'office_js_ast',
            operations: [{ call: 'body.insertText', args: { location: 'end', text: 'x' } }],
          },
        },
      })
      const id = proposals.pending()!.id
      const settled = proposals.waitForDecision(id)
      const confirmation = proposals.confirm(id)
      await vi.advanceTimersByTimeAsync(20_000)
      await expect(confirmation).resolves.toBeUndefined()
      await expect(settled).resolves.toMatchObject({ status: 'applied_unverified' })
      release()
      await vi.runAllTimersAsync()
    } finally {
      vi.useRealTimers()
    }
  })
})
