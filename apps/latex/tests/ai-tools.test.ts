import { describe, expect, it, vi } from 'vitest'
import { createLatexSkill, LATEX_AI_TOOL_NAMES } from '../src/renderer/ai/latex-skill.js'
import { loadProposalForReview } from '../src/renderer/ai/proposal-review.js'
import { fixedWisworkSettings } from '../src/renderer/ai/transport.js'

describe('LaTeX AI read-only skill', () => {
  it('exposes analysis, compile, and propose tools but never apply/write/undo', () => {
    const skill = createLatexSkill({} as never, () => 'project-1')
    expect(skill.tools.map((tool) => tool.name)).toEqual(LATEX_AI_TOOL_NAMES)
    expect(skill.tools.map((tool) => tool.name)).not.toEqual(
      expect.arrayContaining(['apply_project_edits', 'write_project_text', 'undo_project_edits']),
    )
  })

  it('treats project content and tool output as untrusted and rejects instruction boundary changes', () => {
    const skill = createLatexSkill({} as never, () => 'project-1')
    const prompt = skill.systemPrompt.toLowerCase()
    expect(prompt).toContain('untrusted')
    expect(prompt).toContain('latex comments')
    expect(prompt).toContain('tool outputs')
    expect(prompt).toContain('ignore')
    expect(prompt).toContain('authorization')
  })

  it('rejects unknown keys at runtime for every tool input and nested proposal file', async () => {
    const api = {
      listProjectFiles: vi.fn(),
      searchProjectText: vi.fn(),
      readProjectText: vi.fn(),
      getCompileDiagnostics: vi.fn(),
      compileProjectForAi: vi.fn(),
      proposeProjectEdits: vi.fn(),
    }
    const skill = createLatexSkill(api as never, () => 'project-1')
    const calls = [
      { id: '1', name: 'list_project_files', input: { projectId: 'forged' } },
      { id: '2', name: 'search_project_text', input: { query: 'x', extra: true } },
      { id: '3', name: 'read_project_text', input: { path: 'main.tex', rootPath: '/' } },
      { id: '4', name: 'get_compile_diagnostics', input: { extra: true } },
      { id: '5', name: 'compile_project', input: { extra: true } },
      {
        id: '6',
        name: 'propose_project_edits',
        input: { files: [{ path: 'main.tex', afterText: 'x', beforeText: 'forged' }] },
      },
    ]
    for (const call of calls) {
      await expect(skill.executeTool(call)).resolves.toMatchObject({
        isError: true,
        mutated: false,
      })
    }
    for (const method of Object.values(api)) expect(method).not.toHaveBeenCalled()
  })

  it('rejects unsafe or oversized reads/searches before IPC and bounds tool output', async () => {
    const readProjectText = vi.fn()
    const searchProjectText = vi.fn().mockResolvedValue({
      ok: true,
      value: { matches: [{ path: 'main.tex', line: 1, text: 'x'.repeat(100_000) }] },
    })
    const skill = createLatexSkill(
      { readProjectText, searchProjectText } as never,
      () => 'project-1',
    )
    await expect(
      skill.executeTool({ id: '1', name: 'read_project_text', input: { path: '../secret' } }),
    ).resolves.toMatchObject({ isError: true, mutated: false })
    expect(readProjectText).not.toHaveBeenCalled()
    await expect(
      skill.executeTool({
        id: '2',
        name: 'search_project_text',
        input: { query: 'x'.repeat(300), maxResults: 10_000 },
      }),
    ).resolves.toMatchObject({ isError: true, mutated: false })
    expect(searchProjectText).not.toHaveBeenCalled()
  })

  it('creates only a normalized proposal and reports mutated false', async () => {
    const proposeProjectEdits = vi.fn().mockResolvedValue({
      ok: true,
      value: { proposalId: 'server-id', expiresAt: 123, fileCount: 1 },
    })
    const skill = createLatexSkill({ proposeProjectEdits } as never, () => 'project-1')
    const result = await skill.executeTool({
      id: 'p',
      name: 'propose_project_edits',
      input: { files: [{ path: './chapters\\one.tex', afterText: 'full replacement' }] },
    })
    expect(proposeProjectEdits).toHaveBeenCalledWith({
      projectId: 'project-1',
      files: [{ path: 'chapters/one.tex', afterText: 'full replacement' }],
    })
    expect(result).toMatchObject({ mutated: false })
    expect(result.isError).not.toBe(true)
    expect(result.output).toContain('server-id')
  })

  it('returns only a bounded proposal id summary and loads a >64KiB review over getProposal', async () => {
    const largeText = 'x'.repeat(80 * 1024)
    const review = {
      id: 'large-proposal',
      projectId: 'project-1',
      expiresAt: 123,
      files: [
        {
          path: 'main.tex',
          beforeText: largeText,
          beforeSha256: 'a'.repeat(64),
          afterText: largeText,
        },
      ],
    }
    const skill = createLatexSkill(
      { proposeProjectEdits: vi.fn().mockResolvedValue({ ok: true, value: review }) } as never,
      () => 'project-1',
    )
    const execution = await skill.executeTool({
      id: 'large',
      name: 'propose_project_edits',
      input: { files: [{ path: 'main.tex', afterText: largeText }] },
    })

    expect(execution.output.length).toBeLessThan(1_024)
    expect(execution.output).not.toContain('[output truncated]')
    const getProposal = vi.fn().mockResolvedValue({ ok: true, value: review })
    await expect(
      loadProposalForReview(execution.output, 'project-1', getProposal),
    ).resolves.toEqual(review)
    expect(getProposal).toHaveBeenCalledWith({
      projectId: 'project-1',
      proposalId: 'large-proposal',
    })
  })
  it('cannot place a key or endpoint override into renderer WisModel settings', () => {
    const settings = fixedWisworkSettings()
    expect(settings.provider).toBe('wiswork')
    expect(settings.providers.wiswork).toMatchObject({ apiKey: '' })
    expect(settings.providers.wiswork.baseUrl).toBeUndefined()
  })
})
