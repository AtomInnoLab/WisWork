import { describe, expect, it, vi } from 'vitest'
import { proposalForSelection } from '../src/renderer/ai/proposal-review.js'

const proposal = {
  id: 'original',
  projectId: 'project-1',
  expiresAt: 1_000,
  files: [
    { path: 'a.tex', beforeText: 'a0', beforeSha256: 'a-hash', afterText: 'a1' },
    { path: 'b.tex', beforeText: 'b0', beforeSha256: 'b-hash', afterText: 'b1' },
  ],
}

describe('proposal review selection', () => {
  it('keeps the original id only for the unchanged full proposal', async () => {
    const create = vi.fn()
    await expect(proposalForSelection(proposal, new Set(['a.tex', 'b.tex']), create)).resolves.toBe(
      proposal,
    )
    expect(create).not.toHaveBeenCalled()
  })

  it('creates a fresh one-time proposal every time the selected content changes', async () => {
    const create = vi
      .fn()
      .mockResolvedValueOnce({ ...proposal, id: 'subset-1', files: [proposal.files[0]] })
      .mockResolvedValueOnce({ ...proposal, id: 'subset-2', files: [proposal.files[0]] })
    await expect(proposalForSelection(proposal, new Set(['a.tex']), create)).resolves.toMatchObject(
      {
        id: 'subset-1',
      },
    )
    await expect(proposalForSelection(proposal, new Set(['a.tex']), create)).resolves.toMatchObject(
      {
        id: 'subset-2',
      },
    )
    expect(create).toHaveBeenCalledTimes(2)
    expect(create).toHaveBeenLastCalledWith([{ path: 'a.tex', afterText: 'a1' }])
  })

  it('does not authorize an empty selection', async () => {
    await expect(proposalForSelection(proposal, new Set(), vi.fn())).rejects.toThrow(/select/i)
  })
})
