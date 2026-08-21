import { describe, expect, it, vi } from 'vitest'
import type { OfficeRelaySession } from '../src/relay/session.js'
import { createOfficeWebSkill } from '../src/skills/shared/web-skill.js'

const session = (capabilities: readonly string[], fetch = vi.fn()): OfficeRelaySession =>
  ({
    snapshot: () => ({ status: 'connected', capabilities }),
    subscribe: () => () => {},
    connect: async () => {},
    disconnect() {},
    authenticatedFetch: vi.fn(),
    capabilityFetch: fetch,
  }) as unknown as OfficeRelaySession

describe('Office authenticated web skill', () => {
  it('advertises only capabilities negotiated with PC and Relay', () => {
    const skill = createOfficeWebSkill(session(['agent.v1', 'web-search.v1']))
    expect(skill.tools.map((tool) => tool.name)).toEqual(['web_search'])
  })

  it('sends an exact typed request and returns bounded safe JSON', async () => {
    const fetch = vi.fn(
      async () =>
        new Response('{"results":[]}', { headers: { 'content-type': 'application/json' } }),
    )
    const skill = createOfficeWebSkill(session(['web-search.v1'], fetch))
    const result = await skill.executeTool(
      {
        id: 'call_12345678',
        name: 'web_search',
        input: { query: 'office agents', max_results: 5 },
      },
      new AbortController().signal,
    )
    expect(fetch).toHaveBeenCalledWith(
      'web-search.v1',
      { query: 'office agents', max_results: 5 },
      expect.any(AbortSignal),
    )
    expect(result).toEqual({
      output: '{"results":[]}',
      mutated: false,
      summary: 'web_search',
    })
  })

  it('rejects extra fields and unsupported tools without sending', async () => {
    const fetch = vi.fn()
    const skill = createOfficeWebSkill(session(['web-fetch.v1'], fetch))
    const signal = new AbortController().signal
    expect(
      await skill.executeTool(
        { id: 'call_12345678', name: 'web_fetch', input: { url: 'https://example.com', extra: 1 } },
        signal,
      ),
    ).toEqual(expect.objectContaining({ output: 'web_retrieval_failed', isError: true }))
    expect(
      await skill.executeTool(
        { id: 'call_12345679', name: 'web_search', input: { query: 'x', max_results: 1 } },
        signal,
      ),
    ).toEqual(expect.objectContaining({ output: 'invalid_tool_input', isError: true }))
    expect(fetch).not.toHaveBeenCalled()
  })
})
