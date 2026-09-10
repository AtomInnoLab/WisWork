import { describe, expect, it, vi } from 'vitest'
import { createOfficeDesignRequest } from '../src/relay/design-document.js'

const documentId = 'design_document_123'
const markdown = '# DESIGN.md\n\n火山与海洋。'
const revision = 'a'.repeat(64)
const valid = { documentId, markdown, revision }

function client(response: Response) {
  const capabilityFetch = vi.fn(
    async (_capability: string, _body: unknown, _signal?: AbortSignal) => response,
  )
  return { request: createOfficeDesignRequest({ capabilityFetch }), capabilityFetch }
}

describe('connected-PC DESIGN.md request boundary', () => {
  it.each(['open', 'read'] as const)(
    'forwards only the %s request through the negotiated capability with the caller signal',
    async (action) => {
      const { request, capabilityFetch } = client(Response.json(valid))
      const controller = new AbortController()
      const body = action === 'open' ? { action, documentId, markdown } : { action, documentId }

      await expect(request(body, controller.signal)).resolves.toEqual(valid)
      expect(capabilityFetch).toHaveBeenCalledExactlyOnceWith(
        'design-document.v1',
        body,
        controller.signal,
      )
    },
  )

  it('returns only document content and revision, never PC paths or unrelated response fields', async () => {
    const { request } = client(
      Response.json({
        ...valid,
        path: '/private/pc/design/DESIGN.md',
        sessionId: 'another-session',
        arbitrary: { action: 'execute' },
      }),
    )
    const result = await request({ action: 'read', documentId })
    expect(result).toEqual(valid)
    expect(Object.keys(result).sort()).toEqual(['documentId', 'markdown', 'revision'])
  })

  it.each([
    ['wrong document id', { ...valid, documentId: 'other_document_123' }],
    ['missing document id', { markdown, revision }],
    ['null document', null],
    ['array document', [valid]],
    ['non-string Markdown', { ...valid, markdown: 17 }],
    ['empty Markdown', { ...valid, markdown: '' }],
    ['whitespace Markdown', { ...valid, markdown: ' \n\t' }],
    ['NUL in Markdown', { ...valid, markdown: '# DESIGN\0private' }],
    ['oversize ASCII Markdown', { ...valid, markdown: 'x'.repeat(96 * 1024 + 1) }],
    ['oversize UTF-8 Markdown', { ...valid, markdown: '海'.repeat(32 * 1024 + 1) }],
    ['missing digest', { documentId, markdown }],
    ['non-string digest', { ...valid, revision: 123 }],
    ['short digest', { ...valid, revision: 'a'.repeat(63) }],
    ['long digest', { ...valid, revision: 'a'.repeat(65) }],
    ['uppercase digest', { ...valid, revision: 'A'.repeat(64) }],
    ['non-hex digest', { ...valid, revision: 'g'.repeat(64) }],
  ])('rejects %s', async (_label, value) => {
    const { request } = client(Response.json(value))
    await expect(request({ action: 'read', documentId })).rejects.toThrow(
      'design_document_unavailable',
    )
  })

  it.each(['', '{"documentId":', 'not JSON'])('rejects malformed JSON %#', async (text) => {
    const { request } = client(new Response(text))
    await expect(request({ action: 'read', documentId })).rejects.toThrow()
  })

  it('rejects an oversized response even when the Markdown itself is bounded', async () => {
    const { request } = client(Response.json({ ...valid, padding: 'x'.repeat(256 * 1024) }))
    await expect(request({ action: 'read', documentId })).rejects.toThrow(
      'design_document_unavailable',
    )
  })

  it.each([401, 403, 413, 500, 503])('rejects HTTP %i without reading content', async (status) => {
    const response = new Response('upstream details', { status })
    const read = vi.spyOn(response, 'text')
    const { request } = client(response)
    await expect(request({ action: 'read', documentId })).rejects.toThrow(
      'design_document_unavailable',
    )
    expect(read).not.toHaveBeenCalled()
  })

  it('accepts Markdown at the UTF-8 byte limit without truncating it', async () => {
    const exactMarkdown = '海'.repeat(32 * 1024)
    const { request } = client(Response.json({ ...valid, markdown: exactMarkdown }))
    await expect(request({ action: 'read', documentId })).resolves.toEqual({
      ...valid,
      markdown: exactMarkdown,
    })
  })
})
