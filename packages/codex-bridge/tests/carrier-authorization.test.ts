import { describe, expect, it, vi } from 'vitest'
import captured from './fixtures/codex-0147-request.json'
import { createDocumentCarrierIssuer, ProtocolCompatibilityError } from '../src/index.js'

const context = () => ({
  host: 'slides',
  documentId: 'document-1',
  sessionId: 'session_1',
  generation: 7,
})

const turn = () => ({
  turnId: 'turn_1',
  sourceNonce: 'N'.repeat(43),
  capability: { proof: 'private-proof' },
  method: 'mcp__wiswork__wiswork_read_document',
  toolName: 'wiswork_read_document',
  schemaDigest: 'a'.repeat(64),
})

const issuer = () =>
  createDocumentCarrierIssuer(context(), (capability) =>
    capability === undefined ? false : (capability as { proof?: string }).proof === 'private-proof',
  )

async function* chunks(...values: string[]): AsyncGenerator<string> {
  yield* values
}

async function consume(source: AsyncIterable<string>): Promise<void> {
  for await (const _frame of source) void _frame
}

describe('closed document carrier issuer', () => {
  it('keeps authority private and atomically consumes one handle', () => {
    const owner = issuer()
    const handle = owner.issueForTurn(turn())
    expect(Object.keys(handle)).toEqual([])
    expect(Object.getOwnPropertySymbols(handle)).toEqual([])
    expect(JSON.stringify(handle)).toBe('{}')

    const prepared = owner.prepareTurn(structuredClone(captured), {}, handle)
    expect(prepared.messagesRequest.tools).toHaveLength(1)
    expect(Object.keys(prepared).sort()).toEqual(['messagesRequest', 'messagesStreamToResponses'])
    expect(() => owner.prepareTurn(structuredClone(captured), {}, handle)).toThrow(
      'carrier_authorization_consumed',
    )
  })

  it('burns an authentic handle before parsing untrusted request data', () => {
    const owner = issuer()
    const handle = owner.issueForTurn(turn())
    expect(() => owner.prepareTurn({ model: 'invalid' }, {}, handle)).toThrow()
    expect(() => owner.prepareTurn(structuredClone(captured), {}, handle)).toThrow(
      'carrier_authorization_consumed',
    )
  })

  it('rejects cross-issuer, forged, cloned, and metadata-mutated handles', () => {
    const first = issuer()
    const second = issuer()
    const handle = first.issueForTurn(turn())
    expect(() => second.prepareTurn(structuredClone(captured), {}, handle)).toThrow(
      'carrier_authorization_issuer_mismatch',
    )
    expect(() => second.prepareTurn(structuredClone(captured), {}, {} as never)).toThrow(
      'invalid_carrier_authorization',
    )
    expect(() => first.prepareTurn(structuredClone(captured), {}, { ...handle })).toThrow(
      'invalid_carrier_authorization',
    )

    const body = structuredClone(captured) as any
    const metadata = JSON.parse(body.client_metadata['x-codex-turn-metadata'])
    metadata.turn_id = 'other-turn'
    body.client_metadata['x-codex-turn-metadata'] = JSON.stringify(metadata)
    expect(() => first.prepareTurn(body, {}, handle)).toThrow('carrier_authorization_mismatch')
    expect(() => first.prepareTurn(structuredClone(captured), {}, handle)).toThrow(
      'carrier_authorization_consumed',
    )
  })

  it('validates private capability proof before issuing without retaining it publicly', () => {
    const validate = vi.fn((capability: unknown) => capability === 'approved')
    const owner = createDocumentCarrierIssuer(context(), validate)
    expect(() => owner.issueForTurn({ ...turn(), capability: 'denied' })).toThrow(
      'carrier_capability_rejected',
    )
    const handle = owner.issueForTurn({ ...turn(), capability: 'approved' })
    expect(validate).toHaveBeenCalledTimes(2)
    expect(JSON.stringify(handle)).not.toContain('approved')
  })

  it('enforces the bounded call budget in the private issuer ledger', async () => {
    const owner = issuer()
    const prepared = owner.prepareTurn(structuredClone(captured), {}, owner.issueForTurn(turn()))
    const frames = [
      'event: message_start\ndata: {"type":"message_start","message":{"id":"msg","model":"openai/gpt-5.6-sol","usage":{"input_tokens":1}}}\n\n',
      'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"call","name":"exec","input":{}}}\n\n',
      `event: content_block_delta\ndata: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: { type: 'input_json_delta', partial_json: JSON.stringify({ code: 'await tools.mcp__wiswork__wiswork_read_document({})' }) } })}\n\n`,
      'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
      'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"tool_use"},"usage":{"output_tokens":1}}\n\n',
      'event: message_stop\ndata: {"type":"message_stop"}\n\n',
    ]
    for (let index = 0; index < 16; index += 1) {
      await expect(
        consume(prepared.messagesStreamToResponses(chunks(...frames))),
      ).resolves.toBeUndefined()
    }
    await expect(consume(prepared.messagesStreamToResponses(chunks(...frames)))).rejects.toThrow(
      'tool_call_limit_exceeded',
    )
  })

  it.each([
    { ...context(), host: 'unknown' },
    { ...context(), generation: -1 },
    { ...context(), extra: true },
    Object.defineProperty(context(), 'documentId', { get: () => 'document-1' }),
  ])('rejects malformed issuer context', (value) => {
    expect(() => createDocumentCarrierIssuer(value, () => true)).toThrowError(
      expect.objectContaining<Partial<ProtocolCompatibilityError>>({
        message: 'invalid_carrier_issuer',
      }),
    )
  })

  it.each([
    { ...turn(), turnId: '' },
    { ...turn(), sourceNonce: 'short' },
    { ...turn(), schemaDigest: 'not-a-digest' },
    { ...turn(), method: 'exec_command' },
    { ...turn(), extra: true },
    Object.defineProperty(turn(), 'turnId', { get: () => 'turn_1' }),
  ])('rejects malformed turn authority', (value) => {
    expect(() => issuer().issueForTurn(value)).toThrow('invalid_carrier_authorization')
  })
})
