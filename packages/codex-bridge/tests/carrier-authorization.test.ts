import { describe, expect, it } from 'vitest'
import captured from './fixtures/codex-0147-request.json'
import {
  createDocumentCarrierAuthorization,
  prepareResponsesTurn,
  ProtocolCompatibilityError,
} from '../src/index.js'

const rawAuthorization = () => ({
  host: 'slides',
  documentId: 'document-1',
  sessionId: 'session_1',
  generation: 7,
  capabilityToken: 'A'.repeat(43),
  method: 'mcp__wiswork__wiswork_read_document',
  toolName: 'wiswork_read_document',
  schemaDigest: 'a'.repeat(64),
  callBudget: 1,
})

describe('opaque document carrier authorization', () => {
  it('requires caller-created authority before exposing the one-call carrier', () => {
    const withoutAuthority = prepareResponsesTurn(structuredClone(captured))
    expect(withoutAuthority.messagesRequest.tools).toBeUndefined()

    const authority = createDocumentCarrierAuthorization(rawAuthorization())
    const prepared = prepareResponsesTurn(structuredClone(captured), {}, authority)
    expect(prepared.messagesRequest.tools).toHaveLength(1)
    expect(prepared.carrierAuthorization).toBe(authority)
  })

  it.each([
    { ...rawAuthorization(), host: 'unknown' },
    { ...rawAuthorization(), generation: -1 },
    { ...rawAuthorization(), callBudget: 2 },
    { ...rawAuthorization(), schemaDigest: 'not-a-digest' },
    { ...rawAuthorization(), method: 'exec_command' },
    { ...rawAuthorization(), extra: true },
    Object.defineProperty(rawAuthorization(), 'documentId', { get: () => 'document-1' }),
  ])('rejects malformed or expandable caller authority', (value) => {
    expect(() => createDocumentCarrierAuthorization(value)).toThrowError(
      expect.objectContaining<Partial<ProtocolCompatibilityError>>({
        message: 'invalid_carrier_authorization',
      }),
    )
  })

  it('rejects app-server metadata that advertises a different document method', () => {
    const body = structuredClone(captured) as any
    const metadata = JSON.parse(body.client_metadata['x-codex-turn-metadata'])
    metadata.code_mode_tool_names.mcp__wiswork__foreign_document = {
      name: 'foreign_document',
      namespace: 'mcp__wiswork',
    }
    body.client_metadata['x-codex-turn-metadata'] = JSON.stringify(metadata)
    expect(() =>
      prepareResponsesTurn(body, {}, createDocumentCarrierAuthorization(rawAuthorization())),
    ).toThrow('carrier_authorization_mismatch')
  })
})
