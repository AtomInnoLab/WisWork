import {
  createDocumentCarrierIssuer,
  type PreparedResponsesTurn,
  type ProtocolLimits,
} from '../../src/index.js'

export function prepareCarrierTurn(
  input: unknown,
  limits: Partial<ProtocolLimits> = {},
): PreparedResponsesTurn {
  const issuer = createDocumentCarrierIssuer(
    {
      host: 'slides',
      documentId: 'document-1',
      sessionId: 'session_1',
      generation: 7,
    },
    (capability) => capability === 'fixture-proof',
  )
  const handle = issuer.issueForTurn({
    turnId: 'turn_1',
    sourceNonce: 'N'.repeat(43),
    capability: 'fixture-proof',
    method: 'mcp__wiswork__wiswork_read_document',
    toolName: 'wiswork_read_document',
    schemaDigest: 'a'.repeat(64),
  })
  return issuer.prepareTurn(input, limits, handle)
}
