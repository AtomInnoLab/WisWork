import { createDocumentCarrierAuthorization } from '../../src/index.js'

export const carrierAuthorization = createDocumentCarrierAuthorization({
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
