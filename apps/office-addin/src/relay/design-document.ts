import type { OfficeDesignRequest } from '../OfficeDesignPanel.js'
import type { OfficeRelaySession } from './session.js'

export function createOfficeDesignRequest(
  bridge: Pick<OfficeRelaySession, 'capabilityFetch'>,
): OfficeDesignRequest {
  return async (body, signal) => {
    const response = await bridge.capabilityFetch('design-document.v1', body, signal)
    if (!response.ok) throw new Error('design_document_unavailable')
    const text = await response.text()
    if (new TextEncoder().encode(text).byteLength > 256 * 1024)
      throw new Error('design_document_unavailable')
    const value = JSON.parse(text) as Record<string, unknown>
    if (
      !value ||
      value.documentId !== body.documentId ||
      typeof value.markdown !== 'string' ||
      !value.markdown.trim() ||
      value.markdown.includes('\0') ||
      new TextEncoder().encode(value.markdown).byteLength > 96 * 1024 ||
      typeof value.revision !== 'string' ||
      !/^[a-f0-9]{64}$/.test(value.revision)
    )
      throw new Error('design_document_unavailable')
    return { documentId: body.documentId, markdown: value.markdown, revision: value.revision }
  }
}
