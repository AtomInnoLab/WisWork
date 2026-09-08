import { writeFileSync } from 'node:fs'
import { buildPresentationDesignDocument } from '@wiswork/agent-core'

const pendingDesigns = new Map<number, string>()

export function savePresentationDesignSidecar(
  senderId: number,
  deckPath: string | undefined,
  style: string,
): boolean {
  pendingDesigns.set(senderId, style)
  if (!deckPath || !deckPath.endsWith('.pptx')) return true
  return flushPresentationDesignSidecar(senderId, deckPath)
}

export function flushPresentationDesignSidecar(senderId: number, deckPath: string): boolean {
  const style = pendingDesigns.get(senderId)
  if (!style || !deckPath.endsWith('.pptx')) return false
  try {
    writeFileSync(
      deckPath.replace(/\.pptx$/i, '.design.md'),
      buildPresentationDesignDocument(style),
    )
    pendingDesigns.delete(senderId)
    return true
  } catch {
    return false
  }
}

export function clearPresentationDesignSidecar(senderId: number): void {
  pendingDesigns.delete(senderId)
}
