import { existsSync, readFileSync, writeFileSync } from 'node:fs'
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

export function readPresentationDesignSidecar(
  senderId: number,
  deckPath: string | undefined,
): string | undefined {
  const pending = pendingDesigns.get(senderId)
  if (pending) return buildPresentationDesignDocument(pending)
  if (!deckPath || !deckPath.endsWith('.pptx')) return undefined
  const designPath = deckPath.replace(/\.pptx$/i, '.design.md')
  try {
    return existsSync(designPath) ? readFileSync(designPath, 'utf8') : undefined
  } catch {
    return undefined
  }
}

export function clearPresentationDesignSidecar(senderId: number): void {
  pendingDesigns.delete(senderId)
}
