import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { buildPresentationDesignDocument } from '@wiswork/agent-core'

const pendingDesigns = new Map<number, string>()

export function savePresentationDesignSidecar(
  senderId: number,
  deckPath: string | undefined,
  design: string,
): boolean {
  pendingDesigns.set(senderId, design)
  if (!deckPath || !/\.pptx$/i.test(deckPath)) return true
  return flushPresentationDesignSidecar(senderId, deckPath)
}

export function flushPresentationDesignSidecar(senderId: number, deckPath: string): boolean {
  const design = pendingDesigns.get(senderId)
  if (!design || !/\.pptx$/i.test(deckPath)) return false
  try {
    writeFileSync(deckPath.replace(/\.pptx$/i, '.design.md'), normalizeDesignDocument(design))
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
  if (pending) return normalizeDesignDocument(pending)
  if (!deckPath || !/\.pptx$/i.test(deckPath)) return undefined
  const designPath = deckPath.replace(/\.pptx$/i, '.design.md')
  try {
    return existsSync(designPath) ? readFileSync(designPath, 'utf8') : undefined
  } catch {
    return undefined
  }
}

function normalizeDesignDocument(value: string): string {
  return value.trimStart().startsWith('# DESIGN.md')
    ? value.trim()
    : buildPresentationDesignDocument(value)
}

export function clearPresentationDesignSidecar(senderId: number): void {
  pendingDesigns.delete(senderId)
}
