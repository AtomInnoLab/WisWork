import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { buildPresentationDesignDocument } from '@wiswork/agent-core'

const pendingDesigns = new Map<number, string>()
const activeDesignPaths = new Map<number, string>()

export function savePresentationDesignSidecar(
  senderId: number,
  deckPath: string | undefined,
  design: string,
): boolean {
  let normalized: string
  try {
    normalized = normalizeDesignDocument(design)
  } catch {
    return false
  }
  pendingDesigns.set(senderId, normalized)
  if (!deckPath || !/\.pptx$/i.test(deckPath)) {
    const activePath = activeDesignPaths.get(senderId)
    if (!activePath) return true
    try {
      writeFileSync(activePath, normalized)
      pendingDesigns.delete(senderId)
      return true
    } catch {
      return false
    }
  }
  return flushPresentationDesignSidecar(senderId, deckPath)
}

export function flushPresentationDesignSidecar(senderId: number, deckPath: string): boolean {
  const design = pendingDesigns.get(senderId) ?? readActivePresentationDesignSidecar(senderId)
  if (!design || !/\.pptx$/i.test(deckPath)) return false
  try {
    const designPath = deckPath.replace(/\.pptx$/i, '.design.md')
    writeFileSync(designPath, normalizeDesignDocument(design))
    pendingDesigns.delete(senderId)
    activeDesignPaths.set(senderId, designPath)
    return true
  } catch {
    return false
  }
}

export function readPresentationDesignSidecar(
  senderId: number,
  deckPath: string | undefined,
): string | undefined {
  const active = readActivePresentationDesignSidecar(senderId)
  if (active) return active
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

export function materializePresentationDesignSidecar(
  senderId: number,
  deckPath: string | undefined,
  fallbackPath: string,
): string | undefined {
  const design = readPresentationDesignSidecar(senderId, deckPath)
  if (!design) return undefined
  const designPath =
    deckPath && /\.pptx$/i.test(deckPath)
      ? deckPath.replace(/\.pptx$/i, '.design.md')
      : fallbackPath
  try {
    writeFileSync(designPath, normalizeDesignDocument(design))
    activeDesignPaths.set(senderId, designPath)
    pendingDesigns.delete(senderId)
    return designPath
  } catch {
    return undefined
  }
}

function readActivePresentationDesignSidecar(senderId: number): string | undefined {
  const path = activeDesignPaths.get(senderId)
  if (!path) return undefined
  try {
    return existsSync(path) ? normalizeDesignDocument(readFileSync(path, 'utf8')) : undefined
  } catch {
    return undefined
  }
}

function normalizeDesignDocument(value: string): string {
  if (!value.trimStart().startsWith('# DESIGN.md')) return buildPresentationDesignDocument(value)
  const normalized = value.trim()
  if (!normalized.replace(/^#\s*DESIGN\.md\s*/i, '').trim())
    throw new Error('empty_presentation_design')
  return normalized
}

export function clearPresentationDesignSidecar(senderId: number): void {
  pendingDesigns.delete(senderId)
  activeDesignPaths.delete(senderId)
}
