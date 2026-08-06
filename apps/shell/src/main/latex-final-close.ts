import type { WebContents } from 'electron'
import {
  latexQueryDirty,
  releaseLatexEditFlush,
  requestLatexEditFlush,
} from '../../../latex/src/main/latex-main'

type LatexCloseTab = { webContents: WebContents }

export async function prepareLatexCloseTabs(
  tabs: ReadonlyArray<LatexCloseTab>,
  prepare: (contents: WebContents) => Promise<boolean> = requestLatexEditFlush,
): Promise<boolean> {
  const results = await Promise.all(tabs.map((tab) => prepare(tab.webContents).catch(() => false)))
  return results.every(Boolean)
}

export function releaseLatexCloseTabs(
  tabs: ReadonlyArray<LatexCloseTab>,
  release: (contents: WebContents) => unknown = releaseLatexEditFlush,
): void {
  for (const tab of tabs) release(tab.webContents)
}

export async function finalLatexCloseCheck(
  tabs: ReadonlyArray<LatexCloseTab>,
  dirty: (contents: WebContents) => Promise<boolean> = latexQueryDirty,
): Promise<boolean> {
  for (const tab of tabs) if (await dirty(tab.webContents)) return false
  return true
}
