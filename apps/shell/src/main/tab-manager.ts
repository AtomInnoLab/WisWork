import { realpathSync, statSync } from 'node:fs'
import { basename } from 'node:path'
import { BrowserWindow } from 'electron'
import type { Rectangle, WebContents, WebContentsView } from 'electron'

import {
  createDocsView,
  docsQueryDirty,
  markDocsNewBlank,
  requestDocsClose,
  setActiveDocsResolver,
  teardownDocsRenderer,
} from '../../../docs/src/main/docs-main'
import {
  createLatexView,
  latexQueryDirty,
  requestLatexEditFlush,
  releaseLatexEditFlush,
  requestLatexClose,
  teardownLatexRenderer,
} from '../../../latex/src/main/latex-main'
import { createPdfView, pdfIsDirty, requestPdfClose } from '../../../pdf/src/main/pdf-main'
import {
  createSheetsView,
  requestSheetsClose,
  setActiveSheetsWebContents,
  setSheetsNewBlank,
  sheetsPendingEditCount,
} from '../../../sheets/src/main/sheets-main'
import {
  createSlidesView,
  requestSlidesClose,
  setActiveSlidesWebContents,
  slidesIsDirty,
} from '../../../slides/src/main/slides-main'
import type { TabKind, TabSummary } from '../shared/tabs-api'

interface TabRecord {
  id: string
  kind: TabKind
  /** null for the Home tab — it's rendered by the shell window's own webContents */
  view: WebContentsView | null
  title: string
  filePath?: string
}

export interface DocumentCloseContext {
  readonly documentId: string
  readonly kind: TabKind
  readonly webContents: WebContents
}

export interface DocumentClosePreparation {
  commit(): Promise<void>
  rollback(): void
  finalize(): void
}

/** must match the tab strip's rendered height (apps/shell/src/renderer/src/TabBar.tsx) */
const TAB_STRIP_HEIGHT = 40
const HOME_ID = 'home'

/**
 * Owns every open tab (Home + docs + sheets) inside the shell's single
 * BrowserWindow. Docs/sheets tabs are WebContentsView children of that
 * window; only the active one is visible at a time. Home has no view of its
 * own — hiding every other tab reveals the shell window's own content.
 */
export class TabManager {
  private readonly tabs: TabRecord[] = [{ id: HOME_ID, kind: 'home', view: null, title: 'WisWork' }]
  private activeId: string = HOME_ID
  private nextId = 1
  /** tab whose page entered HTML fullscreen (e.g. slides slideshow) — its view covers the tab strip */
  private htmlFullScreenId: string | null = null
  /** tabs mid unsaved-changes prompt, so a second close click doesn't stack dialogs */
  private readonly closingIds = new Set<string>()

  constructor(
    private readonly shellWindow: BrowserWindow,
    private readonly onChanged: () => void,
    private readonly applyMenuFor: (kind: TabKind) => void,
    /** localized placeholder title for a tab that has no file yet */
    private readonly untitledTitleFor?: (kind: TabKind) => string,
    /** Two-phase automation close: quiesce before guards, then commit or resume on cancellation. */
    private readonly prepareDocumentClose?: (
      context: DocumentCloseContext,
    ) => Promise<DocumentClosePreparation>,
  ) {
    shellWindow.on('resize', () => this.layout())
  }

  private untitled(kind: TabKind, fallback: string): string {
    return this.untitledTitleFor?.(kind) ?? fallback
  }

  private contentBounds(): Rectangle {
    const { width, height } = this.shellWindow.getContentBounds()
    if (this.htmlFullScreenId !== null && this.htmlFullScreenId === this.activeId) {
      return { x: 0, y: 0, width, height }
    }
    return { x: 0, y: TAB_STRIP_HEIGHT, width, height: Math.max(0, height - TAB_STRIP_HEIGHT) }
  }

  /**
   * When a tab's page enters HTML fullscreen (the slides slideshow calls requestFullscreen),
   * grow its view over the tab strip so nothing of the shell chrome shows;
   * restore the normal bounds on leave.
   */
  private trackHtmlFullScreen(id: string, view: WebContentsView): void {
    view.webContents.on('enter-html-full-screen', () => {
      this.htmlFullScreenId = id
      this.layout()
    })
    view.webContents.on('leave-html-full-screen', () => {
      if (this.htmlFullScreenId === id) this.htmlFullScreenId = null
      this.layout()
    })
  }

  /** re-fit the active tab's view after a window resize */
  layout(): void {
    const active = this.tabs.find((t) => t.id === this.activeId)
    if (active?.view) active.view.setBounds(this.contentBounds())
  }

  ownsWebContents(senderId: number): boolean {
    return this.tabs.some((tab) => tab.view?.webContents.id === senderId)
  }

  documentIdForWebContents(senderId: number): string | null {
    return this.tabs.find((tab) => tab.view?.webContents.id === senderId)?.id ?? null
  }

  ownsDocument(senderId: number, documentId: string): boolean {
    return this.documentIdForWebContents(senderId) === documentId
  }

  list(): TabSummary[] {
    return this.tabs.map((t) => ({
      id: t.id,
      kind: t.kind,
      title: t.title,
      closable: t.id !== HOME_ID,
      active: t.id === this.activeId,
    }))
  }

  openHomeTab(): void {
    this.activateTab(HOME_ID)
  }

  openDocsTab(openPath?: string, options?: { newBlank?: boolean }): string {
    const view = createDocsView(openPath)
    const id = `t${this.nextId++}`
    if (options?.newBlank) markDocsNewBlank(view.webContents.id)
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.tabs.push({
      id,
      kind: 'docs',
      view,
      title: openPath ? basename(openPath) : this.untitled('docs', 'WisWork Docs'),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  openSheetsTab(openPath?: string, options?: { newBlank?: boolean }): string {
    if (options?.newBlank) setSheetsNewBlank()
    const view = createSheetsView({ includeAiHandlers: false })
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.tabs.push({
      id,
      kind: 'sheets',
      view,
      title: openPath ? basename(openPath) : this.untitled('sheets', 'AI Sheets'),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  openSlidesTab(openPath?: string): string {
    const view = createSlidesView(openPath)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.tabs.push({
      id,
      kind: 'slides',
      view,
      title: openPath ? basename(openPath) : this.untitled('slides', 'AI Slides'),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  openLatexTab(projectPath: string): string {
    const canonicalPath = realpathSync(projectPath)
    if (!statSync(canonicalPath).isDirectory()) throw new Error('LaTeX project must be a directory')
    const existing = this.tabs.find((tab) => tab.kind === 'latex' && tab.filePath === canonicalPath)
    if (existing) {
      this.activateTab(existing.id)
      return existing.id
    }
    const view = createLatexView(canonicalPath)
    const id = 't' + this.nextId++
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.tabs.push({
      id,
      kind: 'latex',
      view,
      title: basename(canonicalPath),
      filePath: canonicalPath,
    })
    this.activateTab(id)
    return id
  }

  openPdfTab(openPath: string): string {
    const view = createPdfView(openPath)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.tabs.push({ id, kind: 'pdf', view, title: basename(openPath), filePath: openPath })
    this.activateTab(id)
    return id
  }

  activateTab(id: string): void {
    const target = this.tabs.find((t) => t.id === id)
    if (!target) return
    for (const t of this.tabs) t.view?.setVisible(t.id === id)
    if (target.view) target.view.setBounds(this.contentBounds())
    this.activeId = id
    setActiveDocsResolver(target.kind === 'docs' ? () => target.view!.webContents : () => null)
    if (target.kind === 'sheets' && target.view) setActiveSheetsWebContents(target.view.webContents)
    if (target.kind === 'slides' && target.view) setActiveSlidesWebContents(target.view.webContents)
    this.applyMenuFor(target.kind)
    this.onChanged()
  }

  /** move a tab to a new index in the strip; Home is pinned at index 0 */
  reorderTab(id: string, toIndex: number): void {
    if (id === HOME_ID) return
    const fromIndex = this.tabs.findIndex((t) => t.id === id)
    if (fromIndex < 0) return
    const clamped = Math.min(Math.max(Math.trunc(toIndex), 1), this.tabs.length - 1)
    if (clamped === fromIndex) return
    const [moved] = this.tabs.splice(fromIndex, 1)
    this.tabs.splice(clamped, 0, moved)
    this.onChanged()
  }

  /** a module opened a file inside an existing tab (⌘O / queued path) — sync title + dedupe path */
  setTabFileFor(webContentsId: number, filePath: string): void {
    const tab = this.tabs.find((t) => t.view?.webContents.id === webContentsId)
    if (!tab) return
    tab.filePath = filePath
    tab.title = basename(filePath)
    this.onChanged()
  }

  /** a file was renamed on disk (rename from the Home list) — sync any open tab's title/path;
   *  returns the affected views so callers can notify the embedded editors */
  renameTabFile(
    oldPath: string,
    newPath: string,
  ): Array<{ kind: TabKind; webContents: WebContents }> {
    const affected: Array<{ kind: TabKind; webContents: WebContents }> = []
    for (const tab of this.tabs) {
      if (tab.filePath !== oldPath) continue
      tab.filePath = newPath
      tab.title = basename(newPath)
      if (tab.view) affected.push({ kind: tab.kind, webContents: tab.view.webContents })
    }
    if (affected.length > 0) this.onChanged()
    return affected
  }

  /** sheets tabs whose renderer reports unsaved journal edits (shell-close guard) */
  dirtySheetsTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter(
        (t) => t.kind === 'sheets' && t.view && sheetsPendingEditCount(t.view.webContents.id) > 0,
      )
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** pdf tabs whose renderer reports unsaved markups/form edits (shell-close guard) */
  dirtyPdfTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'pdf' && t.view && pdfIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** slides tabs whose main-process session has unsaved edits (shell-close guard) */
  dirtySlidesTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'slides' && t.view && slidesIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  latexTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((tab) => tab.kind === 'latex' && tab.view)
      .map((tab) => ({ id: tab.id, webContents: tab.view!.webContents }))
  }

  latexSessionState(): { projectPaths: string[]; activeProjectPath: string | null } {
    const latexTabs = this.tabs.filter((tab) => tab.kind === 'latex' && tab.filePath)
    const active = latexTabs.find((tab) => tab.id === this.activeId)
    return {
      projectPaths: latexTabs.map((tab) => tab.filePath!),
      activeProjectPath: active?.filePath ?? null,
    }
  }

  /** all live docs tabs — dirtiness lives renderer-side, caller queries async (shell-close guard) */
  docsTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'docs' && t.view)
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** closes whichever tab is currently active; no-op for Home (Cmd+W target) */
  closeActiveTab(): void {
    void this.closeTab(this.activeId)
  }

  async closeTab(id: string): Promise<void> {
    if (id === HOME_ID) return
    const tab = this.tabs.find((t) => t.id === id)
    if (!tab || this.closingIds.has(id)) return
    this.closingIds.add(id)
    let preparation: DocumentClosePreparation | undefined
    let preparationCommitted = false
    let identityRemoved = false
    let latexPrepared = false
    try {
      if (tab.view && this.prepareDocumentClose) {
        preparation = await this.prepareDocumentClose({
          documentId: id,
          kind: tab.kind,
          webContents: tab.view.webContents,
        })
      }
      let closeGuard =
        tab.view &&
        (tab.kind === 'sheets' && sheetsPendingEditCount(tab.view.webContents.id) > 0
          ? requestSheetsClose
          : tab.kind === 'pdf' && pdfIsDirty(tab.view.webContents.id)
            ? requestPdfClose
            : tab.kind === 'slides' && slidesIsDirty(tab.view.webContents.id)
              ? requestSlidesClose
              : null)
      // docs dirty state lives in the renderer and needs an async query; skip the guard when clean.
      if (!closeGuard && tab.kind === 'latex' && tab.view) {
        if (!(await requestLatexEditFlush(tab.view.webContents))) return
        latexPrepared = true
        if (await latexQueryDirty(tab.view.webContents)) closeGuard = requestLatexClose
      }
      if (!closeGuard && tab.kind === 'docs' && tab.view) {
        if (await docsQueryDirty(tab.view.webContents)) closeGuard = requestDocsClose
      }
      if (closeGuard && tab.view) {
        // Bring the tab into view so the save prompt has visible context.
        if (this.activeId !== id) this.activateTab(id)
        if (!(await closeGuard(tab.view.webContents, this.shellWindow))) {
          if (latexPrepared) releaseLatexEditFlush(tab.view.webContents)
          return
        }
      }
      await preparation?.commit()
      preparationCommitted = true
      const idx = this.tabs.findIndex((t) => t.id === id)
      if (idx < 0) return
      if (this.htmlFullScreenId === id) this.htmlFullScreenId = null
      const [removed] = this.tabs.splice(idx, 1)
      identityRemoved = true
      if (this.activeId === id) {
        const fallback = this.tabs[idx - 1] ?? this.tabs[0]
        this.activateTab(fallback.id)
      } else {
        this.onChanged()
      }
      if (removed.view) {
        removed.view.setVisible(false)
        this.shellWindow.contentView.removeChildView(removed.view)
        if (removed.kind === 'docs' || removed.kind === 'latex') {
          // WebContentsView teardown can wedge these renderers; detaching lets app quit reclaim them.
          if (removed.kind === 'docs') teardownDocsRenderer(removed.view.webContents)
          else teardownLatexRenderer(removed.view.webContents)
        } else {
          removed.view.webContents.close()
        }
      }
    } catch {
      if (latexPrepared && tab.view) releaseLatexEditFlush(tab.view.webContents)
    } finally {
      if (preparationCommitted && identityRemoved) preparation?.finalize()
      else if (!preparationCommitted) preparation?.rollback()
      this.closingIds.delete(id)
    }
  }

  findDocsTabByPath(path: string): string | undefined {
    return this.tabs.find((t) => t.kind === 'docs' && t.filePath === path)?.id
  }

  findSheetsTab(): string | undefined {
    return this.tabs.find((t) => t.kind === 'sheets')?.id
  }

  findSheetsTabByPath(path: string): string | undefined {
    return this.tabs.find((t) => t.kind === 'sheets' && t.filePath === path)?.id
  }

  findSlidesTabByPath(path: string): string | undefined {
    return this.tabs.find((t) => t.kind === 'slides' && t.filePath === path)?.id
  }

  findLatexTabByPath(path: string): string | undefined {
    let canonical: string
    try {
      canonical = realpathSync(path)
    } catch {
      return undefined
    }
    return this.tabs.find((tab) => tab.kind === 'latex' && tab.filePath === canonical)?.id
  }

  findPdfTabByPath(path: string): string | undefined {
    return this.tabs.find((t) => t.kind === 'pdf' && t.filePath === path)?.id
  }

  /** the active tab's pdf view, if the active tab is a pdf (pdf menu target) */
  activePdfTab(): { id: string; webContents: WebContents; filePath?: string } | undefined {
    const tab = this.tabs.find((t) => t.id === this.activeId)
    return tab?.kind === 'pdf' && tab.view
      ? { id: tab.id, webContents: tab.view.webContents, filePath: tab.filePath }
      : undefined
  }
}
