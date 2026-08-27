import { useState, type ReactNode } from 'react'
import type { LatexEditorCommand } from '../editor/editor-commands.js'
import { useLatexLocale } from '../i18n/locale.js'

type ToolbarEditorCommand = LatexEditorCommand | 'undo' | 'redo'
export type LatexRibbonTab = 'home' | 'insert' | 'compile' | 'pdf' | 'view'

interface WorkbenchToolbarProps {
  activePath: string | null
  mainFile?: string | null
  dirty: boolean
  disabled: boolean
  compileDisabled?: boolean
  compiling: boolean
  diagnosticCount?: number
  compilePanel?: ReactNode
  filesOpen: boolean
  previewOpen: boolean
  aiOpen: boolean
  pdfAvailable?: boolean
  pdfStale?: boolean
  exportingPdf?: boolean
  initialTab?: LatexRibbonTab
  onSave: () => void
  onCompile: () => void
  onCancelCompile?: () => void
  onExportPdf?: () => void
  onEditorCommand?: (command: ToolbarEditorCommand) => void
  onToggleFiles: () => void
  onTogglePreview: () => void
  onToggleAi: () => void
}

function ToolbarIcon({
  name,
}: {
  name: 'compile' | 'compile-running' | 'files' | 'preview' | 'ai' | 'export'
}) {
  if (name === 'compile') {
    return (
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <path d="m9 6 8 6-8 6V6Z" />
      </svg>
    )
  }
  if (name === 'compile-running') {
    return (
      <svg className="latex-compile-spinner" viewBox="0 0 24 24" aria-hidden="true">
        <path d="M19 12a7 7 0 1 1-2.05-4.95" />
        <path d="m16.5 3.75.45 3.3 3.3-.45" />
        <rect x="9.5" y="9.5" width="5" height="5" rx="0.7" />
      </svg>
    )
  }
  if (name === 'files') {
    return (
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <path d="M3.75 5.25h6l2 2h8.5v11.5H3.75V5.25Z" />
      </svg>
    )
  }
  if (name === 'preview') {
    return (
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <path d="M6 3.75h8l4 4v12.5H6V3.75Z" />
        <path d="M14 3.75v4h4M8.75 13h6.5M8.75 16h6.5" />
      </svg>
    )
  }
  if (name === 'export') {
    return (
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <path d="M12 3.75v11M8 11l4 4 4-4M5 19.5h14" />
      </svg>
    )
  }
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M12 3.75 13.7 9l5.3 1.7-5.3 1.7L12 17.75l-1.7-5.35L5 10.7 10.3 9 12 3.75Z" />
      <path d="m18.25 15 .7 2.05L21 17.75l-2.05.7-.7 2.05-.7-2.05-2.05-.7 2.05-.7.7-2.05Z" />
    </svg>
  )
}

function QuickAccessIcon({ name }: { name: 'save' | 'undo' | 'redo' }) {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      {name === 'save' ? (
        <>
          <path d="M3 4.5C3 3.67158 3.67158 3 4.5 3H17.1407L21 6.60325V19.5C21 20.3285 20.3285 21 19.5 21H4.5C3.67158 21 3 20.3285 3 19.5V4.5Z" />
          <path d="M12.0042 3L12 6.6923C12 6.86225 11.7761 7 11.5 7H7.5C7.22385 7 7 6.86225 7 6.6923V3H12.0042Z" />
          <path d="M7 13H17" />
          <path d="M7 17H12.0042" />
        </>
      ) : name === 'undo' ? (
        <>
          <path d="M5.91026 4L2.5 7.14791L5.91026 10.8205" />
          <path d="M3.96154 7.41028H15.1636C18.5169 7.41028 21.3646 10.1484 21.4953 13.5C21.6334 17.0416 18.707 20.0769 15.1636 20.0769H6.88384" />
        </>
      ) : (
        <>
          <path d="M18.0897 4L21.5 7.14791L18.0897 10.8205" />
          <path d="M20.0385 7.41028H8.83636C5.4831 7.41028 2.63537 10.1484 2.5047 13.5C2.36657 17.0416 5.29296 20.0769 8.83636 20.0769H17.1162" />
        </>
      )}
    </svg>
  )
}

function EditorTool({
  label,
  glyph,
  command,
  disabled,
  onCommand,
}: {
  label: string
  glyph: string
  command: ToolbarEditorCommand
  disabled: boolean
  onCommand: (command: ToolbarEditorCommand) => void
}) {
  return (
    <button
      type="button"
      className="latex-toolbar-button"
      title={label}
      aria-label={label}
      disabled={disabled}
      onClick={() => onCommand(command)}
    >
      <span className="latex-toolbar-icon-row latex-toolbar-glyph" aria-hidden="true">
        {glyph}
      </span>
      <span>{label}</span>
    </button>
  )
}

function Separator() {
  return <span className="latex-toolbar-separator" aria-hidden="true" />
}

export function WorkbenchToolbar({
  activePath,
  mainFile = null,
  dirty,
  disabled,
  compileDisabled = disabled,
  compiling,
  diagnosticCount = 0,
  compilePanel = null,
  filesOpen,
  previewOpen,
  aiOpen,
  pdfAvailable = false,
  pdfStale = false,
  exportingPdf = false,
  initialTab = 'home',
  onSave,
  onCompile,
  onCancelCompile = () => undefined,
  onExportPdf = () => undefined,
  onEditorCommand = () => undefined,
  onToggleFiles,
  onTogglePreview,
  onToggleAi,
}: WorkbenchToolbarProps) {
  const { t } = useLatexLocale()
  const [activeTab, setActiveTab] = useState<LatexRibbonTab>(initialTab)
  const [compileDetailsOpen, setCompileDetailsOpen] = useState(false)
  const tabs: Array<{ id: LatexRibbonTab; label: string }> = [
    { id: 'home', label: t('home') },
    { id: 'insert', label: t('insert') },
    { id: 'compile', label: t('compile') },
    { id: 'pdf', label: 'PDF' },
    { id: 'view', label: t('view') },
  ]

  const editorTool = (label: string, glyph: string, command: ToolbarEditorCommand) => (
    <EditorTool
      label={label}
      glyph={glyph}
      command={command}
      disabled={disabled}
      onCommand={onEditorCommand}
    />
  )

  return (
    <header className="latex-workbench-toolbar" role="toolbar" aria-label={t('toolbar')}>
      <div className="latex-toolbar-tabs">
        <button
          type="button"
          className="latex-toolbar-quick-button"
          title={t('save')}
          aria-label={t('save')}
          disabled={disabled || !activePath || !dirty}
          onClick={onSave}
        >
          <QuickAccessIcon name="save" />
        </button>
        <button
          type="button"
          className="latex-toolbar-quick-button"
          title={t('undo')}
          aria-label={t('undo')}
          disabled={disabled}
          onClick={() => onEditorCommand('undo')}
        >
          <QuickAccessIcon name="undo" />
        </button>
        <button
          type="button"
          className="latex-toolbar-quick-button"
          title={t('redo')}
          aria-label={t('redo')}
          disabled={disabled}
          onClick={() => onEditorCommand('redo')}
        >
          <QuickAccessIcon name="redo" />
        </button>
        <span className="latex-toolbar-quick-separator" aria-hidden="true" />
        <nav className="latex-ribbon-tabs" role="tablist" aria-label={t('toolbarTabs')}>
          {tabs.map((tab) => (
            <button
              key={tab.id}
              type="button"
              role="tab"
              aria-selected={activeTab === tab.id}
              aria-controls="latex-ribbon-band"
              className={`latex-ribbon-tab${activeTab === tab.id ? ' active' : ''}`}
              onClick={() => {
                setActiveTab(tab.id)
                setCompileDetailsOpen(false)
              }}
            >
              {tab.label}
            </button>
          ))}
        </nav>
        <span className="latex-toolbar-spacer" />
        {compiling && <span className="latex-toolbar-running">{t('compiling')}</span>}
        <span className="latex-toolbar-document" title={activePath ?? undefined}>
          {activePath ?? t('projectUnavailable')}
          {dirty && <span className="latex-toolbar-dirty" aria-label={t('unsavedChanges')} />}
        </span>
      </div>

      <div
        id="latex-ribbon-band"
        className="latex-toolbar-body"
        role="tabpanel"
        aria-label={tabs.find((tab) => tab.id === activeTab)?.label}
      >
        {activeTab === 'home' && (
          <>
            <div className="latex-toolbar-group">
              <button
                type="button"
                className="latex-toolbar-button ai-entry"
                title={aiOpen ? t('hideAi') : t('showAi')}
                aria-pressed={aiOpen}
                onClick={onToggleAi}
              >
                <span className="latex-toolbar-icon-row">
                  <ToolbarIcon name="ai" />
                </span>
                <span>WisWork AI</span>
              </button>
            </div>
            <Separator />
            <div className="latex-toolbar-group">
              {editorTool(t('bold'), 'B', 'bold')}
              {editorTool(t('italic'), 'I', 'italic')}
              {editorTool(t('underline'), 'U', 'underline')}
            </div>
            <Separator />
            <div className="latex-toolbar-group">
              {editorTool(t('section'), '§', 'section')}
              {editorTool(t('subsection'), '§§', 'subsection')}
              {editorTool(t('bulletedList'), '•', 'itemize')}
              {editorTool(t('numberedList'), '1.', 'enumerate')}
            </div>
          </>
        )}

        {activeTab === 'insert' && (
          <>
            <div className="latex-toolbar-group">
              {editorTool(t('inlineMath'), '$x$', 'inlineMath')}
              {editorTool(t('equation'), '∑', 'equation')}
            </div>
            <Separator />
            <div className="latex-toolbar-group">
              {editorTool(t('figure'), '▧', 'figure')}
              {editorTool(t('table'), '▦', 'table')}
            </div>
            <Separator />
            <div className="latex-toolbar-group">
              {editorTool(t('citation'), '[@]', 'cite')}
              {editorTool(t('reference'), '↗', 'ref')}
            </div>
          </>
        )}

        {activeTab === 'compile' && (
          <>
            <div className="latex-toolbar-group">
              <button
                type="button"
                className={`latex-toolbar-button${compiling ? ' busy' : ''}`}
                data-compile-state={compiling ? 'running' : 'idle'}
                title={compiling ? t('cancel') : t('compile')}
                aria-label={compiling ? t('cancel') : t('compile')}
                disabled={compileDisabled}
                onClick={compiling ? onCancelCompile : onCompile}
              >
                <span className="latex-toolbar-icon-row">
                  <ToolbarIcon name={compiling ? 'compile-running' : 'compile'} />
                </span>
                <span>{compiling ? t('cancel') : t('compile')}</span>
              </button>
              {compilePanel && (
                <button
                  type="button"
                  className="latex-toolbar-button"
                  title={t('problems')}
                  aria-label={`${t('problems')} (${diagnosticCount})`}
                  aria-expanded={compileDetailsOpen}
                  aria-controls="latex-compile-results"
                  onClick={() => setCompileDetailsOpen((open) => !open)}
                >
                  <span className="latex-toolbar-icon-row latex-toolbar-glyph" aria-hidden="true">
                    !
                  </span>
                  <span>
                    {t('problems')} ({diagnosticCount})
                  </span>
                </button>
              )}
            </div>
            <Separator />
            <div className="latex-toolbar-info-group">
              <span>{t('mainFile')}</span>
              <strong title={mainFile ?? undefined}>{mainFile ?? t('projectUnavailable')}</strong>
            </div>
          </>
        )}

        {activeTab === 'pdf' && (
          <div className="latex-toolbar-group">
            <button
              type="button"
              className="latex-toolbar-button"
              title={previewOpen ? t('hidePreview') : t('showPreview')}
              aria-pressed={previewOpen}
              onClick={onTogglePreview}
            >
              <span className="latex-toolbar-icon-row">
                <ToolbarIcon name="preview" />
              </span>
              <span>{t('pdfPreview')}</span>
            </button>
            <button
              type="button"
              className="latex-toolbar-button"
              title={pdfStale ? t('exportPdfStale') : t('exportPdf')}
              disabled={!pdfAvailable || exportingPdf}
              onClick={onExportPdf}
            >
              <span className="latex-toolbar-icon-row">
                <ToolbarIcon name="export" />
              </span>
              <span>{exportingPdf ? t('exportingPdf') : t('exportPdf')}</span>
            </button>
          </div>
        )}

        {activeTab === 'view' && (
          <div className="latex-toolbar-group">
            <button
              type="button"
              className="latex-toolbar-button"
              title={filesOpen ? t('hideFiles') : t('showFiles')}
              aria-pressed={filesOpen}
              onClick={onToggleFiles}
            >
              <span className="latex-toolbar-icon-row">
                <ToolbarIcon name="files" />
              </span>
              <span>{t('files')}</span>
            </button>
            <button
              type="button"
              className="latex-toolbar-button"
              title={previewOpen ? t('hidePreview') : t('showPreview')}
              aria-pressed={previewOpen}
              onClick={onTogglePreview}
            >
              <span className="latex-toolbar-icon-row">
                <ToolbarIcon name="preview" />
              </span>
              <span>{t('pdfPreview')}</span>
            </button>
            <button
              type="button"
              className="latex-toolbar-button ai-entry"
              title={aiOpen ? t('hideAi') : t('showAi')}
              aria-pressed={aiOpen}
              onClick={onToggleAi}
            >
              <span className="latex-toolbar-icon-row">
                <ToolbarIcon name="ai" />
              </span>
              <span>{t('aiPanel')}</span>
            </button>
          </div>
        )}
      </div>

      {compileDetailsOpen && compilePanel && (
        <div
          id="latex-compile-results"
          className="latex-compile-popover"
          role="region"
          aria-label={t('compileResults')}
        >
          {compilePanel}
        </div>
      )}
    </header>
  )
}
