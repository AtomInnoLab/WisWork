import { useState, type ReactNode } from 'react'
import type { LatexEditorCommand } from '../editor/editor-commands.js'
import { useLatexLocale } from '../i18n/locale.js'

type ToolbarEditorCommand = LatexEditorCommand | 'undo' | 'redo'

interface WorkbenchToolbarProps {
  activePath: string | null
  dirty: boolean
  disabled: boolean
  compileDisabled?: boolean
  compiling: boolean
  diagnosticCount?: number
  compilePanel?: ReactNode
  filesOpen: boolean
  previewOpen: boolean
  aiOpen: boolean
  onSave: () => void
  onCompile: () => void
  onCancelCompile?: () => void
  onEditorCommand?: (command: ToolbarEditorCommand) => void
  onToggleFiles: () => void
  onTogglePreview: () => void
  onToggleAi: () => void
}

function ToolbarIcon({ name }: { name: 'save' | 'compile' | 'files' | 'preview' | 'ai' }) {
  if (name === 'save') {
    return (
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <path d="M5 3.75h11.5L20.25 7.5v12.75H3.75V3.75H5Z" />
        <path d="M7.5 3.75v5h8v-5M7 20.25v-7h10v7" />
      </svg>
    )
  }
  if (name === 'compile') {
    return (
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <path d="m9 6 8 6-8 6V6Z" />
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
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M12 3.75 13.7 9l5.3 1.7-5.3 1.7L12 17.75l-1.7-5.35L5 10.7 10.3 9 12 3.75Z" />
      <path d="m18.25 15 .7 2.05L21 17.75l-2.05.7-.7 2.05-.7-2.05-2.05-.7 2.05-.7.7-2.05Z" />
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

export function WorkbenchToolbar({
  activePath,
  dirty,
  disabled,
  compileDisabled = disabled,
  compiling,
  diagnosticCount = 0,
  compilePanel = null,
  filesOpen,
  previewOpen,
  aiOpen,
  onSave,
  onCompile,
  onCancelCompile = () => undefined,
  onEditorCommand = () => undefined,
  onToggleFiles,
  onTogglePreview,
  onToggleAi,
}: WorkbenchToolbarProps) {
  const { t } = useLatexLocale()
  const [compileDetailsOpen, setCompileDetailsOpen] = useState(false)
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
          <ToolbarIcon name="save" />
        </button>
        <span className="latex-toolbar-product">LaTeX</span>
        <span className="latex-toolbar-spacer" />
        <span className="latex-toolbar-document" title={activePath ?? undefined}>
          {activePath ?? t('projectUnavailable')}
          {dirty && <span className="latex-toolbar-dirty" aria-label={t('unsavedChanges')} />}
        </span>
      </div>
      <div className="latex-toolbar-body">
        <div className="latex-toolbar-group">
          <EditorTool
            label={t('undo')}
            glyph="↶"
            command="undo"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
          <EditorTool
            label={t('redo')}
            glyph="↷"
            command="redo"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
        </div>
        <span className="latex-toolbar-separator" aria-hidden="true" />
        <div className="latex-toolbar-group">
          <EditorTool
            label={t('bold')}
            glyph="B"
            command="bold"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
          <EditorTool
            label={t('italic')}
            glyph="I"
            command="italic"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
          <EditorTool
            label={t('underline')}
            glyph="U"
            command="underline"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
        </div>
        <span className="latex-toolbar-separator" aria-hidden="true" />
        <div className="latex-toolbar-group">
          <EditorTool
            label={t('section')}
            glyph="§"
            command="section"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
          <EditorTool
            label={t('subsection')}
            glyph="§§"
            command="subsection"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
          <EditorTool
            label={t('bulletedList')}
            glyph="•"
            command="itemize"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
          <EditorTool
            label={t('numberedList')}
            glyph="1."
            command="enumerate"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
        </div>
        <span className="latex-toolbar-separator" aria-hidden="true" />
        <div className="latex-toolbar-group">
          <EditorTool
            label={t('inlineMath')}
            glyph="$x$"
            command="inlineMath"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
          <EditorTool
            label={t('equation')}
            glyph="∑"
            command="equation"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
          <EditorTool
            label={t('citation')}
            glyph="[@]"
            command="cite"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
          <EditorTool
            label={t('reference')}
            glyph="↗"
            command="ref"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
        </div>
        <span className="latex-toolbar-separator" aria-hidden="true" />
        <div className="latex-toolbar-group">
          <EditorTool
            label={t('figure')}
            glyph="▧"
            command="figure"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
          <EditorTool
            label={t('table')}
            glyph="▦"
            command="table"
            disabled={disabled}
            onCommand={onEditorCommand}
          />
        </div>
        <span className="latex-toolbar-separator" aria-hidden="true" />
        <div className="latex-toolbar-group">
          <button
            type="button"
            className={`latex-toolbar-button${compiling ? ' busy' : ''}`}
            title={compiling ? t('cancel') : t('compile')}
            aria-label={compiling ? t('cancel') : t('compile')}
            disabled={compileDisabled}
            onClick={compiling ? onCancelCompile : onCompile}
          >
            <span className="latex-toolbar-icon-row">
              <ToolbarIcon name="compile" />
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
        <span className="latex-toolbar-separator" aria-hidden="true" />
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
            <span>PDF</span>
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
            <span>WisWork AI</span>
          </button>
        </div>
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
