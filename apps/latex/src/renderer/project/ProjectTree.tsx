import { useLatexLocale } from '../i18n/locale.js'
import { canRenameFile } from '../workbench-coordination.js'

export interface ProjectTreeProps {
  files: readonly string[]
  activePath: string | null
  onOpen: (path: string) => void
  onCreate: () => void
  onRename: (path: string) => void
  mainFile?: string | null
}

export function ProjectTree({
  files,
  activePath,
  onOpen,
  onCreate,
  onRename,
  mainFile,
}: ProjectTreeProps) {
  const { t } = useLatexLocale()
  return (
    <aside className="project-tree">
      <header>
        <span>{t('files')}</span>
        <button type="button" onClick={onCreate} title={t('newFile')}>
          +
        </button>
      </header>
      <ul>
        {files.map((path) => (
          <li key={path} className={path === activePath ? 'active' : undefined}>
            <button type="button" onClick={() => onOpen(path)}>
              {path}
            </button>
            <button
              type="button"
              className="rename-file"
              disabled={!canRenameFile(path, mainFile ?? null)}
              title={
                !canRenameFile(path, mainFile ?? null)
                  ? 'The main file cannot be renamed'
                  : undefined
              }
              onClick={() => onRename(path)}
            >
              ⋯
            </button>
          </li>
        ))}
      </ul>
    </aside>
  )
}
