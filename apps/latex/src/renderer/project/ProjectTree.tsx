import { useEffect, useRef, useState } from 'react'
import { useLatexLocale } from '../i18n/locale.js'
import { canRenameFile } from '../workbench-coordination.js'

export interface ProjectTreeProps {
  files: readonly string[]
  activePath: string | null
  onOpen: (path: string) => void
  onCreate: () => void
  onRename: (path: string) => void
  onDelete: (path: string) => void
  mainFile?: string | null
}

export function ProjectTree({
  files,
  activePath,
  onOpen,
  onCreate,
  onRename,
  onDelete,
  mainFile,
}: ProjectTreeProps) {
  const { t } = useLatexLocale()
  const [menuPath, setMenuPath] = useState<string | null>(null)
  const menuRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const close = (event: MouseEvent) => {
      if (!menuRef.current?.contains(event.target as Node)) setMenuPath(null)
    }
    document.addEventListener('mousedown', close)
    return () => document.removeEventListener('mousedown', close)
  }, [])
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
            <div className="file-actions" ref={menuPath === path ? menuRef : undefined}>
              <button
                type="button"
                className="file-actions-trigger"
                aria-label={`File actions for ${path}`}
                aria-expanded={menuPath === path}
                onClick={() => setMenuPath((current) => (current === path ? null : path))}
              >
                ⋯
              </button>
              {menuPath === path && (
                <div className="file-actions-menu" role="menu">
                  <button
                    type="button"
                    role="menuitem"
                    disabled={!canRenameFile(path, mainFile ?? null)}
                    onClick={() => {
                      setMenuPath(null)
                      onRename(path)
                    }}
                  >
                    {t('rename')}
                  </button>
                  <button
                    type="button"
                    role="menuitem"
                    disabled={!canRenameFile(path, mainFile ?? null)}
                    onClick={() => {
                      setMenuPath(null)
                      onDelete(path)
                    }}
                  >
                    Delete
                  </button>
                </div>
              )}
            </div>
          </li>
        ))}
      </ul>
    </aside>
  )
}
