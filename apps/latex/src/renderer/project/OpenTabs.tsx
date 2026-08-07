export interface OpenTabsProps {
  paths: readonly string[]
  activePath: string | null
  dirty: ReadonlySet<string>
  onActivate: (path: string) => void
  onClose: (path: string) => void
}

export function OpenTabs({ paths, activePath, dirty, onActivate, onClose }: OpenTabsProps) {
  return (
    <nav className="open-tabs" aria-label="Open files">
      {paths.map((path) => (
        <div key={path} className={path === activePath ? 'active' : undefined}>
          <button type="button" onClick={() => onActivate(path)}>
            {path}
            {dirty.has(path) ? ' •' : ''}
          </button>
          <button type="button" aria-label={`Close ${path}`} onClick={() => onClose(path)}>
            ×
          </button>
        </div>
      ))}
    </nav>
  )
}
