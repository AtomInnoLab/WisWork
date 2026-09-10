import { useEffect, useRef, useState, type ReactNode } from 'react'
import { Markdown } from '@wiswork/ui'
export interface OfficeDesignDocument {
  documentId: string
  markdown: string
  revision: string
}
export type OfficeDesignRequest = (
  body:
    | { action: 'open'; documentId: string; markdown: string }
    | { action: 'read'; documentId: string },
  signal?: AbortSignal,
) => Promise<OfficeDesignDocument>
export interface OfficeDesignPanelProps {
  current?: { markdown: string; sourceId: string }
  selection?: { markdown: string; editable: boolean }
  busy: boolean
  request?: OfficeDesignRequest
  onApply(markdown: string): void
}
export function OfficeDesignPanel({
  current,
  selection,
  busy,
  request,
  onApply,
}: OfficeDesignPanelProps): ReactNode {
  const [reader, setReader] = useState<{ historical?: string }>()
  const [remote, setRemote] = useState<OfficeDesignDocument & { baseMarkdown: string }>()
  const [opening, setOpening] = useState(false)
  const [error, setError] = useState('')
  const openController = useRef<AbortController | undefined>(undefined)
  const openedSelection = useRef<OfficeDesignPanelProps['selection']>(undefined)
  const dialog = useRef<HTMLElement>(null)
  const opener = useRef<HTMLElement | null>(null)
  const currentMarkdown = current?.markdown
  const remoteDocumentId = remote?.documentId
  const conflict = Boolean(remote && remote.baseMarkdown !== currentMarkdown)
  const changed = Boolean(remote && remote.markdown !== remote.baseMarkdown)
  const historical = reader?.historical !== undefined
  const readingCurrent = Boolean(reader && !historical)
  const markdown = reader?.historical ?? remote?.markdown ?? currentMarkdown ?? ''

  useEffect(() => {
    if (!selection || selection === openedSelection.current) return
    openedSelection.current = selection
    setReader(
      selection.editable && selection.markdown === currentMarkdown
        ? {}
        : { historical: selection.markdown },
    )
  }, [selection, currentMarkdown])
  useEffect(() => () => openController.current?.abort(), [])
  useEffect(() => {
    if (!reader) return
    opener.current = document.activeElement as HTMLElement
    dialog.current?.focus()
    return () => opener.current?.focus()
  }, [reader])
  useEffect(() => {
    if (!remoteDocumentId || !request || busy || !currentMarkdown || !readingCurrent) return
    const documentId = remoteDocumentId
    const controller = new AbortController()
    let active = true
    let timer: ReturnType<typeof setTimeout>
    const poll = async () => {
      try {
        const saved = await request({ action: 'read', documentId }, controller.signal)
        if (!active) return
        setRemote((value) => (value?.documentId === documentId ? { ...value, ...saved } : value))
        setError('')
      } catch {
        if (active) setError('暂时无法同步 PC 文件，连接恢复后会重试。')
      }
      if (active) timer = setTimeout(() => void poll(), 5000)
    }
    timer = setTimeout(() => void poll(), 5000)
    return () => {
      active = false
      clearTimeout(timer)
      controller.abort()
    }
  }, [remoteDocumentId, request, busy, currentMarkdown, readingCurrent])

  async function openOnPc() {
    if (!request || !currentMarkdown || busy || opening || historical) return
    const controller = new AbortController()
    openController.current?.abort()
    openController.current = controller
    setOpening(true)
    setError('')
    const documentId =
      !conflict && remote ? remote.documentId : `design_${crypto.randomUUID().replaceAll('-', '')}`
    try {
      const opened = await request(
        { action: 'open', documentId, markdown: currentMarkdown },
        controller.signal,
      )
      if (!controller.signal.aborted) setRemote({ ...opened, baseMarkdown: currentMarkdown })
    } catch {
      if (!controller.signal.aborted)
        setError('无法在已连接的 WisWork PC 打开，请确认 PC 已更新且连接正常。')
    } finally {
      if (!controller.signal.aborted) setOpening(false)
    }
  }
  if (!current && !reader) return null
  return (
    <>
      {current && (
        <section className="design-document-entry" aria-label="当前设计文档">
          <button type="button" onClick={() => setReader({})}>
            <strong>DESIGN.md</strong>
            <span>{changed ? 'PC 修改待应用' : '查看设计方案'}</span>
          </button>
        </section>
      )}
      {reader && (
        <div className="design-dialog-backdrop">
          <section
            ref={dialog}
            tabIndex={-1}
            className="design-dialog"
            role="dialog"
            aria-modal="true"
            aria-label="DESIGN.md"
            onKeyDown={(event) => {
              if (event.key === 'Escape') setReader(undefined)
              if (event.key === 'Tab') {
                const buttons = Array.from(
                  event.currentTarget.querySelectorAll<HTMLElement>(
                    'button:not(:disabled), a[href]',
                  ),
                )
                const first = buttons[0],
                  last = buttons.at(-1)
                if (
                  event.shiftKey &&
                  (document.activeElement === first ||
                    document.activeElement === event.currentTarget)
                ) {
                  event.preventDefault()
                  last?.focus()
                } else if (!event.shiftKey && document.activeElement === last) {
                  event.preventDefault()
                  first?.focus()
                }
              }
            }}
          >
            <header>
              <strong>DESIGN.md</strong>
              <button
                type="button"
                className="quiet"
                aria-label="关闭设计文档"
                onClick={() => setReader(undefined)}
              >
                ×
              </button>
            </header>
            <p role="status">
              {historical
                ? '历史设计快照，仅供查看。'
                : conflict
                  ? '设计已更新。PC 草稿已保留，请打开当前版本重新编辑，避免覆盖新方案。'
                  : changed
                    ? '已从 PC 同步，应用后生成新的设计合同。'
                    : remote
                      ? '已在 WisWork PC 打开，保存后会自动同步到这里。'
                      : '在此阅读设计方案，在已连接的 WisWork PC 中编辑。'}
            </p>
            <div className="design-document-content">
              <Markdown text={markdown} />
            </div>
            {error && <p role="alert">{error}</p>}
            <footer>
              {!historical && (
                <>
                  <button
                    type="button"
                    className="secondary"
                    disabled={!request || busy || opening}
                    onClick={() => void openOnPc()}
                  >
                    {opening
                      ? '正在打开…'
                      : conflict
                        ? '打开当前版本重新编辑'
                        : '在 WisWork PC 编辑'}
                  </button>
                  {changed && (
                    <button
                      type="button"
                      disabled={busy || conflict}
                      onClick={() => {
                        if (!remote || conflict || busy) return
                        onApply(remote.markdown)
                        setRemote(undefined)
                        setReader(undefined)
                      }}
                    >
                      应用设计修改
                    </button>
                  )}
                </>
              )}
            </footer>
            {!request && !historical && (
              <p>编辑需要已连接的 PC 和 Relay 支持文件同步，请更新后重新连接。</p>
            )}
          </section>
        </div>
      )}
    </>
  )
}
