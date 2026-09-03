import React, { useEffect, useRef, useState } from 'react'

export type PresentationMessageRole = 'user' | 'assistant' | 'error'

export function PresentationMessage({
  role,
  streaming = false,
  children,
}: {
  readonly role: PresentationMessageRole
  readonly streaming?: boolean
  readonly children?: React.ReactNode
}): React.JSX.Element {
  return (
    <div className={`ai-msg ai-msg-${role}`} {...(role === 'error' ? { role: 'alert' } : {})}>
      {children}
      {streaming && <span className="streaming-cursor" aria-label="Response streaming" />}
    </div>
  )
}

export interface PresentationActivityItem {
  readonly id: string
  readonly label: string
  readonly status: 'running' | 'done' | 'error'
  readonly detail?: React.ReactNode
  readonly tooltip?: string
}

function StepIcon({ status }: { readonly status: PresentationActivityItem['status'] }) {
  if (status === 'running') {
    return <span className="ai-tool-chip-spinner" aria-label="Running" />
  }
  return (
    <span className={`ai-step-icon ${status}`} aria-label={status === 'done' ? 'Done' : 'Error'}>
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        {status === 'done' ? (
          <path d="m3 8 3 3 7-7" stroke="currentColor" strokeWidth="1.5" />
        ) : (
          <>
            <circle cx="8" cy="8" r="5.5" stroke="currentColor" strokeWidth="1.5" />
            <path d="M8 5v3.5M8 11v.1" stroke="currentColor" strokeWidth="1.5" />
          </>
        )}
      </svg>
    </span>
  )
}

export function PresentationActivityGroup({
  items,
  workingLabel,
  workedLabel,
}: {
  readonly items: readonly PresentationActivityItem[]
  readonly workingLabel: string
  readonly workedLabel: (count: number) => string
}): React.JSX.Element | null {
  const running = items.some((item) => item.status === 'running')
  const previousRunning = useRef(running)
  const [open, setOpen] = useState(running)
  const [expanded, setExpanded] = useState<ReadonlySet<string>>(() => new Set())

  useEffect(() => {
    if (running) setOpen(true)
    else if (previousRunning.current) setOpen(false)
    previousRunning.current = running
  }, [running])

  if (items.length === 0) return null
  const label = running ? workingLabel : workedLabel(items.length)

  return (
    <div className="ai-work-group">
      <button
        type="button"
        className={`ai-work-group-summary${running ? ' running' : ''}`}
        aria-expanded={open}
        onClick={() => setOpen((value) => !value)}
      >
        {running && <span className="ai-tool-chip-spinner" aria-hidden="true" />}
        <span className="ai-work-group-label">{label}</span>
        <span className={`ai-tool-chip-caret${open ? ' open' : ''}`} aria-hidden="true">
          ›
        </span>
      </button>
      <div className={`ai-work-group-body${open ? ' open' : ''}`} aria-hidden={!open}>
        <div className="ai-work-group-body-inner">
          {items.map((item) => (
            <div className="ai-step-row" key={item.id}>
              <StepIcon status={item.status} />
              <div className="ai-step-content">
                {item.detail !== undefined ? (
                  <button
                    type="button"
                    className="ai-step-title clickable"
                    title={item.tooltip}
                    aria-expanded={expanded.has(item.id)}
                    onClick={() =>
                      setExpanded((current) => {
                        const next = new Set(current)
                        if (next.has(item.id)) next.delete(item.id)
                        else next.add(item.id)
                        return next
                      })
                    }
                  >
                    {item.label}
                  </button>
                ) : (
                  <span className="ai-step-title" title={item.tooltip}>
                    {item.label}
                  </span>
                )}
                {item.detail !== undefined && expanded.has(item.id) && (
                  <div className="ai-step-detail">{item.detail}</div>
                )}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  )
}

export function PresentationEmptyState({
  title,
  body,
  prompts,
  onChoose,
}: {
  readonly title: string
  readonly body: string
  readonly prompts: readonly string[]
  readonly onChoose: (prompt: string) => void
}): React.JSX.Element {
  return (
    <div className="ai-chat-empty">
      <div className="ai-chat-empty-title">{title}</div>
      <div className="ai-chat-empty-body">{body}</div>
      <div className="ai-starter-list">
        {prompts.map((prompt) => (
          <button
            type="button"
            className="ai-starter"
            key={prompt}
            onClick={() => onChoose(prompt)}
          >
            {prompt}
          </button>
        ))}
      </div>
    </div>
  )
}
