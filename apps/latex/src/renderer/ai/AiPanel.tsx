import { useEffect, useRef, useState } from 'react'
import { AgentLoop } from '@wiswork/agent-core'
import { createLatexSkill } from './latex-skill.js'
import {
  loadProposalForReview,
  proposalForSelection,
  type ReviewProposal,
} from './proposal-review.js'
import { ProposalReview } from './ProposalReview.js'
import { createLatexTransport } from './transport.js'

export function AiPanel({
  projectId,
  disabled = false,
  onProjectFilesChanged,
}: {
  projectId: string
  disabled?: boolean
  onProjectFilesChanged?: () => void | Promise<void>
}) {
  const [input, setInput] = useState('')
  const [text, setText] = useState('')
  const [busy, setBusy] = useState(false)
  const [proposal, setProposal] = useState<ReviewProposal | null>(null)
  const [snapshotId, setSnapshotId] = useState<string | null>(null)
  const [status, setStatus] = useState<string | null>(null)
  const projectRef = useRef(projectId)
  projectRef.current = projectId
  const loopRef = useRef<AgentLoop | null>(null)
  const chatIdsRef = useRef<{ projectId: string; chatId: string } | null>(null)

  useEffect(() => {
    const loop = new AgentLoop({
      transport: createLatexTransport(),
      skill: createLatexSkill(window.latexApi, () => projectRef.current),
      events: {
        onText: setText,
        onToolExecuted: ({ call, execution }) => {
          if (call.name !== 'propose_project_edits' || execution.isError) return
          const projectId = projectRef.current
          void loadProposalForReview(execution.output, projectId, (request) =>
            window.latexApi.getProposal(request),
          )
            .then(setProposal)
            .catch((error: unknown) =>
              setStatus(error instanceof Error ? error.message : String(error)),
            )
        },
        onDone: (result) => {
          setBusy(false)
          const ids = chatIdsRef.current
          if (ids)
            void window.latexApi.appendDirectoryChat({
              projectId: projectRef.current,
              storeProjectId: ids.projectId,
              chatId: ids.chatId,
              role: 'assistant',
              text: result.text,
            })
        },
        onError: (error) => {
          setStatus(error)
          setBusy(false)
        },
      },
    })
    loopRef.current = loop
    return () => {
      loop.cancel()
      loopRef.current = null
    }
  }, [])

  useEffect(() => {
    void window.latexApi.resolveDirectoryChat({ projectId }).then(async (result) => {
      if (!result.ok) return
      chatIdsRef.current = result.value
      const loaded = await window.latexApi.loadDirectoryChat({
        projectId,
        storeProjectId: result.value.projectId,
        chatId: result.value.chatId,
        limit: 200,
      })
      if (loaded.ok)
        loopRef.current?.restore(
          loaded.value.map((message) => ({ role: message.role, text: message.text })),
        )
    })
  }, [projectId])

  const send = () => {
    const instruction = input.trim()
    if (!instruction || busy || disabled) return
    setBusy(true)
    setStatus(null)
    setText('')
    setInput('')
    const ids = chatIdsRef.current
    if (ids)
      void window.latexApi.appendDirectoryChat({
        projectId,
        storeProjectId: ids.projectId,
        chatId: ids.chatId,
        role: 'user',
        text: instruction,
      })
    loopRef.current?.run(instruction)
  }

  const confirm = async (selected: ReadonlySet<string>) => {
    if (!proposal || disabled) return
    setBusy(true)
    try {
      const owned = await proposalForSelection(proposal, selected, async (files) => {
        const result = await window.latexApi.proposeProjectEdits({ projectId, files })
        if (!result.ok) throw new Error(result.error.message)
        return result.value
      })
      const result = await window.latexApi.applyProposal({ projectId, proposalId: owned.id })
      if (!result.ok) throw new Error(result.error.message)
      const value = result.value as {
        snapshotId: string
        compile: { ok: boolean; error?: string; result?: { diagnostics: unknown[] } }
      }
      await onProjectFilesChanged?.()
      setSnapshotId(value.snapshotId)
      setProposal(null)
      setStatus(
        value.compile.ok
          ? `Changes applied and compiled (${value.compile.result?.diagnostics.length ?? 0} diagnostics).`
          : `Changes applied; compile failed: ${value.compile.error ?? 'unknown error'}`,
      )
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error))
    } finally {
      setBusy(false)
    }
  }

  const undo = async () => {
    if (!snapshotId || disabled) return
    setBusy(true)
    const result = await window.latexApi.undoProposal({ projectId, snapshotId })
    if (result.ok) {
      await onProjectFilesChanged?.()
      setSnapshotId(null)
      setStatus('AI changes were undone and the project was compiled again.')
    } else setStatus(result.error.message)
    setBusy(false)
  }

  return (
    <aside className="latex-ai-panel">
      <h2>WisWork AI</h2>
      {text && <div className="ai-response">{text}</div>}
      {proposal && (
        <ProposalReview
          proposal={proposal}
          busy={busy || disabled}
          onConfirm={(selected) => void confirm(selected)}
          onCancel={() => setProposal(null)}
        />
      )}
      {status && <div role="status">{status}</div>}
      {snapshotId && (
        <button disabled={busy || disabled} onClick={() => void undo()}>
          Undo AI changes
        </button>
      )}
      <textarea
        aria-label="Ask WisWork AI"
        value={input}
        disabled={busy || disabled}
        onChange={(event) => setInput(event.target.value)}
      />
      <button disabled={busy || disabled || !input.trim()} onClick={send}>
        {busy ? 'Working…' : 'Send'}
      </button>
    </aside>
  )
}
