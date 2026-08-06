import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { PdfPoint, ViewerLocation } from '@wiswork/pdf-viewer'
import type { CompileDiagnosticInput, EditorDiagnostic } from './compile/diagnostics.js'
import { mapCompileDiagnostics } from './compile/diagnostics.js'
import { CompilePanel } from './compile/CompilePanel.js'
import {
  acceptCompileResult,
  addEditorBuffer,
  beginCompile,
  createEditorState,
  editBuffer,
  reconcileExternalBuffer,
  renameEditorBuffer,
  type EditorState,
} from './editor/editor-state.js'
import { LatexEditor } from './editor/LatexEditor.js'
import { useLatexLocale } from './i18n/locale.js'
import { PdfPreview } from './pdf/PdfPreview.js'
import { OpenTabs } from './project/OpenTabs.js'
import { ProjectTree } from './project/ProjectTree.js'
import {
  canRenameFile,
  completeReverseSync,
  CompileRequestQueue,
  PendingSaveRegistry,
  PendingUpdateRegistry,
  RendererCloseFreeze,
  flushRendererCloseFence,
  persistBuffer,
  remapWorkspacePaths,
  runRenameFlow,
  saveAllForCompile,
  SyncRequestGate,
} from './workbench-coordination.js'

function diagnosticInput(value: unknown): CompileDiagnosticInput | null {
  if (!value || typeof value !== 'object') return null
  const item = value as Record<string, unknown>
  if (
    typeof item.path !== 'string' ||
    typeof item.line !== 'number' ||
    (item.column !== null && typeof item.column !== 'number') ||
    (item.severity !== 'error' && item.severity !== 'warning') ||
    typeof item.message !== 'string'
  ) {
    return null
  }
  return item as unknown as CompileDiagnosticInput
}

async function waitForSession() {
  for (let attempt = 0; attempt < 40; attempt += 1) {
    const result = await window.latexApi.getSession()
    if (result.ok) return result.value
    await new Promise((resolve) => window.setTimeout(resolve, 100))
  }
  throw new Error('The LaTeX project session did not become available.')
}

export function App() {
  const { t } = useLatexLocale()
  const closeFreeze = useRef(new RendererCloseFreeze())
  const [frozen, setFrozen] = useState(false)
  const [projectId, setProjectId] = useState<string | null>(null)
  const [mainFile, setMainFile] = useState<string | null>(null)
  const [files, setFiles] = useState<string[]>([])
  const [openPaths, setOpenPaths] = useState<string[]>([])
  const [activePath, setActivePath] = useState<string | null>(null)
  const activePathRef = useRef<string | null>(activePath)
  activePathRef.current = activePath
  const [editorState, setEditorState] = useState(() => createEditorState([]))
  const editorStateRef = useRef(editorState)
  const [compiling, setCompiling] = useState(false)
  const [diagnostics, setDiagnostics] = useState<EditorDiagnostic[]>([])
  const [log, setLog] = useState('')
  const [error, setError] = useState<string | null>(null)
  const [previewLocation, setPreviewLocation] = useState<ViewerLocation | null>(null)
  const [revealTarget, setRevealTarget] = useState<{ path: string; line: number } | null>(null)
  const saveTimers = useRef(new Map<string, number>())
  const saveQueues = useRef(new Map<string, Promise<boolean>>())
  const pendingSaves = useRef(new PendingSaveRegistry())
  const pendingUpdates = useRef(new PendingUpdateRegistry())
  const autoCompileTimer = useRef<number | null>(null)
  const syncTimer = useRef<number | null>(null)
  const compileLatestRef = useRef<() => void>(() => undefined)
  const compileQueue = useRef(new CompileRequestQueue())
  const forwardGate = useRef(new SyncRequestGate())
  const reverseGate = useRef(new SyncRequestGate())
  const scheduleAutoCompile = useCallback(() => {
    if (closeFreeze.current.isFrozen()) return
    if (autoCompileTimer.current !== null) window.clearTimeout(autoCompileTimer.current)
    autoCompileTimer.current = window.setTimeout(() => compileLatestRef.current(), 2_000)
  }, [])

  const replaceEditorState = useCallback((next: EditorState) => {
    editorStateRef.current = next
    setEditorState(next)
  }, [])

  const updateEditorState = useCallback(
    (update: (current: EditorState) => EditorState) => {
      replaceEditorState(update(editorStateRef.current))
    },
    [replaceEditorState],
  )

  const loadFile = useCallback(
    async (path: string) => {
      if (!projectId) return false
      if (!editorStateRef.current.buffers[path]) {
        const result = await window.latexApi.readFile({ projectId, path })
        if (!result.ok) {
          setError(result.error.message)
          return false
        }
        const before = editorStateRef.current
        const next = addEditorBuffer(
          before,
          {
            path,
            text: result.value.text,
            diskText: result.value.diskText,
            diskSha256: result.value.diskSha256,
            dirty: result.value.dirty,
            conflict: result.value.conflict
              ? {
                  diskText: result.value.conflict.diskText ?? '',
                  diskSha256: result.value.conflict.diskSha256,
                }
              : null,
          },
          'tracked',
        )
        replaceEditorState(next)
        if (next.workspaceRevision !== before.workspaceRevision) scheduleAutoCompile()
      }
      return true
    },
    [projectId, replaceEditorState, scheduleAutoCompile],
  )

  const activateFile = useCallback((path: string) => {
    setOpenPaths((current) => (current.includes(path) ? current : [...current, path]))
    setActivePath(path)
  }, [])

  const openFile = useCallback(
    async (path: string) => {
      if (await loadFile(path)) activateFile(path)
    },
    [activateFile, loadFile],
  )

  useEffect(() => {
    let disposed = false
    void waitForSession()
      .then(async (session) => {
        if (disposed) return
        setProjectId(session.projectId)
        setMainFile(session.mainFile)
        const listed = await window.latexApi.listFiles({ projectId: session.projectId })
        if (!listed.ok) throw new Error(listed.error.message)
        if (disposed) return
        setFiles(listed.value)
        const initial = session.mainFile ?? listed.value[0] ?? null
        if (initial) {
          const read = await window.latexApi.readFile({
            projectId: session.projectId,
            path: initial,
          })
          if (!read.ok) throw new Error(read.error.message)
          const state = createEditorState([
            {
              path: initial,
              text: read.value.text,
              diskText: read.value.diskText,
              diskSha256: read.value.diskSha256,
              dirty: read.value.dirty,
              conflict: read.value.conflict
                ? {
                    diskText: read.value.conflict.diskText ?? '',
                    diskSha256: read.value.conflict.diskSha256,
                  }
                : null,
            },
          ])
          if (!disposed) {
            replaceEditorState(state)
            setOpenPaths([initial])
            setActivePath(initial)
          }
        }
      })
      .catch((reason: unknown) => {
        if (!disposed) setError(reason instanceof Error ? reason.message : String(reason))
      })
    return () => {
      disposed = true
    }
  }, [replaceEditorState])

  const savePath = useCallback(
    (path: string): Promise<boolean> => {
      if (!projectId || closeFreeze.current.isFrozen()) return Promise.resolve(false)
      const previous = saveQueues.current.get(path) ?? Promise.resolve(true)
      const next = previous
        .catch(() => false)
        .then(() =>
          persistBuffer({
            projectId,
            path,
            getState: () => editorStateRef.current,
            setState: replaceEditorState,
            pendingSaves: pendingSaves.current,
            saveFile: async (request) => {
              const result = await window.latexApi.saveFile(request)
              if (!result.ok) setError(result.error.message)
              return result
            },
          }),
        )
      saveQueues.current.set(path, next)
      void next.finally(() => {
        if (saveQueues.current.get(path) === next) saveQueues.current.delete(path)
      })
      return next
    },
    [projectId, replaceEditorState],
  )

  const compileOnce = useCallback(async () => {
    if (closeFreeze.current.isFrozen() || !projectId || !mainFile) return
    const paths = Object.keys(editorStateRef.current.buffers)
    if (!(await saveAllForCompile(paths, savePath, () => editorStateRef.current))) {
      setError('Compilation stopped because not all edits were saved.')
      return
    }
    const begun = beginCompile(editorStateRef.current)
    replaceEditorState(begun.state)
    setCompiling(true)
    setError(null)
    try {
      const result = await window.latexApi.compile({
        projectId,
        mainFile,
        revision: begun.request.revision,
      })
      if (!result.ok) {
        setError(result.error.message)
        return
      }
      if (editorStateRef.current.latestCompileRevision !== result.value.revision) return
      updateEditorState((current) =>
        acceptCompileResult(current, {
          ...result.value,
          workspaceRevision: begun.request.workspaceRevision,
        }),
      )
      setLog(result.value.log)
      setDiagnostics(
        mapCompileDiagnostics(
          result.value.diagnostics.flatMap((value) => {
            const parsed = diagnosticInput(value)
            return parsed ? [parsed] : []
          }),
          new Set(files),
        ),
      )
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    } finally {
      if (editorStateRef.current.latestCompileRevision === begun.request.revision) {
        setCompiling(false)
      }
    }
  }, [files, mainFile, projectId, replaceEditorState, savePath, updateEditorState])
  const compileProject = useCallback(() => compileQueue.current.request(compileOnce), [compileOnce])
  compileLatestRef.current = compileProject

  useEffect(() => {
    const unsubscribe = window.latexApi.onExternalChange((buffer) => {
      const before = editorStateRef.current
      const next = reconcileExternalBuffer(before, buffer, pendingSaves.current.match(buffer))
      if (next === before) return
      replaceEditorState(next)
      if (next.workspaceRevision !== before.workspaceRevision) scheduleAutoCompile()
    })
    return () => unsubscribe()
  }, [replaceEditorState, scheduleAutoCompile])

  useEffect(() => {
    const unsubscribeRequest = window.latexApi.onEditFlushRequest((requestId) => {
      const preparing = closeFreeze.current.prepare(requestId, () =>
        flushRendererCloseFence({
          saveTimers: saveTimers.current,
          saveQueues: saveQueues.current,
          clearTimer: (timer) => window.clearTimeout(timer),
          cancelAutoCompile: () => {
            if (autoCompileTimer.current !== null) window.clearTimeout(autoCompileTimer.current)
            autoCompileTimer.current = null
          },
          settleUpdates: () => pendingUpdates.current.settleAll(),
        }),
      )
      setFrozen(closeFreeze.current.isFrozen())
      void preparing.then((ok) => {
        if (!ok) setFrozen(closeFreeze.current.isFrozen())
      })
      return preparing
    })
    const unsubscribeRelease = window.latexApi.onEditFlushRelease((requestId) => {
      if (closeFreeze.current.release(requestId)) setFrozen(false)
    })
    return () => {
      unsubscribeRequest()
      unsubscribeRelease()
    }
  }, [])

  useEffect(
    () => () => {
      for (const timer of saveTimers.current.values()) window.clearTimeout(timer)
      if (autoCompileTimer.current !== null) window.clearTimeout(autoCompileTimer.current)
      if (syncTimer.current !== null) window.clearTimeout(syncTimer.current)
    },
    [],
  )

  const handleEdit = (text: string) => {
    if (closeFreeze.current.isFrozen() || !activePath || !projectId) return
    updateEditorState((state) => editBuffer(state, activePath, text))
    pendingUpdates.current.track(
      window.latexApi
        .updateFile({ projectId, path: activePath, text })
        .then((result) => {
          if (!result.ok) setError(result.error.message)
          return result.ok
        })
        .catch((reason: unknown) => {
          setError(reason instanceof Error ? reason.message : String(reason))
          return false
        }),
    )
    const previous = saveTimers.current.get(activePath)
    if (previous !== undefined) window.clearTimeout(previous)
    saveTimers.current.set(
      activePath,
      window.setTimeout(() => {
        saveTimers.current.delete(activePath)
        void savePath(activePath)
      }, 600),
    )
    scheduleAutoCompile()
  }

  const createFile = async () => {
    if (closeFreeze.current.isFrozen() || !projectId) return
    const path = window.prompt('New LaTeX file path', 'chapter.tex')?.trim()
    if (!path) return
    const result = await window.latexApi.createFile({ projectId, path, text: '' })
    if (!result.ok) return setError(result.error.message)
    replaceEditorState(
      addEditorBuffer(
        editorStateRef.current,
        {
          path,
          text: result.value.text,
          diskText: result.value.diskText,
          diskSha256: result.value.diskSha256,
          dirty: result.value.dirty,
          conflict: result.value.conflict
            ? {
                diskText: result.value.conflict.diskText ?? '',
                diskSha256: result.value.conflict.diskSha256,
              }
            : null,
        },
        'created',
      ),
    )
    setFiles((current) => [...new Set([...current, path])].sort())
    activateFile(path)
    scheduleAutoCompile()
  }

  const renameFile = async (from: string) => {
    if (closeFreeze.current.isFrozen() || !projectId) return
    if (!canRenameFile(from, mainFile))
      return setError('The configured main file cannot be renamed.')
    const scheduleTimer = () => {
      saveTimers.current.set(
        from,
        window.setTimeout(() => {
          saveTimers.current.delete(from)
          void savePath(from)
        }, 600),
      )
    }
    await runRenameFlow({
      prompt: () => window.prompt('Rename file', from),
      from,
      cancelTimer: () => {
        const timer = saveTimers.current.get(from)
        if (timer !== undefined) window.clearTimeout(timer)
        saveTimers.current.delete(from)
      },
      scheduleTimer,
      save: savePath,
      isClean: () => {
        const buffer = editorStateRef.current.buffers[from]
        return Boolean(buffer && !buffer.dirty && !buffer.conflict)
      },
      rename: async (ownedFrom, to) => {
        const result = await window.latexApi.renameFile({ projectId, from: ownedFrom, to })
        if (!result.ok) setError(result.error.message)
        return result.ok
      },
      apply: (ownedFrom, to) => {
        updateEditorState((state) => renameEditorBuffer(state, ownedFrom, to))
        const remapped = remapWorkspacePaths(files, openPaths, activePathRef.current, ownedFrom, to)
        setFiles(remapped.files)
        setOpenPaths(remapped.openPaths)
        setActivePath(remapped.activePath)
        scheduleAutoCompile()
      },
    })
  }

  const forwardSync = (line: number) => {
    const preview = editorStateRef.current.preview
    if (!projectId || !activePath || !preview) return
    if (syncTimer.current !== null) window.clearTimeout(syncTimer.current)
    const path = activePath
    const token = forwardGate.current.begin(preview.revision, path)
    syncTimer.current = window.setTimeout(() => {
      void window.latexApi
        .syncTexForward({ projectId, revision: preview.revision, path, line })
        .then((result) => {
          const currentPreview = editorStateRef.current.preview
          if (
            result.ok &&
            result.value &&
            currentPreview &&
            forwardGate.current.accept(token, currentPreview.revision, activePathRef.current ?? '')
          )
            setPreviewLocation(result.value)
        })
    }, 250)
  }

  const reverseSync = async (point: PdfPoint) => {
    const preview = editorStateRef.current.preview
    if (!projectId || !preview || !activePath) return
    const token = reverseGate.current.begin(preview.revision, activePath)
    const result = await window.latexApi.syncTexReverse({
      projectId,
      revision: preview.revision,
      ...point,
    })
    if (!result.ok || !result.value) return
    const target = result.value
    await completeReverseSync({
      token,
      gate: reverseGate.current,
      current: () => {
        const currentPreview = editorStateRef.current.preview
        return currentPreview && activePathRef.current
          ? { revision: currentPreview.revision, path: activePathRef.current }
          : null
      },
      load: () => loadFile(target.path),
      activate: () => {
        activateFile(target.path)
        setRevealTarget({ path: target.path, line: target.line })
      },
    })
  }

  const activeBuffer = activePath ? editorState.buffers[activePath] : undefined
  const dirtyPaths = useMemo(
    () =>
      new Set(
        Object.values(editorState.buffers)
          .filter((buffer) => buffer.dirty)
          .map((b) => b.path),
      ),
    [editorState],
  )
  const activeDiagnostics = diagnostics.filter((item) => item.path === activePath)

  useEffect(() => {
    forwardGate.current.invalidate()
    reverseGate.current.invalidate()
    setPreviewLocation(null)
  }, [editorState.preview?.revision])

  return (
    <main className="latex-workbench">
      <ProjectTree
        files={files}
        activePath={activePath}
        onOpen={(path) => void openFile(path)}
        onCreate={() => void createFile()}
        onRename={(path) => void renameFile(path)}
        mainFile={mainFile}
      />
      <section className="editor-workspace">
        <OpenTabs
          paths={openPaths}
          activePath={activePath}
          dirty={dirtyPaths}
          onActivate={setActivePath}
          onClose={(path) => {
            setOpenPaths((current) => current.filter((item) => item !== path))
            if (activePath === path) setActivePath(openPaths.find((item) => item !== path) ?? null)
          }}
        />
        {activeBuffer ? (
          <LatexEditor
            path={activeBuffer.path}
            value={activeBuffer.text}
            diagnostics={activeDiagnostics}
            readOnly={frozen}
            onChange={handleEdit}
            onSave={() => void savePath(activeBuffer.path)}
            onCompile={compileProject}
            onCursorLine={forwardSync}
            revealLine={revealTarget?.path === activeBuffer.path ? revealTarget.line : null}
          />
        ) : (
          <div className="empty-editor">{t('projectUnavailable')}</div>
        )}
        {activeBuffer?.conflict && (
          <div role="alert" className="conflict-banner">
            {t('externalConflict')}
          </div>
        )}
        {error && (
          <div role="alert" className="error-banner">
            {error}
          </div>
        )}
        <CompilePanel
          compiling={compiling}
          diagnostics={diagnostics}
          log={log}
          onCompile={compileProject}
          onCancel={() => {
            compileQueue.current.cancelPending()
            if (projectId) void window.latexApi.cancelCompile({ projectId })
          }}
          onDiagnostic={(diagnostic) => {
            void openFile(diagnostic.path).then(() =>
              setRevealTarget({ path: diagnostic.path, line: diagnostic.lineIndex + 1 }),
            )
          }}
        />
      </section>
      <PdfPreview
        pdfUrl={editorState.preview?.pdfUrl ?? null}
        revision={editorState.preview?.revision ?? null}
        location={previewLocation}
        stale={editorState.previewStale}
        onReverseSync={(point) => void reverseSync(point)}
      />
    </main>
  )
}
