import { describe, expect, it, vi } from 'vitest'
import {
  createEditorState,
  editBuffer,
  reconcileExternalBuffer,
} from '../src/renderer/editor/editor-state.js'
import {
  CompileRequestQueue,
  PendingSaveRegistry,
  ProjectListingRequestGate,
  canRenameFile,
  persistBuffer,
  reconcileProjectListing,
  refreshProjectListing,
  remapWorkspacePaths,
  completeReverseSync,
  runRenameFlow,
  saveAllForCompile,
  SyncRequestGate,
} from '../src/renderer/workbench-coordination.js'

describe('renderer persistence and compile coordination', () => {
  it('does not clean a >2MiB edit when atomic save is rejected and never compiles', async () => {
    const text = 'x'.repeat(2 * 1024 * 1024 + 1)
    let state = editBuffer(
      createEditorState([{ path: 'main.tex', text: 'old', diskSha256: 'old' }]),
      'main.tex',
      text,
    )
    const saveFile = vi.fn(async () => ({
      ok: false as const,
      error: { code: 'LATEX_INVALID_PAYLOAD' as const, message: 'too large' },
    }))
    const compile = vi.fn()
    const saved = await persistBuffer({
      projectId: 'p',
      path: 'main.tex',
      getState: () => state,
      setState: (next) => (state = next),
      saveFile,
    })
    expect(saved).toBe(false)
    expect(state.buffers['main.tex']).toMatchObject({ text, dirty: true })
    expect(
      await saveAllForCompile(
        ['main.tex'],
        () => Promise.resolve(false),
        () => state,
      ),
    ).toBe(false)
    expect(compile).not.toHaveBeenCalled()
  })

  it('reconciles the exact saved snapshot while preserving edits typed during save', async () => {
    let state = editBuffer(
      createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha1' }]),
      'main.tex',
      'v2',
    )
    let release!: (value: any) => void
    const saving = persistBuffer({
      projectId: 'p',
      path: 'main.tex',
      getState: () => state,
      setState: (next) => (state = next),
      saveFile: () => new Promise((resolve) => (release = resolve)),
    })
    state = editBuffer(state, 'main.tex', 'v3')
    release({
      ok: true,
      value: {
        savedText: 'v2',
        diskSha256: 'sha2',
        buffer: { path: 'main.tex', text: 'v3', dirty: true, conflict: null },
      },
    })
    expect(await saving).toBe(true)
    expect(state.buffers['main.tex']).toMatchObject({
      text: 'v3',
      diskText: 'v2',
      diskSha256: 'sha2',
      dirty: true,
    })
  })

  it('settles a watcher self-save echo before the save response and remains compilable', async () => {
    let state = editBuffer(
      createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha1' }]),
      'main.tex',
      'v2',
    )
    let release!: (value: any) => void
    const saving = persistBuffer({
      projectId: 'p',
      path: 'main.tex',
      getState: () => state,
      setState: (next) => (state = next),
      saveFile: () => new Promise((resolve) => (release = resolve)),
    })
    state = reconcileExternalBuffer(state, {
      path: 'main.tex',
      text: 'v2',
      diskSha256: 'sha2',
      dirty: false,
      conflict: null,
    })
    expect(state.buffers['main.tex']).toMatchObject({ dirty: false, conflict: null })
    release({
      ok: true,
      value: {
        savedText: 'v2',
        diskSha256: 'sha2',
        buffer: {
          path: 'main.tex',
          text: 'v2',
          diskSha256: 'sha2',
          dirty: false,
          conflict: null,
        },
      },
    })
    expect(await saving).toBe(true)
    expect(
      await saveAllForCompile(
        ['main.tex'],
        () => Promise.resolve(true),
        () => state,
      ),
    ).toBe(true)
  })

  it('reconciles a v2 watcher echo while v3 is current and allows the next save and compile', async () => {
    let state = editBuffer(
      createEditorState([{ path: 'main.tex', text: 'v1', diskText: 'v1', diskSha256: 'sha1' }]),
      'main.tex',
      'v2',
    )
    const pendingSaves = new PendingSaveRegistry()
    let release!: (value: any) => void
    const saving = persistBuffer({
      projectId: 'p',
      path: 'main.tex',
      getState: () => state,
      setState: (next) => (state = next),
      pendingSaves,
      saveFile: () => new Promise((resolve) => (release = resolve)),
    })
    await vi.waitFor(() => expect(pendingSaves.get('main.tex')).toBeDefined())
    state = editBuffer(state, 'main.tex', 'v3')
    const echo = {
      path: 'main.tex',
      text: 'v2',
      diskText: 'v2',
      diskSha256: 'sha2',
      dirty: false,
      conflict: null,
    }
    const snapshot = pendingSaves.match(echo)
    expect(snapshot).toBeDefined()
    state = reconcileExternalBuffer(state, echo, snapshot)
    expect(state.buffers['main.tex']).toMatchObject({
      text: 'v3',
      diskText: 'v2',
      diskSha256: 'sha2',
      dirty: true,
      conflict: null,
    })
    expect(reconcileExternalBuffer(state, echo, pendingSaves.match(echo))).toBe(state)
    release({
      ok: true,
      value: {
        savedText: 'v2',
        diskSha256: 'sha2',
        buffer: { ...echo, text: 'v3', dirty: true },
      },
    })
    expect(await saving).toBe(true)
    expect(pendingSaves.get('main.tex')).toBeUndefined()
    expect(state.buffers['main.tex']).toMatchObject({ text: 'v3', diskText: 'v2', dirty: true })
    state = reconcileExternalBuffer(state, echo, pendingSaves.match(echo))
    expect(state.buffers['main.tex']).toMatchObject({
      text: 'v3',
      diskText: 'v2',
      dirty: true,
      conflict: null,
    })
  })

  it('coalesces edits during a long compile into exactly one latest rerun and cancellation drops it', async () => {
    const queue = new CompileRequestQueue()
    const releases: Array<() => void> = []
    const run = vi.fn(() => new Promise<void>((resolve) => releases.push(resolve)))
    queue.request(run)
    queue.request(run)
    queue.request(run)
    expect(run).toHaveBeenCalledTimes(1)
    releases.shift()!()
    await vi.waitFor(() => expect(run).toHaveBeenCalledTimes(2))
    queue.request(run)
    queue.cancelPending()
    releases.shift()!()
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(run).toHaveBeenCalledTimes(2)
  })

  it('rejects stale SyncTeX replies after revision or active path changes', () => {
    const gate = new SyncRequestGate()
    const first = gate.begin(2, 'main.tex')
    expect(gate.accept(first, 3, 'main.tex')).toBe(false)
    const second = gate.begin(3, 'main.tex')
    expect(gate.accept(second, 3, 'chapter.tex')).toBe(false)
    const third = gate.begin(3, 'chapter.tex')
    expect(gate.accept(third, 3, 'chapter.tex')).toBe(true)
  })

  it('remaps files, open tabs, and active path while refusing main-file rename', () => {
    expect(canRenameFile('main.tex', 'main.tex')).toBe(false)
    expect(
      remapWorkspacePaths(['a.tex', 'main.tex'], ['a.tex'], 'a.tex', 'a.tex', 'b.tex'),
    ).toEqual({ files: ['b.tex', 'main.tex'], openPaths: ['b.tex'], activePath: 'b.tex' })
  })

  it('adds and removes AI-created files from the tree and tabs without discarding dirty buffers', () => {
    const applied = reconcileProjectListing(
      createEditorState([{ path: 'main.tex', text: 'main', diskSha256: 'main-sha' }]),
      ['main.tex'],
      ['main.tex'],
      'main.tex',
      ['main.tex', 'generated.tex'],
    )
    expect(applied.files).toEqual(['generated.tex', 'main.tex'])

    const withGeneratedOpen = createEditorState([
      { path: 'main.tex', text: 'main', diskSha256: 'main-sha' },
      { path: 'generated.tex', text: 'generated', diskSha256: 'generated-sha' },
    ])
    const undone = reconcileProjectListing(
      withGeneratedOpen,
      ['generated.tex', 'main.tex'],
      ['main.tex', 'generated.tex'],
      'generated.tex',
      ['main.tex'],
    )
    expect(undone.files).toEqual(['main.tex'])
    expect(undone.openPaths).toEqual(['main.tex'])
    expect(undone.activePath).toBe('main.tex')
    expect(undone.editorState.buffers['generated.tex']).toBeUndefined()

    const dirty = editBuffer(withGeneratedOpen, 'generated.tex', 'unsaved local edit')
    const preserved = reconcileProjectListing(
      dirty,
      ['generated.tex', 'main.tex'],
      ['main.tex', 'generated.tex'],
      'generated.tex',
      ['main.tex'],
    )
    expect(preserved.files).toContain('generated.tex')
    expect(preserved.openPaths).toContain('generated.tex')
    expect(preserved.editorState.buffers['generated.tex']?.text).toBe('unsaved local edit')
    expect(preserved.editorState.buffers['generated.tex']?.dirty).toBe(true)
  })

  it('ignores deferred listing responses made stale by a local create or rename', async () => {
    const gate = new ProjectListingRequestGate()
    let resolveFirst!: (files: string[]) => void
    const firstResponse = new Promise<string[]>((resolve) => (resolveFirst = resolve))
    const applyFirst = vi.fn()
    const firstRefresh = refreshProjectListing(gate, () => firstResponse, applyFirst)
    gate.invalidate()
    resolveFirst(['main.tex'])
    await expect(firstRefresh).resolves.toBe(false)
    expect(applyFirst).not.toHaveBeenCalled()

    let resolveSecond!: (files: string[]) => void
    const secondResponse = new Promise<string[]>((resolve) => (resolveSecond = resolve))
    const applySecond = vi.fn()
    const secondRefresh = refreshProjectListing(gate, () => secondResponse, applySecond)
    gate.invalidate()
    resolveSecond(['main.tex', 'old-name.tex'])
    await expect(secondRefresh).resolves.toBe(false)
    expect(applySecond).not.toHaveBeenCalled()
  })

  it('aborts compile when a dirty buffer is added during a deferred save', async () => {
    let state = editBuffer(
      createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha1' }]),
      'main.tex',
      'v2',
    )
    let release!: () => void
    const gate = new Promise<void>((resolve) => (release = resolve))
    const saving = saveAllForCompile(
      ['main.tex'],
      async () => {
        await gate
        return true
      },
      () => state,
    )
    state = {
      ...editBuffer(state, 'main.tex', 'v3'),
      buffers: {
        ...state.buffers,
        'new.tex': { ...state.buffers['main.tex']!, path: 'new.tex', text: 'dirty', dirty: true },
      },
    }
    release()
    expect(await saving).toBe(false)
  })

  it('does not activate reverse SyncTeX after its deferred load becomes stale', async () => {
    const gate = new SyncRequestGate()
    const token = gate.begin(1, 'main.tex')
    let release!: () => void
    const loaded = new Promise<void>((resolve) => (release = resolve))
    const activate = vi.fn()
    const completion = completeReverseSync({
      token,
      gate,
      current: () => ({ revision: 1, path: 'main.tex' }),
      load: () => loaded,
      activate,
    })
    gate.begin(2, 'main.tex')
    release()
    expect(await completion).toBe(false)
    expect(activate).not.toHaveBeenCalled()
  })

  it('keeps debounce on cancelled rename and flushes dirty text before rename', async () => {
    const cancelTimer = vi.fn()
    const scheduleTimer = vi.fn()
    const save = vi.fn(async () => true)
    const rename = vi.fn(async () => true)
    expect(
      await runRenameFlow({
        prompt: () => null,
        from: 'a.tex',
        cancelTimer,
        scheduleTimer,
        save,
        isClean: () => false,
        rename,
        apply: vi.fn(),
      }),
    ).toBe(false)
    expect(cancelTimer).not.toHaveBeenCalled()
    expect(
      await runRenameFlow({
        prompt: () => 'b.tex',
        from: 'a.tex',
        cancelTimer,
        scheduleTimer,
        save,
        isClean: () => true,
        rename,
        apply: vi.fn(),
      }),
    ).toBe(true)
    expect(cancelTimer).toHaveBeenCalledOnce()
    expect(save).toHaveBeenCalledWith('a.tex')
    expect(rename).toHaveBeenCalledWith('a.tex', 'b.tex')
    expect(scheduleTimer).not.toHaveBeenCalled()
  })
})
