import { describe, expect, it } from 'vitest'
import {
  acceptCompileResult,
  addEditorBuffer,
  beginCompile,
  beginSave,
  completeSave,
  createEditorState,
  editBuffer,
  recordExternalChange,
  reconcileExternalBuffer,
  renameEditorBuffer,
} from '../src/renderer/editor/editor-state.js'

describe('LaTeX editor three-layer revisions', () => {
  it('keeps a newer edit dirty when an older autosave completes', () => {
    let state = createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha-v1' }])
    state = editBuffer(state, 'main.tex', 'v2')
    const save = beginSave(state, 'main.tex')
    state = editBuffer(state, 'main.tex', 'v3')
    state = completeSave(state, save, 'sha-v2')
    expect(state.buffers['main.tex']).toMatchObject({
      text: 'v3',
      diskSha256: 'sha-v2',
      dirty: true,
    })
  })

  it('never lets a stale compile result refresh the current preview', () => {
    let state = createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha-v1' }])
    const first = beginCompile(state)
    state = first.state
    state = editBuffer(state, 'main.tex', 'v2')
    state = completeSave(state, beginSave(state, 'main.tex'), 'sha-v2')
    const second = beginCompile(state)
    state = acceptCompileResult(second.state, {
      revision: second.request.revision,
      pdfUrl: 'wiswork-latex-pdf://project/2',
    })
    state = acceptCompileResult(state, {
      revision: first.request.revision,
      pdfUrl: 'wiswork-latex-pdf://project/1',
    })
    expect(state.preview).toEqual({
      revision: second.request.revision,
      pdfUrl: 'wiswork-latex-pdf://project/2',
      compiledWorkspaceRevision: second.request.workspaceRevision,
    })
  })

  it('marks a PDF compiled from an older workspace revision as stale', () => {
    let state = createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha1' }])
    const compile = beginCompile(state)
    state = editBuffer(compile.state, 'main.tex', 'v2')
    state = acceptCompileResult(state, {
      revision: compile.request.revision,
      workspaceRevision: compile.request.workspaceRevision,
      pdfUrl: 'wiswork-latex-pdf://p/1',
    })
    expect(state.preview).toMatchObject({ compiledWorkspaceRevision: 0 })
    expect(state.previewStale).toBe(true)
  })

  it('does not overwrite a dirty buffer when disk changes externally', () => {
    let state = createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha-v1' }])
    state = editBuffer(state, 'main.tex', 'local')
    state = recordExternalChange(state, 'main.tex', 'disk', 'sha-disk')
    expect(state.buffers['main.tex']).toMatchObject({
      text: 'local',
      dirty: true,
      conflict: { diskText: 'disk', diskSha256: 'sha-disk' },
    })
  })

  it('keeps an external conflict that arrives while an autosave is pending', () => {
    let state = createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha-v1' }])
    state = editBuffer(state, 'main.tex', 'local')
    const save = beginSave(state, 'main.tex')
    state = recordExternalChange(state, 'main.tex', 'external', 'sha-external')
    state = completeSave(state, save, 'sha-local')
    expect(state.buffers['main.tex']).toMatchObject({
      text: 'local',
      dirty: true,
      conflict: { diskText: 'external', diskSha256: 'sha-external' },
    })
  })

  it('atomically remaps a renamed buffer without leaving a ghost path', () => {
    let state = editBuffer(
      createEditorState([{ path: 'a.tex', text: 'v1', diskSha256: 'sha1' }]),
      'a.tex',
      'dirty',
    )
    state = recordExternalChange(state, 'a.tex', 'external', 'sha2')
    state = renameEditorBuffer(state, 'a.tex', 'b.tex')
    expect(state.buffers['a.tex']).toBeUndefined()
    expect(state.buffers['b.tex']).toMatchObject({
      path: 'b.tex',
      text: 'dirty',
      dirty: true,
      conflict: { diskText: 'external' },
    })
  })

  it('bumps one workspace generation for external clean reloads and dirty conflicts', () => {
    let clean = createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha1' }])
    clean = recordExternalChange(clean, 'main.tex', 'external', 'sha2')
    expect(clean.workspaceRevision).toBe(1)
    let dirty = editBuffer(
      createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha1' }]),
      'main.tex',
      'local',
    )
    dirty = recordExternalChange(dirty, 'main.tex', 'external', 'sha2')
    expect(dirty.workspaceRevision).toBe(2)
  })

  it('distinguishes initial buffers from tracked loads and creates, bumping rename too', () => {
    let state = createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha1' }])
    expect(state.workspaceRevision).toBe(0)
    const preview = beginCompile(state)
    state = acceptCompileResult(preview.state, {
      revision: preview.request.revision,
      workspaceRevision: preview.request.workspaceRevision,
      pdfUrl: 'wiswork-latex-pdf://p/1',
    })
    state = addEditorBuffer(
      state,
      {
        path: 'existing.tex',
        text: 'local',
        diskSha256: 'e1',
        dirty: true,
        conflict: { diskText: 'disk', diskSha256: 'e2' },
      },
      'tracked',
    )
    expect(state.workspaceRevision).toBe(1)
    expect(state.previewStale).toBe(true)
    expect(state.buffers['existing.tex']).toMatchObject({
      dirty: true,
      conflict: { diskText: 'disk', diskSha256: 'e2' },
    })
    state = addEditorBuffer(state, { path: 'created.tex', text: '', diskSha256: 'c1' }, 'created')
    expect(state.workspaceRevision).toBe(2)
    state = renameEditorBuffer(state, 'created.tex', 'renamed.tex')
    expect(state.workspaceRevision).toBe(3)
  })

  it('settles a self-save watcher echo once without changing workspace revision', () => {
    let state = editBuffer(
      createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha1' }]),
      'main.tex',
      'v2',
    )
    const revision = state.workspaceRevision
    state = reconcileExternalBuffer(state, {
      path: 'main.tex',
      text: 'v2',
      diskSha256: 'sha2',
      dirty: false,
      conflict: null,
    })
    expect(state.workspaceRevision).toBe(revision)
    expect(state.buffers['main.tex']).toMatchObject({ dirty: false, conflict: null })
    const duplicate = reconcileExternalBuffer(state, {
      path: 'main.tex',
      text: 'v2',
      diskSha256: 'sha2',
      dirty: false,
      conflict: null,
    })
    expect(duplicate).toBe(state)
  })

  it('uses authoritative disk text for initial and tracked dirty buffers', () => {
    let initial = createEditorState([
      {
        path: 'main.tex',
        text: 'local',
        diskText: 'disk',
        diskSha256: 'sha-disk',
        dirty: true,
        conflict: null,
      },
    ])
    expect(initial.buffers['main.tex'].dirty).toBe(true)
    initial = editBuffer(initial, 'main.tex', 'local')
    expect(initial.buffers['main.tex'].dirty).toBe(true)
    initial = editBuffer(initial, 'main.tex', 'disk')
    expect(initial.buffers['main.tex'].dirty).toBe(false)

    let tracked = createEditorState([])
    tracked = addEditorBuffer(
      tracked,
      {
        path: 'chapter.tex',
        text: 'local',
        diskText: 'disk',
        diskSha256: 'sha-disk',
        dirty: true,
        conflict: { diskText: 'external', diskSha256: 'sha-external' },
      },
      'tracked',
    )
    tracked = editBuffer(tracked, 'chapter.tex', 'disk')
    expect(tracked.buffers['chapter.tex']).toMatchObject({
      dirty: false,
      conflict: { diskText: 'external', diskSha256: 'sha-external' },
    })
  })

  it('bumps and stales tracked dirty or conflict buffers but not tracked clean buffers', () => {
    let clean = createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha1' }])
    clean = addEditorBuffer(
      clean,
      { path: 'clean.tex', text: 'disk', diskText: 'disk', diskSha256: 'clean' },
      'tracked',
    )
    expect(clean.workspaceRevision).toBe(0)

    const compile = beginCompile(
      createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha1' }]),
    )
    let dirty = addEditorBuffer(
      compile.state,
      {
        path: 'hidden.tex',
        text: 'local',
        diskText: 'disk',
        diskSha256: 'dirty',
        dirty: true,
        conflict: null,
      },
      'tracked',
    )
    expect(dirty.workspaceRevision).toBe(1)
    dirty = acceptCompileResult(dirty, {
      revision: compile.request.revision,
      workspaceRevision: compile.request.workspaceRevision,
      pdfUrl: 'wiswork-latex-pdf://p/1',
    })
    expect(dirty.previewStale).toBe(true)

    let conflict = createEditorState([])
    conflict = addEditorBuffer(
      conflict,
      {
        path: 'conflict.tex',
        text: 'disk',
        diskText: 'disk',
        diskSha256: 'base',
        dirty: false,
        conflict: { diskText: 'external', diskSha256: 'external' },
      },
      'tracked',
    )
    expect(conflict.workspaceRevision).toBe(1)
  })

  it('keeps an in-flight compile stale after a tracked create or external change', () => {
    let state = createEditorState([{ path: 'main.tex', text: 'v1', diskSha256: 'sha1' }])
    const compile = beginCompile(state)
    state = addEditorBuffer(
      compile.state,
      { path: 'new.tex', text: '', diskSha256: 'new' },
      'created',
    )
    state = acceptCompileResult(state, {
      revision: compile.request.revision,
      workspaceRevision: compile.request.workspaceRevision,
      pdfUrl: 'wiswork-latex-pdf://p/1',
    })
    expect(state.previewStale).toBe(true)
  })
})
