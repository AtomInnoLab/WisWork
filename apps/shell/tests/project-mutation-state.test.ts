import { describe, expect, it } from 'vitest'
import {
  initialProjectMutationState,
  projectMutationReducer,
} from '../src/renderer/src/project-mutation-state'

describe('project mutation state', () => {
  it('surfaces only a stable operation code when an IPC call fails', () => {
    const started = projectMutationReducer(initialProjectMutationState, {
      type: 'start',
      requestId: 1,
      operation: 'create',
    })
    const failed = projectMutationReducer(started, {
      type: 'fail',
      requestId: 1,
      operation: 'create',
    })

    expect(failed).toEqual({
      activeRequestId: null,
      operation: null,
      errorCode: 'PROJECT_CREATE_FAILED',
    })
    expect(JSON.stringify(failed)).not.toContain('/Users/alice/private')
    expect(JSON.stringify(failed)).not.toContain('secret-token')
  })

  it('ignores late completions from superseded requests', () => {
    const first = projectMutationReducer(initialProjectMutationState, {
      type: 'start',
      requestId: 1,
      operation: 'rename',
    })
    const second = projectMutationReducer(first, {
      type: 'start',
      requestId: 2,
      operation: 'delete',
    })

    expect(
      projectMutationReducer(second, {
        type: 'fail',
        requestId: 1,
        operation: 'rename',
      }),
    ).toBe(second)
    expect(projectMutationReducer(second, { type: 'succeed', requestId: 1 })).toBe(second)
  })

  it('clears an earlier failure when a retry starts and settles success', () => {
    const failed = projectMutationReducer(
      projectMutationReducer(initialProjectMutationState, {
        type: 'start',
        requestId: 1,
        operation: 'delete',
      }),
      { type: 'fail', requestId: 1, operation: 'delete' },
    )
    const retrying = projectMutationReducer(failed, {
      type: 'start',
      requestId: 2,
      operation: 'delete',
    })

    expect(retrying.errorCode).toBeNull()
    expect(retrying.activeRequestId).toBe(2)
    expect(projectMutationReducer(retrying, { type: 'succeed', requestId: 2 })).toEqual(
      initialProjectMutationState,
    )
  })
})
