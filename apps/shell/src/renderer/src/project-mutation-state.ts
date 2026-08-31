export type ProjectMutationOperation = 'create' | 'rename' | 'delete'

export type ProjectMutationErrorCode =
  'PROJECT_CREATE_FAILED' | 'PROJECT_RENAME_FAILED' | 'PROJECT_DELETE_FAILED'

export interface ProjectMutationState {
  activeRequestId: number | null
  operation: ProjectMutationOperation | null
  errorCode: ProjectMutationErrorCode | null
}

export type ProjectMutationAction =
  | { type: 'start'; requestId: number; operation: ProjectMutationOperation }
  | { type: 'succeed'; requestId: number }
  | {
      type: 'fail'
      requestId: number
      operation: ProjectMutationOperation
    }

export const initialProjectMutationState: ProjectMutationState = {
  activeRequestId: null,
  operation: null,
  errorCode: null,
}

const failureCodes: Record<ProjectMutationOperation, ProjectMutationErrorCode> = {
  create: 'PROJECT_CREATE_FAILED',
  rename: 'PROJECT_RENAME_FAILED',
  delete: 'PROJECT_DELETE_FAILED',
}

export function projectMutationReducer(
  state: ProjectMutationState,
  action: ProjectMutationAction,
): ProjectMutationState {
  if (action.type === 'start') {
    return {
      activeRequestId: action.requestId,
      operation: action.operation,
      errorCode: null,
    }
  }
  if (state.activeRequestId !== action.requestId) return state
  if (action.type === 'succeed') return initialProjectMutationState
  return {
    activeRequestId: null,
    operation: null,
    errorCode: failureCodes[action.operation],
  }
}
