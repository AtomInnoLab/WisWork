export interface SaveCoordinator {
  enqueue(save: () => Promise<boolean>): Promise<boolean>
  flushDirty(isDirty: () => boolean, save: () => Promise<boolean>): Promise<boolean>
}

const MAX_FLUSH_SAVES = 8

/**
 * Owns renderer save ordering. Close-save requests join the same promise chain,
 * so they cannot race an autosave or menu save and do not need polling.
 */
export function createSaveCoordinator(): SaveCoordinator {
  let tail: Promise<void> = Promise.resolve()

  const enqueue = (save: () => Promise<boolean>): Promise<boolean> => {
    const result = tail.then(save, save)
    tail = result.then(
      () => undefined,
      () => undefined,
    )
    return result
  }

  return {
    enqueue,
    flushDirty(isDirty, save) {
      return enqueue(async () => {
        for (let attempt = 0; attempt < MAX_FLUSH_SAVES && isDirty(); attempt += 1) {
          if (!(await save())) return false
        }
        return !isDirty()
      })
    },
  }
}
