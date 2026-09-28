import { readUntilConverged } from '../shared/office-write-transaction.js'

export interface PresentationImportSnapshot {
  slideIds: string[]
  /** Structural baseline only; document identity is checked by the caller. */
  fingerprint: string
}

export interface PresentationImportReceipt {
  /** IDs of the newly appended slides, in presentation order. */
  slideIds: string[]
}

export interface PresentationImportAdapter {
  available(): boolean
  snapshot(signal?: AbortSignal): Promise<PresentationImportSnapshot>
  /** Read-only export used to prove an interrupted single-page append. */
  exportPage?(slideId: string, signal?: AbortSignal): Promise<string>
  insert(
    base64: string,
    expectedSlideCount: number,
    before: PresentationImportSnapshot,
    signal?: AbortSignal,
  ): Promise<PresentationImportReceipt>
  insertPage?(
    base64: string,
    sourceSlideId: string,
    before: PresentationImportSnapshot,
    signal?: AbortSignal,
  ): Promise<PresentationImportReceipt>
  verify(
    receipt: PresentationImportReceipt,
    before: PresentationImportSnapshot,
    signal?: AbortSignal,
  ): Promise<boolean>
}

function cancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new Error('cancelled')
}

function validIds(ids: unknown): ids is string[] {
  return (
    Array.isArray(ids) &&
    ids.every((id) => typeof id === 'string' && id.length > 0) &&
    new Set(ids).size === ids.length
  )
}

function validBaseline(before: PresentationImportSnapshot): boolean {
  return Boolean(
    before && validIds(before.slideIds) && before.fingerprint === JSON.stringify(before.slideIds),
  )
}

async function read(context: PowerPoint.RequestContext): Promise<PresentationImportSnapshot> {
  context.presentation.slides.load('items/id')
  await context.sync()
  const slideIds = context.presentation.slides.items.map((slide) => slide.id)
  if (!validIds(slideIds)) throw new Error('office_state_uncertain')
  return { slideIds, fingerprint: JSON.stringify(slideIds) }
}

export function createBrowserPresentationImportAdapter(): PresentationImportAdapter {
  const available = () => {
    try {
      return (
        typeof PowerPoint !== 'undefined' &&
        typeof PowerPoint.run === 'function' &&
        typeof Office !== 'undefined' &&
        String(Office.context.host) === 'PowerPoint' &&
        Office.context.requirements.isSetSupported('PowerPointApi', '1.2')
      )
    } catch {
      return false
    }
  }
  const requireApi = () => {
    if (!available()) throw new Error('office_api_unsupported')
  }
  const snapshot = async (signal?: AbortSignal) => {
    cancelled(signal)
    requireApi()
    const value = await PowerPoint.run(read)
    cancelled(signal)
    return value
  }
  const insert = async (
    base64: string,
    expectedSlideCount: number,
    before: PresentationImportSnapshot,
    signal?: AbortSignal,
    sourceSlideId?: string,
  ): Promise<PresentationImportReceipt> => {
    cancelled(signal)
    if (
      typeof base64 !== 'string' ||
      base64.length === 0 ||
      base64.length > 16 * 1024 * 1024 ||
      base64.length % 4 !== 0 ||
      !/^[A-Za-z0-9+/]+={0,2}$/.test(base64) ||
      !Number.isSafeInteger(expectedSlideCount) ||
      expectedSlideCount < 1 ||
      !validBaseline(before) ||
      (sourceSlideId !== undefined &&
        (!/^[1-9][0-9]{0,9}#$/.test(sourceSlideId) ||
          Number(sourceSlideId.slice(0, -1)) < 256 ||
          Number(sourceSlideId.slice(0, -1)) > 4294967295))
    ) {
      throw new Error('invalid_tool_input')
    }
    requireApi()
    const baseline = { slideIds: [...before.slideIds], fingerprint: before.fingerprint }
    let attempted = false
    try {
      return await PowerPoint.run(async (context) => {
        if (typeof context.presentation.insertSlidesFromBase64 !== 'function')
          throw new Error('office_api_unsupported')
        const current = await read(context)
        cancelled(signal)
        if (current.fingerprint !== baseline.fingerprint) throw new Error('proposal_stale')
        const lastId = baseline.slideIds.at(-1)
        attempted = true
        context.presentation.insertSlidesFromBase64(base64, {
          formatting: 'KeepSourceFormatting',
          ...(lastId
            ? { targetSlideId: sourceSlideId && !lastId.includes('#') ? `${lastId}#` : lastId }
            : {}),
          ...(sourceSlideId ? { sourceSlideIds: [sourceSlideId] } : {}),
        })
        let syncFailed = false
        try {
          await context.sync()
        } catch {
          syncFailed = true
        }
        // Once queued, cancellation does not establish whether Office committed the write.
        // Always reconcile without the caller's abort signal, and never delete uncertain pages.
        const expected = (state: PresentationImportSnapshot) =>
          state.slideIds.length === baseline.slideIds.length + expectedSlideCount &&
          baseline.slideIds.every((id, index) => state.slideIds[index] === id)
        const after = await readUntilConverged({ read: () => read(context), accept: expected })
        if (syncFailed || !expected(after)) throw new Error('office_state_uncertain')
        return { slideIds: after.slideIds.slice(baseline.slideIds.length) }
      })
    } catch (cause) {
      if (attempted) throw new Error('office_state_uncertain', { cause })
      throw cause
    }
  }
  return {
    available,
    snapshot,
    async exportPage(slideId, signal) {
      cancelled(signal)
      requireApi()
      if (!slideId || slideId.length > 256) throw new Error('invalid_tool_input')
      const base64 = await PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItem(slideId)
        if (typeof slide.exportAsBase64 !== 'function') throw new Error('office_api_unsupported')
        const exported = slide.exportAsBase64()
        await context.sync()
        return exported.value
      })
      cancelled(signal)
      if (typeof base64 !== 'string' || !base64) throw new Error('office_state_uncertain')
      return base64
    },
    insert,
    insertPage: (base64, sourceSlideId, before, signal) =>
      insert(base64, 1, before, signal, sourceSlideId),
    async verify(receipt, before, _signal) {
      if (
        !validBaseline(before) ||
        !receipt ||
        !validIds(receipt.slideIds) ||
        receipt.slideIds.length === 0
      )
        return false
      const expected = [...before.slideIds, ...receipt.slideIds]
      if (!validIds(expected)) return false
      // Verification is reconciliation after a write, including when Stop raced the commit.
      return (await snapshot()).fingerprint === JSON.stringify(expected)
    },
  }
}
