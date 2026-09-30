import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import {
  selectionFingerprint,
  type StructuredProposalController,
} from '../../agent/proposal-controller.js'
import {
  savePackageBackup,
  readPackageBackup,
  releasePackageBackup,
  packageBackupRefForBytes,
  type PackageBackupRef,
} from './presentation-package-backup.js'
import {
  editPowerPointPackage,
  verifyImportedPowerPointPackageContent,
  presentationPackageDigest,
  type XmlReplacement,
  type PackageEditResult,
} from './powerpoint-package.js'
import {
  readPackageSourceSlideId,
  type PackageHostSnapshot,
  type BrowserPresentationPackageEditAdapter,
} from './browser-presentation-package-edit-adapter.js'
import {
  validatePresentationPackageChange,
  validPackageTransition,
  validPackageHostId,
  type PresentationPackageChange,
  type PackageAction,
} from './presentation-package-change.js'
import {
  validatePowerPointPageScreenshot,
  type PowerPointAdapter,
} from './browser-powerpoint-adapter.js'
interface Options {
  documentId(): Promise<string>
  assertDocumentId?(expected: string): void
  available(): boolean
  adapter: Pick<BrowserPresentationPackageEditAdapter, 'inspect' | 'stage' | 'remove'> &
    Partial<Pick<PowerPointAdapter, 'screenshotSlide'>>
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  proposals: StructuredProposalController
  readPackageChange(id: string): PresentationPackageChange | undefined
  writePackageChange(
    next: PresentationPackageChange,
    expected: PresentationPackageChange | undefined,
  ): Promise<void>
}
type Proof = PackageHostSnapshot
type Matcher = Omit<PackageEditResult, 'base64'>
interface Snapshot {
  version: 1
  documentId: string
  changeId: string
  sourceKind: 'slide' | 'chart'
  sourceSlideId: string
  packageSourceSlideId: string
  original: Proof
  expected: Matcher
  restoreExpected: Matcher
}
interface Context {
  documentId: string
  changeId: string
  epoch: number
  record?: PresentationPackageChange
  signal?: AbortSignal
  originalSignal?: AbortSignal
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const encode = (bytes: Uint8Array) =>
  btoa(Array.from(bytes, (b) => String.fromCharCode(b)).join(''))
const decode = (base64: string) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
const json = (v: unknown) => new TextEncoder().encode(JSON.stringify(v))
const proof = (v: Proof): Proof => ({
  slideIds: [...v.slideIds],
  pages: v.pages.map((p) => ({ slideId: p.slideId, digest: p.digest })),
})
const validProof = (v: any): v is Proof =>
  v &&
  Object.keys(v).sort().join(',') === 'pages,slideIds' &&
  Array.isArray(v.slideIds) &&
  v.slideIds.length > 0 &&
  v.slideIds.every(validPackageHostId) &&
  new Set(v.slideIds).size === v.slideIds.length &&
  Array.isArray(v.pages) &&
  v.pages.length === v.slideIds.length &&
  v.pages.every(
    (p: any, i: number) =>
      Object.keys(p).sort().join(',') === 'digest,slideId' &&
      p.slideId === v.slideIds[i] &&
      /^[a-f0-9]{64}$/.test(p.digest),
  )
const validMatcher = (v: any, kind: 'slide' | 'chart'): v is Matcher => {
  if (
    !v ||
    typeof v !== 'object' ||
    Array.isArray(v) ||
    Object.keys(v).sort().join(',') !==
      'afterHashes,afterXml,beforeHashes,beforeXml,changedPaths,preservedHashes' ||
    !Array.isArray(v.changedPaths) ||
    v.changedPaths.length < 1 ||
    v.changedPaths.length > 32 ||
    new Set(v.changedPaths).size !== v.changedPaths.length ||
    v.changedPaths.some(
      (p: any) =>
        typeof p !== 'string' ||
        (kind === 'slide'
          ? p !== 'ppt/slides/slide1.xml'
          : !/^ppt\/charts\/(chart|style|colors)[0-9]+\.xml$/.test(p)),
    )
  )
    return false
  const keys = [...v.changedPaths].sort()
  for (const key of ['beforeHashes', 'afterHashes', 'beforeXml', 'afterXml'])
    if (
      !v[key] ||
      typeof v[key] !== 'object' ||
      Array.isArray(v[key]) ||
      !same(Object.keys(v[key]).sort(), keys) ||
      Object.values(v[key]).some((x) => typeof x !== 'string')
    )
      return false
  return (
    !!v.preservedHashes &&
    typeof v.preservedHashes === 'object' &&
    !Array.isArray(v.preservedHashes) &&
    Object.entries(v.preservedHashes).every(([p, x]) => !!p && typeof x === 'string')
  )
}
const actions = ['inspect', 'reconcile', 'resume', 'undo', 'discard', 'capture', 'review'] as const
const tools: AgentToolDef[] = actions.map((action) => ({
  name: `${action}_package_xml_change`,
  description: `${action} a persistent PowerPoint XML package change; screenshot review is historical and never whole-deck QA.`,
  inputSchema: {
    type: 'object',
    properties: {
      change_id: { type: 'string' },
      ...(action === 'capture' || action === 'review' ? { slide_id: { type: 'string' } } : {}),
      ...(action === 'review'
        ? {
            screenshot_digest: { type: 'string' },
            status: { type: 'string', enum: ['pass', 'fail'] },
            notes: { type: 'string' },
          }
        : {}),
    },
    required: [
      'change_id',
      ...(action === 'capture' || action === 'review' ? ['slide_id'] : []),
      ...(action === 'review' ? ['screenshot_digest', 'status', 'notes'] : []),
    ],
    additionalProperties: false,
  },
}))
export function createPresentationPackageEditingSkill(options: Options) {
  let epoch = 0
  let visualEpoch = 0
  const invalidateVisual = () => {
    visualEpoch++
    captures.clear()
  }
  const captures = new Map<string, { digest: string; record: string; epoch: number }>()
  const syncGuard = (c: Context) => {
    if (c.signal?.aborted || c.originalSignal?.aborted || c.epoch !== epoch)
      throw Error('cancelled')
    if (!options.available()) throw Error('presentation_package_persistence_unavailable')
    options.assertDocumentId?.(c.documentId)
    if (!same(options.readPackageChange(c.changeId), c.record))
      throw Error('presentation_package_stale')
  }
  const guard = async (c: Context) => {
    syncGuard(c)
    if ((await options.documentId()) !== c.documentId) throw Error('presentation_document_changed')
    syncGuard(c)
  }
  const awaited = async <T>(c: Context, fn: () => Promise<T>): Promise<T> => {
    await guard(c)
    const result = await fn()
    await guard(c)
    return result
  }
  const request = (c: Context) => (body: unknown, signal?: AbortSignal) =>
    awaited(c, () => options.request(structuredClone(body), signal))
  const read = (c: Context, backup: PackageBackupRef) =>
    awaited(c, () =>
      readPackageBackup({
        request: request(c),
        documentId: c.documentId,
        changeId: c.changeId,
        backup: structuredClone(backup),
        signal: c.signal,
      }),
    )
  const save = async (c: Context, key: string, bytes: Uint8Array) => {
    const ref = await awaited(c, () =>
      savePackageBackup({
        request: request(c),
        documentId: c.documentId,
        changeId: c.changeId,
        key,
        bytes: Uint8Array.from(bytes),
        signal: c.signal,
      }),
    )
    await read(c, ref)
    return ref
  }
  const readJson = async (c: Context, ref: PackageBackupRef) =>
    JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(await read(c, ref)))
  const inspect = async (c: Context, exports: string[] = []) => {
    const actual = await awaited(c, async () =>
      structuredClone(await options.adapter.inspect([...exports], c.signal)),
    )
    if (!validProof(proof(actual))) throw Error('presentation_package_invalid')
    return actual
  }
  const store = async (c: Context, next: PresentationPackageChange) => {
    const owned = structuredClone(next)
    if (!validPackageTransition(c.record, owned)) throw Error('presentation_package_state_invalid')
    await guard(c)
    await options.writePackageChange(structuredClone(owned), c.record && structuredClone(c.record))
    c.record = owned
    await guard(c)
    return owned
  }
  const record = (id: string) => {
    const r = options.readPackageChange(id)
    if (!validatePresentationPackageChange(r)) throw Error('presentation_package_missing')
    return structuredClone(r)
  }
  const context = (r: PresentationPackageChange, signal?: AbortSignal): Context => ({
    documentId: r.documentId,
    changeId: r.changeId,
    record: r,
    epoch,
    signal,
  })
  const loadSnapshot = async (c: Context, r: PresentationPackageChange): Promise<Snapshot> => {
    const s = await readJson(c, r.snapshotRef)
    if (
      !s ||
      Object.keys(s).sort().join(',') !==
        'changeId,documentId,expected,original,packageSourceSlideId,restoreExpected,sourceKind,sourceSlideId,version' ||
      s.version !== 1 ||
      s.documentId !== r.documentId ||
      s.changeId !== r.changeId ||
      s.sourceKind !== r.sourceKind ||
      s.sourceSlideId !== r.sourceSlideId ||
      s.packageSourceSlideId !== r.packageSourceSlideId ||
      !validProof(s.original) ||
      !s.original.slideIds.includes(r.sourceSlideId) ||
      !validMatcher(s.expected, r.sourceKind) ||
      !validMatcher(s.restoreExpected, r.sourceKind) ||
      !same(s.expected.changedPaths, s.restoreExpected.changedPaths)
    )
      throw Error('presentation_package_backup_invalid')
    return s
  }
  const loadProof = async (c: Context, ref: PackageBackupRef) => {
    const p = await readJson(c, ref)
    if (!validProof(p)) throw Error('presentation_package_backup_invalid')
    return p
  }
  const check = async (c: Context, r: PresentationPackageChange) => {
    const p = await loadProof(c, r.currentProofRef)
    if (!same(proof(await inspect(c)), p)) throw Error('presentation_package_conflict')
    return p
  }
  const backups = async (c: Context, r: PresentationPackageChange, s: Snapshot) => {
    if (!same(await loadSnapshot(c, r), s)) throw Error('presentation_package_backup_invalid')
    const original = encode(await read(c, r.originalRef)),
      prepared = encode(await read(c, r.preparedRef))
    if (
      (await awaited(c, () => readPackageSourceSlideId(original, c.signal))) !==
        r.packageSourceSlideId ||
      (await awaited(c, () => readPackageSourceSlideId(prepared, c.signal))) !==
        r.packageSourceSlideId ||
      (await awaited(c, () => presentationPackageDigest(original, c.signal))) !==
        s.original.pages.find((p) => p.slideId === r.sourceSlideId)!.digest
    )
      throw Error('presentation_package_backup_invalid')
    const recomputed = await awaited(c, () =>
      editPowerPointPackage(
        original,
        r.sourceKind,
        s.expected.changedPaths.map((path) => ({ path, xml: s.expected.afterXml[path]! })),
        c.signal,
      ),
    )
    const { base64: recomputedBase64, ...metadata } = recomputed
    if (
      !same(metadata, s.expected) ||
      (await awaited(c, () => presentationPackageDigest(recomputedBase64, c.signal))) !==
        (await awaited(c, () => presentationPackageDigest(prepared, c.signal)))
    )
      throw Error('presentation_package_backup_invalid')
    const restored = await awaited(c, () =>
      editPowerPointPackage(
        prepared,
        r.sourceKind,
        s.expected.changedPaths.map((path) => ({ path, xml: s.expected.beforeXml[path]! })),
        c.signal,
      ),
    )
    const { base64: restoredBase64, ...restoreMetadata } = restored
    if (
      !same(restoreMetadata, s.restoreExpected) ||
      (await awaited(c, () => presentationPackageDigest(restoredBase64, c.signal))) !==
        s.original.pages.find((p) => p.slideId === r.sourceSlideId)!.digest
    )
      throw Error('presentation_package_backup_invalid')
    return { original, prepared }
  }
  const expectedAfter = (
    before: Proof,
    action: PackageAction,
    sourceId: string,
    inserted?: string,
  ) => {
    const ids = [...before.slideIds]
    if (action === 'import' || action === 'restore') {
      if (!inserted || ids.includes(inserted)) throw Error('presentation_package_unknown')
      ids.splice(ids.indexOf(sourceId) + 1, 0, inserted)
    } else ids.splice(ids.indexOf(sourceId), 1)
    return ids
  }
  const accepted = async (
    c: Context,
    r: PresentationPackageChange,
    s: Snapshot,
    before: Proof,
    action: PackageAction,
    actual: Proof,
    inserted?: string,
  ) => {
    const sourceId =
      action === 'restore' || action === 'delete_applied' || action === 'discard'
        ? r.replacementSlideId!
        : r.sourceSlideId
    if (!same(actual.slideIds, expectedAfter(before, action, sourceId, inserted))) return false
    if (
      before.pages.some(
        (p) =>
          actual.slideIds.includes(p.slideId) &&
          actual.pages.find((a) => a.slideId === p.slideId)!.digest !== p.digest,
      )
    )
      return false
    if (action === 'import' || action === 'restore') {
      const page = actual.pages.find((p) => p.slideId === inserted),
        data = await backups(c, r, s)
      if (!page?.base64) return false
      const expected = {
        ...(action === 'import' ? s.expected : s.restoreExpected),
        base64: action === 'import' ? data.prepared : data.original,
      }
      if (
        !(await awaited(c, () =>
          verifyImportedPowerPointPackageContent(page.base64!, expected, c.signal),
        ))
      )
        return false
    }
    return true
  }
  const observePending = async (c: Context, r: PresentationPackageChange, s: Snapshot) => {
    if (!r.pending) {
      await check(c, r)
      return { status: 'ready' as const }
    }
    const before = await loadProof(c, r.pending.beforeProofRef)
    const inserted = r.pending.insertedSlideId
    let actual = await inspect(c)
    if (same(proof(actual), before)) return { status: 'before' as const, actual: proof(actual) }
    if (r.pending.afterProofRef) {
      const after = await loadProof(c, r.pending.afterProofRef)
      const sourceId = ['restore', 'delete_applied', 'discard'].includes(r.pending.action)
        ? r.replacementSlideId!
        : r.sourceSlideId
      if (
        !same(
          after.slideIds,
          expectedAfter(before, r.pending.action, sourceId, r.pending.insertedSlideId),
        )
      )
        throw Error('presentation_package_backup_invalid')
      return {
        status: same(proof(actual), after) ? ('after' as const) : ('unknown' as const),
        actual: proof(actual),
        inserted,
      }
    }
    let candidate = inserted
    if (['import', 'restore'].includes(r.pending.action)) {
      const added = actual.slideIds.filter((id) => !before.slideIds.includes(id))
      if (added.length !== 1 || (candidate && candidate !== added[0]))
        return { status: 'unknown' as const }
      candidate = added[0]
      actual = await inspect(c, [candidate!])
    }
    if (await accepted(c, r, s, before, r.pending.action, actual, candidate))
      return { status: 'after' as const, actual: proof(actual), inserted: candidate }
    return { status: 'unknown' as const }
  }
  const finish = async (c: Context, r: PresentationPackageChange) => {
    const p = r.pending!
    if (!p.afterProofRef) throw Error('presentation_package_unknown')
    const states = {
      import: 'staged',
      delete_source: 'applied',
      restore: 'restore_staged',
      delete_applied: 'undone',
      discard: 'discarded',
    } as const
    return store(c, {
      ...r,
      state: states[p.action],
      currentProofRef: p.afterProofRef,
      pending: undefined,
      ...(p.action === 'import'
        ? { replacementSlideId: p.insertedSlideId }
        : p.action === 'restore'
          ? { restoredSlideId: p.insertedSlideId }
          : {}),
      receipts: [
        ...r.receipts,
        {
          action: p.action,
          proofRef: p.afterProofRef,
          ...(p.insertedSlideId ? { slideId: p.insertedSlideId } : {}),
        },
      ],
      reviews: [],
    })
  }
  const perform = async (
    c: Context,
    r: PresentationPackageChange,
    s: Snapshot,
    action: PackageAction,
  ) => {
    if (r.pending) throw Error('presentation_package_unknown')
    const data = await backups(c, r, s)
    await check(c, r)
    r = await store(c, {
      ...r,
      pending: { action, beforeProofRef: r.currentProofRef },
      reviews: [],
    })
    const before = await check(c, r)
    invalidateVisual()
    const sourceId =
      action === 'restore' || action === 'delete_applied' || action === 'discard'
        ? r.replacementSlideId!
        : r.sourceSlideId
    if (action === 'import' || action === 'restore') {
      await options.adapter.stage(
        {
          base64: action === 'import' ? data.prepared : data.original,
          sourceSlideId: sourceId,
          packageSourceSlideId: r.packageSourceSlideId,
          preimage: structuredClone(before),
        },
        async (actualId) => {
          await guard(c)
          r = await store(c, { ...r, pending: { ...r.pending!, insertedSlideId: actualId } })
        },
        () => guard(c),
        () => syncGuard(c),
        c.signal,
      )
    } else
      await options.adapter.remove(
        { slideId: sourceId, preimage: structuredClone(before) },
        () => guard(c),
        () => syncGuard(c),
        c.signal,
      )
    await guard(c)
    const actual = await inspect(c, r.pending!.insertedSlideId ? [r.pending!.insertedSlideId] : [])
    if (!(await accepted(c, r, s, before, action, actual, r.pending!.insertedSlideId)))
      throw Error('presentation_package_unknown')
    const after = proof(actual),
      afterProofRef = await save(c, `receipt-${r.receipts.length + 1}`, json(after))
    if (!same(proof(await inspect(c)), after)) throw Error('presentation_package_unknown')
    r = await store(c, { ...r, pending: { ...r.pending!, afterProofRef } })
    return finish(c, r)
  }
  const drive = async (
    c: Context,
    r: PresentationPackageChange,
    s: Snapshot,
    action: 'resume' | 'undo' | 'discard',
  ) => {
    if (r.pending) throw Error('presentation_package_unknown')
    if (action === 'discard') {
      if (r.state === 'prepared') return store(c, { ...r, state: 'discarded' })
      if (r.state !== 'staged') throw Error('presentation_package_state_invalid')
      return perform(c, r, s, 'discard')
    }
    if (action === 'undo') {
      if (r.state === 'applied') r = await perform(c, r, s, 'restore')
      if (r.state === 'restore_staged') r = await perform(c, r, s, 'delete_applied')
      return r
    }
    if (r.state === 'prepared') r = await perform(c, r, s, 'import')
    if (r.state === 'staged') r = await perform(c, r, s, 'delete_source')
    return r
  }
  const summary = (r: PresentationPackageChange) => ({
    changeId: r.changeId,
    state: r.state,
    pending: r.pending?.action,
    sourceSlideId: r.sourceSlideId,
    replacementSlideId: r.replacementSlideId,
    restoredSlideId: r.restoredSlideId,
    qaPassed: false,
  })
  const targets = (r: PresentationPackageChange) => [
    ...new Set([
      ...(r.state === 'prepared' || r.state === 'staged' || r.state === 'discarded'
        ? [r.sourceSlideId]
        : []),
      ...(r.replacementSlideId && r.state !== 'undone' && r.state !== 'discarded'
        ? [r.replacementSlideId]
        : []),
      ...(r.restoredSlideId ? [r.restoredSlideId] : []),
    ]),
  ]
  const mutationProposal = async (
    r: PresentationPackageChange,
    s: Snapshot,
    action: 'resume' | 'undo' | 'discard' | 'reconcile',
    signal?: AbortSignal,
  ) => {
    const c = context(r, signal),
      observed = await observePending(c, r, s)
    if (
      action === 'reconcile'
        ? observed.status === 'ready' || observed.status === 'unknown'
        : observed.status !== 'ready'
    )
      throw Error('presentation_package_unknown')
    if (
      (action === 'resume' && !['prepared', 'staged'].includes(r.state)) ||
      (action === 'undo' && !['applied', 'restore_staged'].includes(r.state)) ||
      (action === 'discard' && !['prepared', 'staged'].includes(r.state))
    )
      throw Error('presentation_package_state_invalid')
    return options.proposals.propose({
      operation: `${action}_package_xml_change`,
      toolName: `${action}_package_xml_change`,
      title: `${action} XML package change`,
      preview: {
        ...summary(r),
        qaScope: {
          basis: 'package_xml_savepoint',
          hostSlideIds: targets(r),
        },
      },
      impact: {
        host:
          action === 'reconcile' || (action === 'discard' && r.state === 'prepared')
            ? 'local_checkpoint'
            : 'powerpoint',
        targets: targets(r),
        count: 1,
      },
      fingerprint: selectionFingerprint(JSON.stringify(r)),
      validate: async (signal) => {
        try {
          return same(await observePending({ ...c, signal }, r, s), observed)
        } catch {
          return false
        }
      },
      execute: async (signal) => {
        const ctx = { ...c, signal }
        await guard(ctx)
        if (!same(await observePending(ctx, r, s), observed)) throw Error('proposal_stale')
        if (action !== 'reconcile') {
          await drive(ctx, r, s, action)
          return
        }
        if (observed.status === 'before') {
          await store(ctx, {
            ...r,
            pending: undefined,
            state: r.pending!.action === 'import' ? 'discarded' : r.state,
          })
          return
        }
        let next = r
        if (observed.inserted && !next.pending!.insertedSlideId)
          next = await store(ctx, {
            ...next,
            pending: { ...next.pending!, insertedSlideId: observed.inserted },
          })
        if (!next.pending!.afterProofRef) {
          const afterProofRef = await save(
            ctx,
            `receipt-${next.receipts.length + 1}`,
            json(observed.actual),
          )
          if (!same(proof(await inspect(ctx)), observed.actual)) throw Error('proposal_stale')
          next = await store(ctx, { ...next, pending: { ...next.pending!, afterProofRef } })
        }
        await finish(ctx, next)
      },
      verify: async (signal) => {
        const current = record(r.changeId)
        await check(context(current, signal), current)
      },
    })
  }
  return {
    id: 'powerpoint-package-editing',
    systemPrompt:
      'Inspect persistent XML package changes before recovery. Unknown native acknowledgements require explicit reconciliation. Never replay unresolved imports/deletes or claim screenshot review certifies the deck.',
    tools,
    clear() {
      epoch++
      invalidateVisual()
    },
    beginMutation: invalidateVisual,
    endMutation: invalidateVisual,
    async propose(
      kind: 'slide' | 'chart',
      slideIndex: number,
      replacements: XmlReplacement[],
      explanation?: string,
      signal?: AbortSignal,
    ) {
      const owned = structuredClone(replacements)
      if (
        !['slide', 'chart'].includes(kind) ||
        !Number.isSafeInteger(slideIndex) ||
        slideIndex < 0 ||
        slideIndex > 100000 ||
        owned.length < 1 ||
        owned.length > 32 ||
        json({ version: 1, operations: owned.map((r) => ({ op: 'replace_xml', ...r })) }).length >
          32 * 1024
      )
        throw Error('invalid_tool_input')
      const documentId = await options.documentId(),
        c: Context = {
          documentId,
          changeId: crypto.randomUUID(),
          epoch,
          signal,
          originalSignal: signal,
        }
      await guard(c)
      const all = await inspect(c)
      const sourceSlideId = all.slideIds[slideIndex]
      if (!sourceSlideId) throw Error('invalid_tool_input')
      const exported = await inspect(c, [sourceSlideId])
      if (!same(proof(exported), proof(all))) throw Error('proposal_stale')
      const originalBase64 = exported.pages[slideIndex]!.base64!
      const packageSourceSlideId = await awaited(c, () =>
        readPackageSourceSlideId(originalBase64, signal),
      )
      const edited = await awaited(c, () =>
        editPowerPointPackage(originalBase64, kind, owned, signal),
      )
      const restored = await awaited(c, () =>
        editPowerPointPackage(
          edited.base64,
          kind,
          edited.changedPaths.map((path) => ({ path, xml: edited.beforeXml[path]! })),
          signal,
        ),
      )
      const { base64: preparedBase64, ...expected } = edited,
        { base64: restoredBase64, ...restoreExpected } = restored
      if (
        (await awaited(c, () => presentationPackageDigest(restoredBase64, signal))) !==
        all.pages[slideIndex]!.digest
      )
        throw Error('presentation_package_backup_invalid')
      const snapshot: Snapshot = {
        version: 1,
        documentId,
        changeId: c.changeId,
        sourceKind: kind,
        sourceSlideId,
        packageSourceSlideId,
        original: proof(all),
        expected,
        restoreExpected,
      }
      if (!same(proof(await inspect(c)), proof(all))) throw Error('proposal_stale')
      const intent = explanation ?? `Edit slide ${kind} XML`
      return options.proposals.propose({
        operation: kind === 'chart' ? 'edit_slide_chart' : 'edit_slide_xml',
        toolName: kind === 'chart' ? 'edit_slide_chart' : 'edit_slide_xml',
        title: intent,
        preview: {
          changeId: c.changeId,
          state: 'prepared',
          sourceSlideId,
          qaPassed: false,
          changedPaths: edited.changedPaths,
          qaScope: { basis: 'package_xml_savepoint', hostSlideIds: [sourceSlideId] },
        },
        impact: { host: 'powerpoint', targets: [sourceSlideId], count: 1 },
        fingerprint: selectionFingerprint(JSON.stringify(proof(all))),
        validate: async (signal) => {
          try {
            const ctx = { ...c, signal }
            return same(proof(await inspect(ctx)), proof(all))
          } catch (error) {
            if (signal?.aborted || c.originalSignal?.aborted || c.epoch !== epoch) throw error
            return false
          }
        },
        execute: async (signal) => {
          const ctx = { ...c, signal }
          if (!same(proof(await inspect(ctx)), proof(all))) throw Error('proposal_stale')
          const blobs = [
            { key: 'page-0', bytes: decode(originalBase64) },
            { key: 'page-1', bytes: decode(preparedBase64) },
            { key: 'snapshot', bytes: json(snapshot) },
            { key: 'receipt-0', bytes: json(proof(all)) },
          ]
          const refs = await Promise.all(
            blobs.map(({ key, bytes }) => packageBackupRefForBytes(key, bytes)),
          )
          let persistenceAttempted = false
          try {
            for (let i = 0; i < blobs.length; i++)
              if (!same(await save(ctx, blobs[i]!.key, blobs[i]!.bytes), refs[i]))
                throw Error('presentation_package_backup_invalid')
            const [originalRef, preparedRef, snapshotRef, currentProofRef] = refs as [
              PackageBackupRef,
              PackageBackupRef,
              PackageBackupRef,
              PackageBackupRef,
            ]
            const r: PresentationPackageChange = {
              version: 1,
              kind: 'package_xml',
              documentId,
              changeId: c.changeId,
              intent,
              sourceKind: kind,
              sourceSlideId,
              packageSourceSlideId,
              snapshotRef,
              originalRef,
              preparedRef,
              currentProofRef,
              state: 'prepared',
              receipts: [],
              reviews: [],
            }
            if (!validatePresentationPackageChange(r))
              throw Error('presentation_package_state_invalid')
            await backups(ctx, r, snapshot)
            if (!same(proof(await inspect(ctx)), proof(all))) throw Error('proposal_stale')
            persistenceAttempted = true
            await store(ctx, r)
            await drive(ctx, r, snapshot, 'resume')
          } catch (error) {
            if (!persistenceAttempted)
              await Promise.allSettled(
                refs.map((backup) =>
                  releasePackageBackup({
                    request: options.request,
                    documentId,
                    changeId: c.changeId,
                    backup,
                  }),
                ),
              )
            throw error
          }
        },
        verify: async (signal) => {
          const current = record(c.changeId)
          await check(context(current, signal), current)
        },
      })
    },
    async executeTool(inputCall: any, signal?: AbortSignal) {
      try {
        const call = structuredClone(inputCall)
        const action = actions.find((a) => call.name === `${a}_package_xml_change`)
        const fields =
          action === 'review'
            ? ['change_id', 'slide_id', 'screenshot_digest', 'status', 'notes']
            : action === 'capture'
              ? ['change_id', 'slide_id']
              : ['change_id']
        if (
          !action ||
          !call.input ||
          typeof call.input !== 'object' ||
          Array.isArray(call.input) ||
          Object.keys(call.input).length !== fields.length ||
          fields.some((k) => !Object.hasOwn(call.input, k)) ||
          typeof call.input.change_id !== 'string'
        )
          throw Error('invalid_tool_input')
        const r = record(call.input.change_id),
          c = context(r, signal),
          s = await loadSnapshot(c, r)
        if (action === 'inspect') {
          const observed = await observePending(c, r, s)
          return {
            output: JSON.stringify({
              ...summary(r),
              status: observed.status,
              observedSlideId: 'inserted' in observed ? observed.inserted : undefined,
            }),
            mutated: false,
            summary: 'Inspected XML package checkpoint',
          }
        }
        if (action === 'capture' || action === 'review') {
          const slideId = call.input.slide_id
          if (
            !validPackageHostId(slideId) ||
            ![
              r.sourceSlideId,
              r.replacementSlideId,
              r.restoredSlideId,
              r.pending?.insertedSlideId,
            ].includes(slideId) ||
            !options.adapter.screenshotSlide
          )
            throw Error('invalid_tool_input')
          const observed = await observePending(c, r, s)
          const stable =
            observed.status === 'ready' && !r.pending && ['applied', 'undone'].includes(r.state)
          const captureEpoch = visualEpoch
          const before = proof(await inspect(c)),
            index = before.slideIds.indexOf(slideId)
          if (index < 0) throw Error('presentation_package_qa_stale')
          const screenshot = await awaited(c, async () =>
            structuredClone(await options.adapter.screenshotSlide!(index, signal)),
          )
          if (
            !validatePowerPointPageScreenshot(screenshot.base64) ||
            screenshot.slideId !== slideId ||
            screenshot.mime !== 'image/png'
          )
            throw Error('presentation_package_qa_stale')
          const digest = Array.from(
            new Uint8Array(
              await awaited(c, () => crypto.subtle.digest('SHA-256', decode(screenshot.base64))),
            ),
            (b) => b.toString(16).padStart(2, '0'),
          ).join('')
          if (!same(proof(await inspect(c)), before) || captureEpoch !== visualEpoch)
            throw Error('presentation_package_qa_stale')
          const key = r.changeId + '/' + slideId
          if (action === 'capture') {
            captures.set(key, { digest, record: JSON.stringify(r), epoch: visualEpoch })
            return {
              output: JSON.stringify({
                slideId,
                screenshotDigest: digest,
                base64: screenshot.base64,
                qaPassed: false,
                stateUncertain: !stable,
              }),
              mutated: false,
              summary: 'Captured historical XML change page',
            }
          }
          const capture = captures.get(key)
          if (
            !stable ||
            !capture ||
            capture.digest !== digest ||
            capture.digest !== call.input.screenshot_digest ||
            capture.record !== JSON.stringify(r) ||
            capture.epoch !== visualEpoch ||
            !['pass', 'fail'].includes(call.input.status) ||
            typeof call.input.notes !== 'string' ||
            call.input.notes.length > 2000
          )
            throw Error('presentation_package_qa_stale')
          const reviewRef = await save(
            c,
            `receipt-${Date.now()}`,
            json({
              slideId,
              screenshotDigest: digest,
              status: call.input.status,
              notes: call.input.notes,
              capturedAt: new Date().toISOString(),
              qaPassed: false,
            }),
          )
          await check(c, r)
          if (captureEpoch !== visualEpoch) throw Error('presentation_package_qa_stale')
          await store(c, {
            ...r,
            reviews: [...r.reviews.filter((v) => v.slideId !== slideId), { slideId, reviewRef }],
          })
          captures.delete(key)
          return {
            output: JSON.stringify({ qaPassed: false, historical: true }),
            mutated: false,
            summary: 'Recorded historical page review',
          }
        }
        const proposal = await mutationProposal(r, s, action, signal)
        return {
          output: JSON.stringify({ proposalId: proposal.id, ...summary(r) }),
          mutated: false,
          summary: 'XML change confirmation required',
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error)
        return {
          output: JSON.stringify({ error: message }),
          isError: true,
          mutated: false,
          summary: message,
        }
      }
    },
  } satisfies AgentSkill & {
    propose(
      kind: 'slide' | 'chart',
      slideIndex: number,
      replacements: XmlReplacement[],
      explanation?: string,
      signal?: AbortSignal,
    ): Promise<unknown>
    clear(): void
    beginMutation(): void
    endMutation(): void
  }
}
