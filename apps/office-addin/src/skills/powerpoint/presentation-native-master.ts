import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import {
  validatePowerPointPageScreenshot,
  type PowerPointAdapter,
  type PowerPointMasterOperation,
  type PowerPointMasterState,
} from './browser-powerpoint-adapter.js'
import {
  inverseMasterOperation,
  masterOperationKey,
  masterOperationValue,
  projectedMasterState,
  sameMasterOperationValue,
} from './presentation-master-program.js'
import {
  affectedStyleSlideIds,
  parsePowerPointStyleDependencies,
  type PowerPointStyleDependencies,
} from './presentation-style-dependencies.js'
import { saveMasterBackup, readMasterBackup } from './presentation-master-backup.js'
import {
  validatePresentationNativeMasterChange,
  validNativeMasterTransition,
  validMasterBackupRef,
  validMasterHostId,
  validStoredMasterOperation,
  type PresentationNativeMasterChange,
  type MasterBackupRef,
  type StoredMasterOperation,
} from './presentation-native-master-change.js'
import { presentationPackageDigest } from './powerpoint-package.js'
import {
  prepareMasterPackageProtection,
  verifyMasterPackageProtection,
  validMasterPackageProtectionTargets,
  type MasterPackageProtection,
} from './presentation-native-master-package-proof.js'

interface Options {
  documentId(): Promise<string>
  assertDocumentId?(expected: string): void
  available(): boolean
  adapter: PowerPointAdapter
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  proposals: StructuredProposalController
  readNativeMasterChange(id: string): PresentationNativeMasterChange | undefined
  writeNativeMasterChange(
    next: PresentationNativeMasterChange,
    expected: PresentationNativeMasterChange | undefined,
  ): Promise<void>
}
interface Snapshot {
  version: 1
  documentId: string
  changeId: string
  before: PowerPointMasterState
  operations: StoredMasterOperation[]
  inverseOperations: StoredMasterOperation[]
  beforeSlideIds: string[]
  dependencies: PowerPointStyleDependencies
  pages: Array<{
    hostSlideId: string
    slideIndex: number
    backup: MasterBackupRef
    originalPackageDigest: string
    protection: MasterPackageProtection
  }>
}
interface Proof {
  version: 1
  documentId: string
  changeId: string
  nextIndex: number
  targetValuesDigest: string
  unaffectedValuesDigest: string
  packages: Array<{ hostSlideId: string; packageDigest: string }>
}
interface Context {
  documentId: string
  changeId: string
  token: number
  signal?: AbortSignal
  originalSignal?: AbortSignal
  record?: PresentationNativeMasterChange
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const encode = (bytes: Uint8Array) =>
  btoa(Array.from(bytes, (x) => String.fromCharCode(x)).join(''))
const decode = (base64: string) => Uint8Array.from(atob(base64), (x) => x.charCodeAt(0))
const jsonBytes = (value: unknown) => new TextEncoder().encode(JSON.stringify(value))
const sha = async (bytes: Uint8Array) =>
  Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes).buffer)),
    (x) => x.toString(16).padStart(2, '0'),
  ).join('')
const canonical = (value: unknown): unknown =>
  Array.isArray(value)
    ? value.map(canonical)
    : value && typeof value === 'object'
      ? Object.fromEntries(
          Object.entries(value)
            .sort(([a], [b]) => a.localeCompare(b))
            .map(([key, child]) => [key, canonical(child)]),
        )
      : value
const id = (v: unknown): v is string => typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
const digest = (v: unknown): v is string => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v)
const object = (v: unknown): v is Record<string, unknown> =>
  !!v && typeof v === 'object' && !Array.isArray(v)
const exact = (v: unknown, fields: string[]): v is Record<string, unknown> =>
  object(v) &&
  Object.keys(v).length === fields.length &&
  fields.every((field) => Object.hasOwn(v, field))
const metadataOperation = (op: StoredMasterOperation): PowerPointMasterOperation =>
  op.op === 'set_master_background' && op.fill.type === 'picture_or_texture'
    ? {
        ...op,
        fill: { type: 'picture_or_texture', transparency: op.fill.transparency, image_base64: '' },
      }
    : (structuredClone(op) as PowerPointMasterOperation)
const nativeStateValid = (v: unknown): v is PowerPointMasterState => {
  if (!exact(v, ['masters']) || !Array.isArray(v.masters) || v.masters.length > 32) return false
  return (
    new Set(v.masters.map((m) => m?.id)).size === v.masters.length &&
    v.masters.every(
      (m) =>
        exact(m, ['id', 'name', 'background', 'themeColors', 'layouts']) &&
        validMasterHostId(m.id) &&
        typeof m.name === 'string' &&
        object(m.background) &&
        typeof m.background.type === 'string' &&
        Object.entries(m.background).every(([key, value]) =>
          [
            'type',
            'color',
            'gradientType',
            'pattern',
            'foregroundColor',
            'backgroundColor',
          ].includes(key)
            ? typeof value === 'string'
            : ['transparency', 'pictureTransparency'].includes(key) &&
              typeof value === 'number' &&
              Number.isFinite(value) &&
              value >= 0 &&
              value <= 1,
        ) &&
        object(m.themeColors) &&
        Object.values(m.themeColors).every((value) => typeof value === 'string') &&
        Array.isArray(m.layouts) &&
        m.layouts.length <= 128 &&
        new Set(m.layouts.map((l) => l?.id)).size === m.layouts.length &&
        m.layouts.every(
          (l) =>
            exact(l, [
              'id',
              'name',
              'isMasterBackgroundFollowed',
              'areBackgroundGraphicsHidden',
              'background',
            ]) &&
            validMasterHostId(l.id) &&
            typeof l.name === 'string' &&
            typeof l.isMasterBackgroundFollowed === 'boolean' &&
            typeof l.areBackgroundGraphicsHidden === 'boolean' &&
            exact(l.background, ['type']) &&
            typeof l.background.type === 'string',
        ),
    )
  )
}
const tools: AgentToolDef[] = ['inspect', 'reconcile', 'resume', 'undo'].map((action) => ({
  name: `${action}_slide_master_change`,
  description:
    action === 'inspect'
      ? 'Read durable native master state and exact package proofs without writing. Historical receipts are not current QA.'
      : action === 'reconcile'
        ? 'Propose confirmed reconciliation of a durable observed receipt, or close an exact unchanged pending operation without replay.'
        : action === 'undo'
          ? 'Propose confirmed inverse native master writes from verified durable receipts. Ambiguous writes or page drift stop.'
          : 'Propose confirmed continuation only from a verified durable checkpoint without a pending write.',
  inputSchema: {
    type: 'object',
    properties: { change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' } },
    required: ['change_id'],
    additionalProperties: false,
  },
}))
tools.push(
  ...(['capture_slide_master_page', 'record_slide_master_page_review'] as const).map((name) => ({
    name,
    description: name.startsWith('capture')
      ? 'Capture an affected native page for visual review. No QA pass is implied.'
      : 'Store a historical assessment of an exact fresh terminal page capture. Does not certify current or whole-deck QA.',
    inputSchema: {
      type: 'object',
      properties: {
        change_id: { type: 'string' },
        slide_id: { type: 'string' },
        ...(name.startsWith('record')
          ? {
              screenshot_digest: { type: 'string' },
              status: { type: 'string', enum: ['pass', 'fail'] },
              notes: { type: 'string', maxLength: 2000 },
            }
          : {}),
      },
      required: name.startsWith('record')
        ? ['change_id', 'slide_id', 'screenshot_digest', 'status', 'notes']
        : ['change_id', 'slide_id'],
      additionalProperties: false,
    },
  })),
)

export function createPresentationNativeMasterSkill(options: Options) {
  let epoch = 0,
    visualEpoch = 0
  const captures = new Map<
    string,
    {
      screenshotDigest: string
      capturedAt: string
      record: string
      visualEpoch: number
      terminal: boolean
    }
  >()
  const invalidateVisual = () => {
    visualEpoch++
    captures.clear()
  }
  const syncGuard = (ctx: Context) => {
    if (ctx.signal?.aborted || ctx.originalSignal?.aborted || ctx.token !== epoch)
      throw Error('cancelled')
    if (!options.available()) throw Error('presentation_existing_persistence_unavailable')
    options.assertDocumentId?.(ctx.documentId)
    if (!same(options.readNativeMasterChange(ctx.changeId), ctx.record))
      throw Error('presentation_native_master_stale')
  }
  const guard = async (ctx: Context) => {
    syncGuard(ctx)
    if ((await options.documentId()) !== ctx.documentId)
      throw Error('presentation_document_changed')
    syncGuard(ctx)
  }
  const awaited = async <T>(ctx: Context, task: () => Promise<T>): Promise<T> => {
    await guard(ctx)
    const value = await task()
    await guard(ctx)
    return value
  }
  const requestFor = (ctx: Context) => async (body: unknown, signal?: AbortSignal) =>
    awaited(ctx, () => options.request(structuredClone(body), signal))
  const read = async (ctx: Context, backup: MasterBackupRef) =>
    awaited(ctx, () =>
      readMasterBackup({
        request: requestFor(ctx),
        documentId: ctx.documentId,
        changeId: ctx.changeId,
        backup: structuredClone(backup),
        signal: ctx.signal,
      }),
    )
  const save = async (ctx: Context, key: string, bytes: Uint8Array) => {
    const saved = await awaited(ctx, () =>
      saveMasterBackup({
        request: requestFor(ctx),
        documentId: ctx.documentId,
        changeId: ctx.changeId,
        key,
        bytes: Uint8Array.from(bytes),
        signal: ctx.signal,
      }),
    )
    const verified = await read(ctx, saved)
    if (!same(Array.from(verified), Array.from(bytes)))
      throw Error('presentation_master_backup_invalid')
    return saved
  }
  const readJson = async (ctx: Context, ref: MasterBackupRef): Promise<unknown> => {
    const bytes = await read(ctx, ref)
    try {
      return JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes))
    } catch {
      throw Error('presentation_master_backup_invalid')
    }
  }
  const order = async (ctx: Context) => {
    if (!options.adapter.readSlideOrder) throw Error('office_api_unsupported')
    const result = await awaited(ctx, async () => [
      ...(await options.adapter.readSlideOrder!(ctx.signal)),
    ])
    if (
      !result.length ||
      !result.every(validMasterHostId) ||
      new Set(result).size !== result.length
    )
      throw Error('office_read_failed')
    return result
  }
  const dependencies = async (ctx: Context) => {
    if (!options.adapter.inspectStyleDependencies) throw Error('office_api_unsupported')
    return awaited(ctx, async () =>
      parsePowerPointStyleDependencies(
        structuredClone(await options.adapter.inspectStyleDependencies!(ctx.signal)),
      ),
    )
  }
  const state = async (ctx: Context) => {
    const value = await awaited(ctx, async () =>
      structuredClone(await options.adapter.inspectSlideMasters(ctx.signal)),
    )
    if (!nativeStateValid(value)) throw Error('office_read_failed')
    return value
  }
  const materialize = async (
    ctx: Context,
    op: StoredMasterOperation,
  ): Promise<PowerPointMasterOperation> => {
    if (op.op !== 'set_master_background' || op.fill.type !== 'picture_or_texture')
      return structuredClone(op) as PowerPointMasterOperation
    const bytes = await read(ctx, op.fill.imageRef)
    return {
      op: op.op,
      master_id: op.master_id,
      fill: { type: op.fill.type, transparency: op.fill.transparency, image_base64: encode(bytes) },
    }
  }
  const unaffected = (native: PowerPointMasterState, ops: StoredMasterOperation[]) => {
    const copy = structuredClone(native)
    // Layout fill type is derived from inheritance and can change after a master write.
    for (const master of copy.masters)
      for (const layout of master.layouts) delete (layout as Partial<typeof layout>).background
    for (const op of ops) {
      const m = copy.masters.find((m) => m.id === op.master_id)
      if (!m) throw Error('presentation_native_master_conflict')
      if (op.op === 'set_master_background') delete (m as Partial<typeof m>).background
      else if (op.op === 'set_master_theme_color') delete m.themeColors[op.theme_color]
      else {
        const l = m.layouts.find((l) => l.id === op.layout_id)
        if (!l) throw Error('presentation_native_master_conflict')
        delete (l as Partial<typeof l>).isMasterBackgroundFollowed
        delete (l as Partial<typeof l>).areBackgroundGraphicsHidden
      }
    }
    return canonical(copy)
  }
  const values = async (
    ctx: Context,
    native: PowerPointMasterState,
    ops: StoredMasterOperation[],
  ) => ({
    targetValuesDigest: await awaited(ctx, () =>
      sha(
        jsonBytes(
          canonical(
            ops.map((op) => [
              masterOperationKey(metadataOperation(op)),
              masterOperationValue(native, metadataOperation(op)),
            ]),
          ),
        ),
      ),
    ),
    unaffectedValuesDigest: await awaited(ctx, () => sha(jsonBytes(unaffected(native, ops)))),
  })
  const loadSnapshot = async (
    ctx: Context,
    r: PresentationNativeMasterChange,
  ): Promise<Snapshot> => {
    const raw = await readJson(ctx, r.snapshotRef)
    if (
      !exact(raw, [
        'version',
        'documentId',
        'changeId',
        'before',
        'operations',
        'inverseOperations',
        'beforeSlideIds',
        'dependencies',
        'pages',
      ]) ||
      raw.version !== 1 ||
      raw.documentId !== r.documentId ||
      raw.changeId !== r.changeId ||
      !nativeStateValid(raw.before) ||
      !same(raw.operations, r.operations) ||
      !same(raw.inverseOperations, r.inverseOperations) ||
      !Array.isArray(raw.beforeSlideIds) ||
      !raw.beforeSlideIds.length ||
      !raw.beforeSlideIds.every(validMasterHostId) ||
      new Set(raw.beforeSlideIds).size !== raw.beforeSlideIds.length ||
      !Array.isArray(raw.pages)
    )
      throw Error('presentation_master_backup_invalid')
    const snapshot = raw as unknown as Snapshot
    snapshot.dependencies = parsePowerPointStyleDependencies(snapshot.dependencies)
    if (
      snapshot.dependencies.slides.length !== snapshot.beforeSlideIds.length ||
      !same(
        [...snapshot.dependencies.slides.map((p) => p.slideId)].sort(),
        [...snapshot.beforeSlideIds].sort(),
      )
    )
      throw Error('presentation_master_backup_invalid')
    const affected = affectedStyleSlideIds(
      snapshot.dependencies,
      r.operations.map(metadataOperation),
    )
    if (
      new Set(snapshot.pages.map((p) => p.hostSlideId)).size !== snapshot.pages.length ||
      snapshot.pages.length !== r.scope.affectedPageCount ||
      !same([...affected].sort(), snapshot.pages.map((p) => p.hostSlideId).sort()) ||
      !snapshot.pages.every(
        (p) =>
          exact(p, [
            'hostSlideId',
            'slideIndex',
            'backup',
            'originalPackageDigest',
            'protection',
          ]) &&
          validMasterHostId(p.hostSlideId) &&
          Number.isSafeInteger(p.slideIndex) &&
          snapshot.beforeSlideIds[p.slideIndex] === p.hostSlideId &&
          validMasterBackupRef(p.backup) &&
          digest(p.originalPackageDigest) &&
          validMasterPackageProtectionTargets(
            p.protection,
            snapshot.dependencies.slides.find((d) => d.slideId === p.hostSlideId)!,
            snapshot.operations,
          ),
      )
    )
      throw Error('presentation_master_backup_invalid')
    const computedInverse = r.operations.map((op) =>
      inverseMasterOperation(snapshot.before, metadataOperation(op)),
    )
    if (!same(computedInverse, r.inverseOperations))
      throw Error('presentation_master_backup_invalid')
    return snapshot
  }
  const loadProof = async (
    ctx: Context,
    ref: MasterBackupRef,
    snapshot: Snapshot,
  ): Promise<Proof> => {
    const raw = await readJson(ctx, ref)
    if (
      !exact(raw, [
        'version',
        'documentId',
        'changeId',
        'nextIndex',
        'targetValuesDigest',
        'unaffectedValuesDigest',
        'packages',
      ]) ||
      raw.version !== 1 ||
      raw.documentId !== ctx.documentId ||
      raw.changeId !== ctx.changeId ||
      !Number.isSafeInteger(raw.nextIndex) ||
      Number(raw.nextIndex) < 0 ||
      Number(raw.nextIndex) > snapshot.operations.length ||
      !digest(raw.targetValuesDigest) ||
      !digest(raw.unaffectedValuesDigest) ||
      !Array.isArray(raw.packages) ||
      raw.packages.length !== snapshot.pages.length ||
      !raw.packages.every(
        (p, i) =>
          exact(p, ['hostSlideId', 'packageDigest']) &&
          p.hostSlideId === snapshot.pages[i]!.hostSlideId &&
          digest(p.packageDigest),
      )
    )
      throw Error('presentation_master_backup_invalid')
    return raw as unknown as Proof
  }
  const exportPage = async (ctx: Context, p: Snapshot['pages'][number]) => {
    const page = await awaited(ctx, async () =>
      structuredClone(await options.adapter.exportSlidePackage(p.slideIndex, ctx.signal)),
    )
    if (page.slideId !== p.hostSlideId) throw Error('presentation_native_master_conflict')
    return page.base64
  }
  const stableStructure = async (ctx: Context, snapshot: Snapshot) => {
    if (
      !same(await order(ctx), snapshot.beforeSlideIds) ||
      !same(await dependencies(ctx), snapshot.dependencies)
    )
      throw Error('presentation_native_master_conflict')
  }
  const observe = async (
    ctx: Context,
    snapshot: Snapshot,
    nextIndex: number,
    protectedCheck = true,
  ): Promise<Proof> => {
    await stableStructure(ctx, snapshot)
    const before = await state(ctx)
    const firstValues = await values(ctx, before, snapshot.operations)
    const packages: Proof['packages'] = []
    for (const p of snapshot.pages) {
      const base64 = await exportPage(ctx, p)
      if (protectedCheck) {
        const dependency = snapshot.dependencies.slides.find((d) => d.slideId === p.hostSlideId)!
        if (
          !(await awaited(ctx, () =>
            verifyMasterPackageProtection(
              base64,
              p.protection,
              dependency,
              snapshot.operations,
              snapshot.operations.slice(0, nextIndex),
              ctx.signal,
              () => guard(ctx),
            ),
          ))
        )
          throw Error('presentation_native_master_package_unproven')
      }
      packages.push({
        hostSlideId: p.hostSlideId,
        packageDigest: await awaited(ctx, () => presentationPackageDigest(base64, ctx.signal)),
      })
    }
    await stableStructure(ctx, snapshot)
    const lastValues = await values(ctx, await state(ctx), snapshot.operations)
    if (!same(firstValues, lastValues)) throw Error('presentation_native_master_conflict')
    return {
      version: 1,
      documentId: ctx.documentId,
      changeId: ctx.changeId,
      nextIndex,
      ...lastValues,
      packages,
    }
  }
  const assertExpected = async (ctx: Context, snapshot: Snapshot, proof: Proof) => {
    const expected = projectedMasterState(
      snapshot.before,
      snapshot.operations.slice(0, proof.nextIndex).map(metadataOperation),
    )
    const expectedValues = await values(ctx, expected, snapshot.operations)
    if (
      proof.targetValuesDigest !== expectedValues.targetValuesDigest ||
      proof.unaffectedValuesDigest !== expectedValues.unaffectedValuesDigest
    )
      throw Error('presentation_native_master_conflict')
  }
  const check = async (ctx: Context, r: PresentationNativeMasterChange, snapshot: Snapshot) => {
    const proof = await loadProof(ctx, r.currentProofRef, snapshot)
    if (proof.nextIndex !== r.nextIndex) throw Error('presentation_master_backup_invalid')
    await assertExpected(ctx, snapshot, proof)
    // Compare fresh SHA-256 of every package entry byte to the immutable previously verified proof.
    if (!same(await observe(ctx, snapshot, r.nextIndex, false), proof))
      throw Error('presentation_native_master_conflict')
    return proof
  }
  const backups = async (ctx: Context, snapshot: Snapshot) => {
    for (const p of snapshot.pages) {
      const original = await read(ctx, p.backup)
      if (
        (await awaited(ctx, () => presentationPackageDigest(encode(original), ctx.signal))) !==
        p.originalPackageDigest
      )
        throw Error('presentation_master_backup_invalid')
    }
    for (const op of snapshot.operations)
      if (op.op === 'set_master_background' && op.fill.type === 'picture_or_texture')
        await read(ctx, op.fill.imageRef)
  }
  const store = async (ctx: Context, next: PresentationNativeMasterChange) => {
    const owned = structuredClone(next)
    if (!validNativeMasterTransition(ctx.record, owned))
      throw Error('presentation_native_master_state_invalid')
    await guard(ctx)
    await options.writeNativeMasterChange(
      structuredClone(owned),
      ctx.record && structuredClone(ctx.record),
    )
    ctx.record = owned
    await guard(ctx)
    return owned
  }
  const finishPending = async (ctx: Context, r: PresentationNativeMasterChange) => {
    const pending = r.pending
    if (!pending?.afterProofRef) throw Error('presentation_native_master_write_uncertain')
    const { pending: _pending, ...rest } = r
    const nextIndex = r.nextIndex + (pending.direction === 'forward' ? 1 : -1)
    return store(ctx, {
      ...rest,
      nextIndex,
      currentProofRef: pending.afterProofRef,
      state:
        pending.direction === 'forward'
          ? nextIndex === r.operations.length
            ? 'applied'
            : 'applying'
          : nextIndex === 0
            ? 'undone'
            : 'undoing',
      receipts: [
        ...r.receipts,
        { direction: pending.direction, index: pending.index, proofRef: pending.afterProofRef },
      ],
      reviews: [],
    })
  }
  const drive = async (ctx: Context, r: PresentationNativeMasterChange, snapshot: Snapshot) => {
    if (r.pending) throw Error('presentation_native_master_write_uncertain')
    await backups(ctx, snapshot)
    while (r.state === 'applying' || r.state === 'undoing') {
      await check(ctx, r, snapshot)
      const direction = r.state === 'undoing' ? 'undo' : 'forward'
      const index = direction === 'forward' ? r.nextIndex : r.nextIndex - 1
      const storedOperation =
        direction === 'forward' ? r.operations[index]! : r.inverseOperations[index]!
      const operation = await materialize(ctx, storedOperation)
      await check(ctx, r, snapshot)
      r = await store(ctx, {
        ...r,
        pending: { direction, index, beforeProofRef: r.currentProofRef },
        reviews: [],
      })
      await check(ctx, r, snapshot)
      // Pending is durable before every SDK write. Any exception leaves it intact.
      await guard(ctx)
      syncGuard(ctx)
      invalidateVisual()
      await options.adapter.executeMasterOperations(
        [structuredClone(operation)],
        ctx.signal,
        {
          before: projectedMasterState(
            snapshot.before,
            r.operations.slice(0, r.nextIndex).map(metadataOperation),
          ),
          operations: r.operations.map(metadataOperation),
          slideIds: [...snapshot.beforeSlideIds],
          dependencies: structuredClone(snapshot.dependencies),
        },
        () => guard(ctx),
        () => syncGuard(ctx),
      )
      await guard(ctx)
      const nextIndex = r.nextIndex + (direction === 'forward' ? 1 : -1)
      const expected = projectedMasterState(
        snapshot.before,
        r.operations.slice(0, nextIndex).map(metadataOperation),
      )
      const actual = await state(ctx)
      if (
        !r.operations.every((op) =>
          sameMasterOperationValue(actual, expected, metadataOperation(op)),
        )
      )
        throw Error('presentation_native_master_write_uncertain')
      if (
        direction === 'undo' &&
        r.operations
          .slice(nextIndex)
          .some(
            (op) =>
              op.op !== 'set_layout_background_following' &&
              !snapshot.pages.some((p) =>
                snapshot.dependencies.slides.some(
                  (d) => d.slideId === p.hostSlideId && d.masterId === op.master_id,
                ),
              ),
          )
      )
        throw Error('presentation_native_master_package_unproven')
      const proof = await observe(ctx, snapshot, nextIndex)
      await assertExpected(ctx, snapshot, proof)
      const afterProofRef = await save(ctx, `receipt-${r.receipts.length + 1}`, jsonBytes(proof))
      // The first read above verifies protected XML; this fresh read requires identical complete package bytes.
      if (!same(await observe(ctx, snapshot, nextIndex, false), proof))
        throw Error('presentation_native_master_write_uncertain')
      r = await store(ctx, { ...r, pending: { ...r.pending!, afterProofRef } })
      r = await finishPending(ctx, r)
    }
    await check(ctx, r, snapshot)
    return r
  }
  const record = (changeId: string) => {
    const saved = options.readNativeMasterChange(changeId)
    if (!saved || !validatePresentationNativeMasterChange(saved))
      throw Error('presentation_native_master_missing')
    return structuredClone(saved)
  }
  const context = (r: PresentationNativeMasterChange, signal?: AbortSignal): Context => ({
    documentId: r.documentId,
    changeId: r.changeId,
    token: epoch,
    record: r,
    signal,
  })
  const summary = (r: PresentationNativeMasterChange) => ({
    changeId: r.changeId,
    state: r.state,
    nextIndex: r.nextIndex,
    pending: r.pending && {
      direction: r.pending.direction,
      index: r.pending.index,
      observed: !!r.pending.afterProofRef,
    },
    scope: r.scope,
    qaPassed: false,
  })
  const output = (value: unknown) => {
    const encoded = JSON.stringify(value)
    if (new TextEncoder().encode(encoded).length > 256 * 1024) throw Error('office_read_failed')
    return encoded
  }
  const inspect = async (ctx: Context, r: PresentationNativeMasterChange, snapshot: Snapshot) => {
    if (!r.pending) {
      await check(ctx, r, snapshot)
      return { status: 'ready' as const }
    }
    const actual = await observe(ctx, snapshot, r.nextIndex, false)
    if (r.pending.afterProofRef) {
      const proof = await loadProof(ctx, r.pending.afterProofRef, snapshot)
      if (proof.nextIndex !== r.nextIndex + (r.pending.direction === 'forward' ? 1 : -1))
        throw Error('presentation_master_backup_invalid')
      await assertExpected(ctx, snapshot, proof)
      return {
        status: same({ ...actual, nextIndex: proof.nextIndex }, proof)
          ? ('pending_observed' as const)
          : ('unknown' as const),
      }
    }
    const before = await loadProof(ctx, r.pending.beforeProofRef, snapshot)
    if (same(actual, before)) return { status: 'pending_before' as const }
    return { status: 'unknown' as const }
  }
  const mutationProposal = async (
    r: PresentationNativeMasterChange,
    snapshot: Snapshot,
    action: 'resume' | 'undo' | 'reconcile',
    signal?: AbortSignal,
  ) => {
    const ctx = context(r, signal),
      observed = await inspect(ctx, r, snapshot)
    if (action === 'reconcile') {
      if (!r.pending || !['pending_observed', 'pending_before'].includes(observed.status))
        throw Error('presentation_native_master_write_uncertain')
    } else {
      if (r.pending) throw Error('presentation_native_master_write_uncertain')
      if (action === 'resume' && !['applying', 'undoing'].includes(r.state))
        throw Error('presentation_native_master_state_invalid')
      if (action === 'undo' && !['applying', 'applied', 'undoing'].includes(r.state))
        throw Error('presentation_native_master_state_invalid')
    }
    const toolName = `${action}_slide_master_change`
    return options.proposals.propose({
      operation: toolName,
      toolName,
      title:
        action === 'undo'
          ? 'Undo native master change'
          : action === 'resume'
            ? 'Continue native master change'
            : 'Reconcile native master receipt',
      preview: {
        ...summary(r),
        ...(action === 'reconcile'
          ? { reconcile: observed.status }
          : {
              qaScope: {
                basis: 'native_master_layout',
                hostSlideIds: snapshot.pages.map((p) => p.hostSlideId),
              },
            }),
      },
      impact: {
        host: action === 'reconcile' ? 'local_checkpoint' : 'powerpoint',
        targets: r.scope.masterIds.map((id) => `master:${id}`),
        count: r.scope.masterIds.length,
      },
      fingerprint: selectionFingerprint(JSON.stringify(r)),
      validate: async (s) => {
        const c = { ...ctx, signal: s }
        try {
          return (await inspect(c, r, snapshot)).status === observed.status
        } catch {
          if (s?.aborted || ctx.token !== epoch) throw Error('cancelled')
          return false
        }
      },
      execute: async (s) => {
        const c = { ...ctx, signal: s }
        await guard(c)
        if ((await inspect(c, r, snapshot)).status !== observed.status)
          throw Error('proposal_stale')
        if (action === 'reconcile') {
          if (observed.status === 'pending_observed') await finishPending(c, r)
          else {
            const { pending: _pending, ...rest } = r
            await store(c, {
              ...rest,
              state: r.nextIndex === 0 ? 'undone' : 'undoing',
              reviews: [],
            })
          }
        } else {
          let next = r
          if (action === 'undo' && r.state !== 'undoing')
            next = await store(c, {
              ...r,
              state: r.nextIndex === 0 ? 'undone' : 'undoing',
              reviews: [],
            })
          await drive(c, next, snapshot)
        }
      },
      verify: async (s) => {
        const saved = record(r.changeId),
          c = context(saved, s)
        if (c.token !== ctx.token) throw Error('cancelled')
        await check(c, saved, snapshot)
      },
    })
  }
  return {
    id: 'presentation-native-master',
    systemPrompt:
      'Native master writes use a standalone durable PC savepoint with complete dependency scope. Inspect ambiguous writes before recovery. Never replay unknown host writes or auto-undo them. QA captures and reviews are historical evidence, not current or whole-deck certification.',
    tools,
    clear() {
      epoch++
      invalidateVisual()
    },
    beginMutation: invalidateVisual,
    endMutation: invalidateVisual,
    async propose(input: PowerPointMasterOperation[], explanation?: string, signal?: AbortSignal) {
      const token = epoch
      if (signal?.aborted) throw Error('cancelled')
      if (!options.available()) throw Error('presentation_existing_persistence_unavailable')
      const operations = structuredClone(input)
      if (
        !Array.isArray(operations) ||
        !operations.length ||
        operations.length > 32 ||
        new Set(operations.map(masterOperationKey)).size !== operations.length ||
        (explanation !== undefined &&
          (typeof explanation !== 'string' || !explanation.trim() || explanation.length > 300))
      )
        throw Error('invalid_tool_input')
      const documentId = await options.documentId()
      const ctx: Context = {
        documentId,
        changeId: crypto.randomUUID(),
        token,
        signal,
        originalSignal: signal,
      }
      await guard(ctx)
      const before = await state(ctx),
        beforeSlideIds = await order(ctx),
        deps = await dependencies(ctx)
      if (
        deps.slides.length !== beforeSlideIds.length ||
        !same([...deps.slides.map((p) => p.slideId)].sort(), [...beforeSlideIds].sort())
      )
        throw Error('office_read_failed')
      const inverses = operations.map((op) => inverseMasterOperation(before, op))
      const stored: StoredMasterOperation[] = []
      for (const [i, op] of operations.entries()) {
        if (op.op === 'set_master_background' && op.fill.type === 'picture_or_texture') {
          const bytes = decode(op.fill.image_base64)
          if (
            !bytes.length ||
            bytes.length > 8 * 1024 * 1024 ||
            encode(bytes) !== op.fill.image_base64
          )
            throw Error('invalid_tool_input')
          const imageRef = await save(ctx, `image-${i}`, bytes)
          stored.push({
            op: op.op,
            master_id: op.master_id,
            fill: { type: op.fill.type, transparency: op.fill.transparency, imageRef },
          })
        } else stored.push(structuredClone(op) as StoredMasterOperation)
      }
      if (!stored.every(validStoredMasterOperation) || !inverses.every(validStoredMasterOperation))
        throw Error('invalid_tool_input')
      const storedInverses = inverses as StoredMasterOperation[]
      const snapshot: Snapshot = {
        version: 1,
        documentId,
        changeId: ctx.changeId,
        before,
        operations: stored,
        inverseOperations: storedInverses,
        beforeSlideIds,
        dependencies: deps,
        pages: [],
      }
      if (
        stored.some(
          (op) =>
            op.op !== 'set_layout_background_following' &&
            !deps.slides.some((d) => d.masterId === op.master_id),
        )
      )
        throw Error('presentation_native_master_inverse_unproven')
      const affected = affectedStyleSlideIds(deps, operations)
      for (const [i, hostSlideId] of affected.entries()) {
        const p = {
          hostSlideId,
          slideIndex: beforeSlideIds.indexOf(hostSlideId),
          backup: undefined as unknown as MasterBackupRef,
          originalPackageDigest: '',
          protection: undefined as unknown as MasterPackageProtection,
        }
        const base64 = await exportPage(ctx, p)
        p.originalPackageDigest = await awaited(ctx, () =>
          presentationPackageDigest(base64, signal),
        )
        const dependency = deps.slides.find((d) => d.slideId === hostSlideId)!
        p.protection = await awaited(ctx, () =>
          prepareMasterPackageProtection(
            base64,
            dependency,
            stored,
            signal,
            () => guard(ctx),
            storedInverses,
          ),
        )
        p.backup = await save(ctx, `page-${i}`, decode(base64))
        snapshot.pages.push(p)
      }
      const proof = await observe(ctx, snapshot, 0)
      await assertExpected(ctx, snapshot, proof)
      if (
        !snapshot.pages.every(
          (p, i) => p.originalPackageDigest === proof.packages[i]!.packageDigest,
        )
      )
        throw Error('presentation_native_master_conflict')
      const snapshotRef = await save(ctx, 'snapshot', jsonBytes(snapshot)),
        currentProofRef = await save(ctx, 'receipt-0', jsonBytes(proof))
      const r: PresentationNativeMasterChange = {
        version: 1,
        kind: 'native_master',
        changeId: ctx.changeId,
        documentId,
        intent: explanation ?? 'Edit native PowerPoint master',
        snapshotRef,
        operations: stored,
        inverseOperations: storedInverses,
        scope: {
          masterIds: [...new Set(stored.map((op) => op.master_id))],
          affectedPageCount: snapshot.pages.length,
        },
        nextIndex: 0,
        state: 'applying',
        currentProofRef,
        receipts: [],
        reviews: [],
      }
      if (!validatePresentationNativeMasterChange(r))
        throw Error('presentation_native_master_state_invalid')
      await backups(ctx, snapshot)
      if (!same(await observe(ctx, snapshot, 0), proof))
        throw Error('presentation_native_master_conflict')
      return options.proposals.propose({
        operation: 'edit_slide_master',
        toolName: 'edit_slide_master',
        title: r.intent,
        preview: {
          ...summary(r),
          operations: r.operations,
          qaScope: { basis: 'native_master_layout', hostSlideIds: affected },
        },
        impact: {
          host: 'powerpoint',
          targets: r.scope.masterIds.map((id) => `master:${id}`),
          count: r.scope.masterIds.length,
        },
        fingerprint: selectionFingerprint(JSON.stringify(r)),
        validate: async (s) => {
          const c = { ...ctx, signal: s }
          try {
            return same(await observe(c, snapshot, 0), proof)
          } catch {
            if (s?.aborted || signal?.aborted || token !== epoch) throw Error('cancelled')
            return false
          }
        },
        execute: async (s) => {
          const c = { ...ctx, signal: s }
          await backups(c, snapshot)
          if (!same(await observe(c, snapshot, 0), proof)) throw Error('proposal_stale')
          const saved = await store(c, r)
          await drive(c, saved, snapshot)
        },
        verify: async (s) => {
          const saved = record(r.changeId)
          const c = context(saved, s)
          if (token !== epoch) throw Error('cancelled')
          await check(c, saved, snapshot)
        },
      })
    },
    async executeTool(call, signal?: AbortSignal) {
      if (!tools.some((tool) => tool.name === call.name))
        return { output: 'Unknown tool', isError: true, mutated: false, summary: call.name }
      try {
        const input = call.input
        const fields =
          call.name === 'record_slide_master_page_review'
            ? ['change_id', 'slide_id', 'screenshot_digest', 'status', 'notes']
            : call.name === 'capture_slide_master_page'
              ? ['change_id', 'slide_id']
              : ['change_id']
        if (!exact(input, fields) || !id(input.change_id)) throw Error('invalid_tool_input')
        const r = record(input.change_id),
          ctx = context(r, signal)
        await guard(ctx)
        const snapshot = await loadSnapshot(ctx, r)
        if (call.name === 'inspect_slide_master_change') {
          const observed = await inspect(ctx, r, snapshot)
          return {
            output: output({
              ...summary(r),
              ...observed,
              affectedSlideIds: snapshot.pages.map((p) => p.hostSlideId),
            }),
            mutated: false,
            summary: 'Inspected native master checkpoint',
          }
        }
        if (
          call.name === 'capture_slide_master_page' ||
          call.name === 'record_slide_master_page_review'
        ) {
          if (!validMasterHostId(input.slide_id)) throw Error('invalid_tool_input')
          const p = snapshot.pages.find((p) => p.hostSlideId === input.slide_id)
          if (!p) throw Error('invalid_tool_input')
          const status = await inspect(ctx, r, snapshot),
            terminal =
              !r.pending && ['applied', 'undone'].includes(r.state) && status.status === 'ready'
          const captureEpoch = visualEpoch,
            key = `${r.changeId}/${p.hostSlideId}`
          if (
            call.name === 'record_slide_master_page_review' &&
            (!terminal ||
              !digest(input.screenshot_digest) ||
              !['pass', 'fail'].includes(input.status as string) ||
              typeof input.notes !== 'string' ||
              input.notes.length > 2000)
          )
            throw Error('presentation_native_master_qa_stale')
          const screenshot = await awaited(ctx, async () =>
            structuredClone(await options.adapter.screenshotSlide(p.slideIndex, signal)),
          )
          if (screenshot.slideId !== p.hostSlideId || screenshot.mime !== 'image/png')
            throw Error('office_read_failed')
          validatePowerPointPageScreenshot(screenshot.base64)
          const screenshotDigest = await awaited(ctx, () => sha(decode(screenshot.base64)))
          if (
            (await inspect(ctx, r, snapshot)).status !== status.status ||
            captureEpoch !== visualEpoch
          )
            throw Error('presentation_native_master_qa_stale')
          if (call.name === 'capture_slide_master_page') {
            const capturedAt = new Date().toISOString()
            captures.set(key, {
              screenshotDigest,
              capturedAt,
              record: JSON.stringify(r),
              visualEpoch: captureEpoch,
              terminal,
            })
            return {
              output: output({
                ...summary(r),
                slideId: p.hostSlideId,
                screenshotDigest,
                capturedAt,
                stateUncertain: !terminal,
              }),
              mutated: false,
              summary: 'Captured native master affected page',
              modelContent: [
                { type: 'image', image: { mime: 'image/png', base64: screenshot.base64 } },
              ],
              display: {
                kind: 'images',
                items: [{ url: `data:image/png;base64,${screenshot.base64}` }],
              },
            }
          }
          const capture = captures.get(key)
          if (
            !capture ||
            !capture.terminal ||
            capture.visualEpoch !== visualEpoch ||
            capture.record !== JSON.stringify(r) ||
            capture.screenshotDigest !== screenshotDigest ||
            input.screenshot_digest !== screenshotDigest
          )
            throw Error('presentation_native_master_qa_stale')
          const reviewRef = await save(
            ctx,
            `receipt-${Date.now()}${Math.floor(Math.random() * 100000)}`,
            jsonBytes({
              version: 1,
              documentId: r.documentId,
              changeId: r.changeId,
              slideId: p.hostSlideId,
              screenshotDigest,
              capturedAt: capture.capturedAt,
              reviewedAt: new Date().toISOString(),
              status: input.status,
              notes: input.notes,
            }),
          )
          await check(ctx, r, snapshot)
          const next = {
            ...r,
            reviews: [
              ...r.reviews.filter((v) => v.slideId !== p.hostSlideId),
              { slideId: p.hostSlideId, reviewRef },
            ],
          }
          await store(ctx, next)
          captures.delete(key)
          return {
            output: output({ ...summary(next), historicalReview: true }),
            mutated: false,
            summary: 'Recorded historical native master page review',
          }
        }
        const action = call.name.split('_')[0] as 'resume' | 'undo' | 'reconcile'
        return {
          output: output(await mutationProposal(r, snapshot, action, signal)),
          mutated: false,
          summary: 'Proposed native master recovery',
        }
      } catch (error) {
        const message =
          error instanceof Error ? error.message : 'presentation_native_master_operation_failed'
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
      input: PowerPointMasterOperation[],
      explanation?: string,
      signal?: AbortSignal,
    ): Promise<ReturnType<StructuredProposalController['propose']>>
    clear(): void
    beginMutation(): void
    endMutation(): void
  }
}
