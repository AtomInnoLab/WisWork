import type { AgentSkill, AgentToolDef, ToolExecutionOutcome } from '@wiswork/agent-core'
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
  presentationPackageDigest,
  type XmlReplacement,
  type PackageEditResult,
} from './powerpoint-package.js'
import {
  inspectMasterXmlPackage,
  deriveMasterXmlCarrier,
  assertMasterXmlPreparation,
  proveMasterXmlLayoutMapping,
  assertMasterXmlPagePreserved,
  type MasterXmlPackageInventory,
  type MasterXmlLayoutMapping,
} from './presentation-master-xml-package.js'
import type {
  BrowserPresentationMasterXmlAdapter,
  MasterXmlHostSnapshot,
} from './browser-presentation-master-xml-adapter.js'
import {
  validatePresentationMasterXmlChange,
  validMasterXmlTransition,
  type PresentationMasterXmlChange,
  type MasterXmlNativeSource,
  type MasterXmlAction,
  type MasterXmlState,
} from './presentation-master-xml-change.js'
import { validPackageHostId, validPackageBackupRef } from './presentation-package-change.js'
import {
  validatePowerPointPageScreenshot,
  type PowerPointAdapter,
} from './browser-powerpoint-adapter.js'
export interface PresentationMasterXmlOptions {
  documentId(): Promise<string>
  assertDocumentId?(expected: string): void
  available(): boolean
  adapter: Pick<
    BrowserPresentationMasterXmlAdapter,
    'inspect' | 'stage' | 'remove' | 'applyLayout' | 'readPage'
  > &
    Partial<Pick<PowerPointAdapter, 'screenshotSlide'>>
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  proposals: StructuredProposalController
  readMasterXmlChange(id: string): PresentationMasterXmlChange | undefined
  writeMasterXmlChange(
    next: PresentationMasterXmlChange,
    expected: PresentationMasterXmlChange | undefined,
  ): Promise<void>
}
type CommandKind = 'stage' | 'layout' | 'remove' | 'checkpoint'
type Host = MasterXmlHostSnapshot
type OriginalPage = {
  slideId: string
  packageRef: PackageBackupRef
  masterId: string
  layoutId: string
}
interface Snapshot {
  version: 1
  documentId: string
  changeId: string
  sourceSlideId: string
  packageSourceSlideId: string
  original: Host
  pages: OriginalPage[]
  carrierRef: PackageBackupRef
  preparedRef: PackageBackupRef
  originalInventory: MasterXmlPackageInventory
  preparedInventory: MasterXmlPackageInventory
  affectedMasterPaths: string[]
  replacements: XmlReplacement[]
  expected: Omit<PackageEditResult, 'base64'>
}
interface Progress {
  masterCarriers?: {
    mode: 'imported' | 'restored'
    packageMasterPath: string
    identity: MasterXmlNativeSource
    packageRef: PackageBackupRef
  }[]
  importedMasterIds: string[]
  restoredMasterIds: string[]
  introducedMasterIds: string[]
  originalRepresentatives: { masterId: string; layoutId: string; packageRef: PackageBackupRef }[]
  originalMappings: { packageMasterPath: string; mapping: MasterXmlLayoutMapping }[]
  importedRepresentatives: { masterId: string; layoutId: string; packageRef: PackageBackupRef }[]
  importedMappings: { packageMasterPath: string; mapping: MasterXmlLayoutMapping }[]
  restoredRepresentatives: { masterId: string; layoutId: string; packageRef: PackageBackupRef }[]
  restoredMappings: { packageMasterPath: string; mapping: MasterXmlLayoutMapping }[]
  roles: Partial<
    Record<
      | 'stagedSource'
      | 'restoredSource'
      | 'originalProbeSource'
      | 'importedProbeSource'
      | 'restoredProbeSource',
      MasterXmlNativeSource
    >
  >
  ownedSources: MasterXmlNativeSource[]
  mode: 'forward' | 'undo' | 'discard'
  currentBackups?: {
    slideId: string
    originalSlideId?: string
    packageRef: PackageBackupRef
    digest: string
  }[]
  fallbackOriginalIds: string[]
  restorationStaged?: { originalSlideId: string; identity: MasterXmlNativeSource }
  pageMappings: { originalSlideId: string; currentSlideId: string }[]
}
interface Proof {
  version: 1
  documentId: string
  changeId: string
  sequence: number
  previous?: PackageBackupRef
  host: Host
  progress: Progress
  action?: MasterXmlAction
  next?: { state: MasterXmlState; phase: MasterXmlAction; index: number }
  step?: {
    kind: CommandKind
    index: number
    targetSlideId?: string
    masterId?: string
    layoutId?: string
    sourceId?: string
    originalId?: string
    inserted?: MasterXmlNativeSource
    beforeBackupRef?: PackageBackupRef
    recoveryPreimageRef?: PackageBackupRef
  }
}
interface Context {
  documentId: string
  changeId: string
  epoch: number
  record?: PresentationMasterXmlChange
  signal?: AbortSignal
  originalSignal?: AbortSignal
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const json = (value: unknown) => new TextEncoder().encode(JSON.stringify(value))
const decode = (base64: string) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
const encode = (bytes: Uint8Array) =>
  btoa(Array.from(bytes, (b) => String.fromCharCode(b)).join(''))
const host = (v: Host): Host => ({
  slideIds: [...v.slideIds],
  pages: v.pages.map(({ slideId, digest }) => ({ slideId, digest })),
  masters: structuredClone(v.masters),
  dependencies: structuredClone(v.dependencies),
})
function fail(suffix: string): never {
  throw Error('presentation_master_xml_' + suffix)
}
const SAFE_ERRORS = new Set([
  'cancelled',
  'aborted',
  'invalid_tool_input',
  'office_api_unsupported',
  'office_read_failed',
  'office_verify_failed',
  'proposal_stale',
  'proposal_missing',
  'proposal_confirmation_in_progress',
  'selection_too_large',
  'presentation_document_changed',
  'presentation_document_identity_unavailable',
  'presentation_change_history_full',
  'presentation_change_history_pending',
  'presentation_change_history_state_invalid',
  'presentation_lock_review_pending',
  'presentation_lock_review_unavailable',
  'presentation_package_backup_invalid',
  'presentation_package_backup_capacity',
  ...[
    'backup_budget_exceeded',
    'backup_invalid',
    'drift',
    'invalid',
    'mapping_unproven',
    'persistence_unavailable',
    'qa_stale',
    'stale',
    'state_invalid',
    'unknown',
    'package_unproven',
  ].map((s) => 'presentation_master_xml_' + s),
])
const safeError = (error: unknown) => {
  try {
    const message = error instanceof Error ? error.message : undefined
    if (typeof message === 'string' && SAFE_ERRORS.has(message)) return message
  } catch {
    /* Error accessors are untrusted transport data. */
  }
  return 'presentation_master_xml_failed'
}
function validHost(v: Host): boolean {
  return (
    !!v &&
    json(v).length <= 8 * 1024 * 1024 &&
    Array.isArray(v.slideIds) &&
    v.slideIds.length > 0 &&
    v.slideIds.every(validPackageHostId) &&
    new Set(v.slideIds).size === v.slideIds.length &&
    Array.isArray(v.pages) &&
    v.pages.length === v.slideIds.length &&
    v.pages.every((p, i) => p.slideId === v.slideIds[i] && /^[a-f0-9]{64}$/.test(p.digest)) &&
    Array.isArray(v.masters) &&
    v.masters.length > 0 &&
    new Set(v.masters.map((m) => m.masterId)).size === v.masters.length &&
    v.masters.every(
      (m) =>
        validPackageHostId(m.masterId) &&
        typeof m.name === 'string' &&
        Array.isArray(m.layouts) &&
        m.layouts.length > 0 &&
        m.layouts.every((l) => validPackageHostId(l.layoutId) && typeof l.name === 'string'),
    ) &&
    new Set(v.masters.flatMap((m) => m.layouts.map((l) => l.layoutId))).size ===
      v.masters.reduce((n, m) => n + m.layouts.length, 0) &&
    Array.isArray(v.dependencies) &&
    v.dependencies.length === v.slideIds.length &&
    v.dependencies.every(
      (d, i) =>
        d.slideId === v.slideIds[i] &&
        v.masters.some(
          (m) => m.masterId === d.masterId && m.layouts.some((l) => l.layoutId === d.layoutId),
        ),
    )
  )
}
const actions = ['inspect', 'reconcile', 'resume', 'undo', 'discard', 'capture', 'review'] as const
const tools: AgentToolDef[] = actions.map((action) => ({
  name: `${action}_master_xml_change`,
  description: `${action} a durable master XML change; screenshot reviews are historical and do not certify the deck.`,
  inputSchema: {
    type: 'object',
    properties: {
      change_id: { type: 'string' },
      ...(['capture', 'review'].includes(action) ? { slide_id: { type: 'string' } } : {}),
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
      ...(['capture', 'review'].includes(action) ? ['slide_id'] : []),
      ...(action === 'review' ? ['screenshot_digest', 'status', 'notes'] : []),
    ],
    additionalProperties: false,
  },
}))
export function createPresentationMasterXmlSkill(options: PresentationMasterXmlOptions) {
  let epoch = 0,
    visualEpoch = 0
  const captures = new Map<string, { digest: string; record: string; epoch: number }>()
  const invalidateVisual = () => {
    visualEpoch++
    captures.clear()
  }
  const syncGuard = (c: Context) => {
    if (c.signal?.aborted || c.originalSignal?.aborted || c.epoch !== epoch)
      throw Error('cancelled')
    if (!options.available()) fail('persistence_unavailable')
    options.assertDocumentId?.(c.documentId)
    if (!same(options.readMasterXmlChange(c.changeId), c.record)) fail('stale')
  }
  const guard = async (c: Context) => {
    syncGuard(c)
    if ((await options.documentId()) !== c.documentId) throw Error('presentation_document_changed')
    syncGuard(c)
  }
  const awaited = async <T>(c: Context, run: () => Promise<T>): Promise<T> => {
    await guard(c)
    const result = await run()
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
    const owned = Uint8Array.from(bytes)
    const ref = await awaited(c, () =>
      savePackageBackup({
        request: request(c),
        documentId: c.documentId,
        changeId: c.changeId,
        key,
        bytes: owned,
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
    if (!validHost(host(actual))) fail('invalid')
    if (exports.some((id) => !actual.pages.find((p) => p.slideId === id)?.base64)) fail('invalid')
    return actual
  }
  const store = async (c: Context, next: PresentationMasterXmlChange) => {
    const owned = structuredClone(next)
    if (!validMasterXmlTransition(c.record, owned)) fail('state_invalid')
    await guard(c)
    await options.writeMasterXmlChange(
      structuredClone(owned),
      c.record && structuredClone(c.record),
    )
    c.record = owned
    await guard(c)
    return owned
  }
  const context = (r: PresentationMasterXmlChange, signal?: AbortSignal): Context => ({
    documentId: r.documentId,
    changeId: r.changeId,
    epoch,
    record: structuredClone(r),
    signal,
    originalSignal: signal,
  })
  const record = (id: string) => {
    const r = structuredClone(options.readMasterXmlChange(id))
    if (!validatePresentationMasterXmlChange(r)) fail('state_invalid')
    return r
  }
  const summary = (r: PresentationMasterXmlChange) => ({
    changeId: r.changeId,
    state: r.state,
    cursor: r.cursor,
    receiptCount: r.receiptCount,
    pending: r.pending?.action,
    affectedPageCount: r.scope.affectedPageCount,
    affectedMasterCount: r.scope.affectedMasterCount,
    introducedMasterCount: r.introducedMasterCount,
    inventoryCleanupVerified: false,
    qaPassed: false,
  })
  let verifiedTip: PackageBackupRef | undefined
  const exact = (v: any, required: string[], optional: string[] = []) =>
    v &&
    typeof v === 'object' &&
    !Array.isArray(v) &&
    required.every((k) => Object.hasOwn(v, k)) &&
    Object.keys(v).every((k) => required.includes(k) || optional.includes(k))
  const validSource = (v: any) =>
    exact(v, ['slideId', 'masterId', 'layoutId']) &&
    [v.slideId, v.masterId, v.layoutId].every(validPackageHostId)
  const validateProof = (value: any, s: Snapshot, sequence: number): value is Proof => {
    if (
      !exact(
        value,
        ['version', 'documentId', 'changeId', 'sequence', 'host', 'progress'],
        ['previous', 'action', 'step', 'next'],
      ) ||
      value.version !== 1 ||
      value.documentId !== s.documentId ||
      value.changeId !== s.changeId ||
      value.sequence !== sequence ||
      !validHost(value.host)
    )
      return false
    const p = value.progress,
      required = [
        'originalRepresentatives',
        'originalMappings',
        'importedRepresentatives',
        'importedMappings',
        'restoredRepresentatives',
        'restoredMappings',
        'ownedSources',
        'roles',
        'pageMappings',
        'mode',
        'fallbackOriginalIds',
        'importedMasterIds',
        'restoredMasterIds',
        'introducedMasterIds',
      ]
    if (
      !exact(p, required, ['restorationStaged', 'currentBackups', 'masterCarriers']) ||
      !['forward', 'undo', 'discard'].includes(p.mode) ||
      !Array.isArray(p.pageMappings) ||
      p.pageMappings.length !== s.pages.length ||
      p.pageMappings.some(
        (m: any, i: number) =>
          !exact(m, ['originalSlideId', 'currentSlideId']) ||
          m.originalSlideId !== s.pages[i]!.slideId ||
          !validPackageHostId(m.currentSlideId),
      ) ||
      new Set(p.pageMappings.map((m: any) => m.currentSlideId)).size !== p.pageMappings.length
    )
      return false
    for (const key of [
      'importedMasterIds',
      'restoredMasterIds',
      'introducedMasterIds',
      'fallbackOriginalIds',
    ])
      if (
        !Array.isArray(p[key]) ||
        !p[key].every(validPackageHostId) ||
        new Set(p[key]).size !== p[key].length
      )
        return false
    if (
      p.masterCarriers !== undefined &&
      (!Array.isArray(p.masterCarriers) ||
        p.masterCarriers.some(
          (v: any) =>
            !exact(v, ['mode', 'packageMasterPath', 'identity', 'packageRef']) ||
            !['imported', 'restored'].includes(v.mode) ||
            !s.originalInventory.masters.some((m) => m.path === v.packageMasterPath) ||
            !validSource(v.identity) ||
            !validPackageBackupRef(v.packageRef) ||
            !p.ownedSources.some((o: any) => same(o, v.identity)),
        ))
    )
      return false
    if (
      p.currentBackups !== undefined &&
      (!Array.isArray(p.currentBackups) ||
        p.currentBackups.some(
          (v: any) =>
            !exact(v, ['slideId', 'packageRef', 'digest'], ['originalSlideId']) ||
            !validPackageHostId(v.slideId) ||
            !validPackageBackupRef(v.packageRef) ||
            typeof v.digest !== 'string' ||
            !/^[a-f0-9]{64}$/.test(v.digest) ||
            (v.originalSlideId !== undefined &&
              !s.pages.some((o) => o.slideId === v.originalSlideId)),
        ))
    )
      return false
    if (p.fallbackOriginalIds.some((id: string) => !s.pages.some((v) => v.slideId === id)))
      return false
    if (
      !Array.isArray(p.ownedSources) ||
      !p.ownedSources.every(validSource) ||
      new Set(p.ownedSources.map((v: any) => v.slideId)).size !== p.ownedSources.length ||
      p.ownedSources.some((v: any) => s.original.slideIds.includes(v.slideId))
    )
      return false
    if (
      !exact(
        p.roles,
        [],
        [
          'stagedSource',
          'restoredSource',
          'originalProbeSource',
          'importedProbeSource',
          'restoredProbeSource',
        ],
      ) ||
      Object.values(p.roles).some(
        (v: any) => !validSource(v) || !p.ownedSources.some((o: any) => same(o, v)),
      )
    )
      return false
    if (
      p.restorationStaged &&
      (!exact(p.restorationStaged, ['originalSlideId', 'identity']) ||
        !s.pages.some((v) => v.slideId === p.restorationStaged.originalSlideId) ||
        !validSource(p.restorationStaged.identity) ||
        !p.ownedSources.some((v: any) => same(v, p.restorationStaged.identity)))
    )
      return false
    for (const key of [
      'originalRepresentatives',
      'importedRepresentatives',
      'restoredRepresentatives',
    ])
      if (
        !Array.isArray(p[key]) ||
        p[key].some(
          (v: any) =>
            !exact(v, ['masterId', 'layoutId', 'packageRef']) ||
            ![v.masterId, v.layoutId].every(validPackageHostId) ||
            !validPackageBackupRef(v.packageRef) ||
            !v.packageRef.key.startsWith('page-'),
        ) ||
        new Set(p[key].map((v: any) => `${v.masterId}:${v.layoutId}`)).size !== p[key].length
      )
        return false
    for (const key of ['originalMappings', 'importedMappings', 'restoredMappings']) {
      if (
        !Array.isArray(p[key]) ||
        (p[key].length !== 0 && p[key].length !== s.originalInventory.masters.length) ||
        new Set(p[key].map((v: any) => v.packageMasterPath)).size !== p[key].length ||
        new Set(p[key].map((v: any) => v.mapping?.masterId)).size !== p[key].length
      )
        return false
      for (const v of p[key]) {
        const expected = s.originalInventory.masters.find((m) => m.path === v.packageMasterPath),
          m = v.mapping
        if (
          !exact(v, ['packageMasterPath', 'mapping']) ||
          !expected ||
          !exact(m, ['masterId', 'sourceLayoutId', 'layouts']) ||
          ![m.masterId, m.sourceLayoutId].every(validPackageHostId) ||
          !Array.isArray(m.layouts) ||
          m.layouts.length !== expected.orderedLayouts.length ||
          m.layouts.some(
            (l: any, i: number) =>
              !exact(l, ['packageLayoutPath', 'nativeLayoutId']) ||
              l.packageLayoutPath !== expected.orderedLayouts[i]!.path ||
              !validPackageHostId(l.nativeLayoutId),
          ) ||
          new Set(m.layouts.map((l: any) => l.nativeLayoutId)).size !== m.layouts.length ||
          !m.layouts.some((l: any) => l.nativeLayoutId === m.sourceLayoutId)
        )
          return false
      }
    }
    if (sequence === 0)
      return (
        !value.previous &&
        !value.action &&
        !value.step &&
        !value.next &&
        same(value.host, s.original) &&
        !p.ownedSources.length &&
        !Object.keys(p.roles).length &&
        !p.originalMappings.length &&
        !p.importedMappings.length &&
        !p.restoredMappings.length &&
        p.mode === 'forward' &&
        !p.fallbackOriginalIds.length &&
        p.pageMappings.every((m: any) => m.originalSlideId === m.currentSlideId) &&
        same(
          p.originalRepresentatives,
          s.pages
            .filter(
              (v, i, all) =>
                all.findIndex((x) => x.masterId === v.masterId && x.layoutId === v.layoutId) === i,
            )
            .map((v) => ({ masterId: v.masterId, layoutId: v.layoutId, packageRef: v.packageRef })),
        )
      )
    return (
      validPackageBackupRef(value.previous) &&
      value.previous.key === `receipt-${sequence - 1}` &&
      exact(
        value.step,
        ['kind', 'index'],
        [
          'targetSlideId',
          'masterId',
          'layoutId',
          'sourceId',
          'originalId',
          'inserted',
          'beforeBackupRef',
          'recoveryPreimageRef',
        ],
      ) &&
      ['stage', 'layout', 'remove', 'checkpoint'].includes(value.step.kind) &&
      Number.isSafeInteger(value.step.index) &&
      value.step.index >= 0 &&
      ['targetSlideId', 'masterId', 'layoutId', 'sourceId', 'originalId'].every(
        (k) => value.step[k] === undefined || validPackageHostId(value.step[k]),
      ) &&
      (!value.step.inserted || validSource(value.step.inserted)) &&
      (!value.step.beforeBackupRef || validPackageBackupRef(value.step.beforeBackupRef)) &&
      (!value.step.recoveryPreimageRef || validPackageBackupRef(value.step.recoveryPreimageRef)) &&
      typeof value.action === 'string' &&
      exact(value.next, ['state', 'phase', 'index']) &&
      Number.isSafeInteger(value.next.index) &&
      value.next.index >= 0
    )
  }
  const validateEdge = (before: Proof, after: Proof, s: Snapshot) => {
    const step = after.step!,
      a = after.host,
      b = before.host,
      p = before.progress,
      n = after.progress
    if (step.kind === 'stage') {
      const inserted = step.inserted
      if (
        !inserted ||
        !step.sourceId ||
        !b.slideIds.includes(step.sourceId) ||
        b.slideIds.includes(inserted.slideId)
      )
        fail('backup_invalid')
      const ids = [...b.slideIds]
      ids.splice(ids.indexOf(step.sourceId) + 1, 0, inserted.slideId)
      if (
        !same(ids, a.slideIds) ||
        !unchanged(b, a, []) ||
        !same(
          a.dependencies.find((d) => d.slideId === inserted.slideId),
          inserted,
        ) ||
        !same(n.ownedSources, [...p.ownedSources, inserted])
      )
        fail('backup_invalid')
    } else {
      if (!same(p.ownedSources, n.ownedSources)) fail('backup_invalid')
      if (step.kind === 'layout') {
        if (
          !step.targetSlideId ||
          !step.masterId ||
          !step.layoutId ||
          !same(a.slideIds, b.slideIds) ||
          !same(a.masters, b.masters) ||
          !unchanged(
            b,
            a,
            after.next?.state === 'recovery_required'
              ? targetsFor(s, before)
              : [step.targetSlideId],
          ) ||
          (after.next?.state !== 'recovery_required' &&
            !same(
              a.dependencies.find((d) => d.slideId === step.targetSlideId),
              { slideId: step.targetSlideId, masterId: step.masterId, layoutId: step.layoutId },
            ))
        )
          fail('backup_invalid')
      } else if (step.kind === 'remove') {
        if (
          !step.targetSlideId ||
          !b.slideIds.includes(step.targetSlideId) ||
          !same(
            a.slideIds,
            b.slideIds.filter((id) => id !== step.targetSlideId),
          ) ||
          !unchanged(b, a, [step.targetSlideId]) ||
          ![
            ...p.ownedSources.map((v) => v.slideId),
            ...(after.action === 'delete_source' ? [s.sourceSlideId] : []),
            ...(after.action === 'restore_page_delete' && step.originalId
              ? p.pageMappings
                  .filter((v) => v.originalSlideId === step.originalId)
                  .map((v) => v.currentSlideId)
              : []),
          ].includes(step.targetSlideId)
        )
          fail('backup_invalid')
      } else if (!same(a, b)) fail('backup_invalid')
    }
    if (
      Object.entries(p.roles).some(([k, v]) => !same(v, n.roles[k as keyof typeof n.roles])) ||
      p.introducedMasterIds.some((id) => !n.introducedMasterIds.includes(id))
    )
      fail('backup_invalid')
    const addedBackups = (n.currentBackups ?? []).slice(p.currentBackups?.length ?? 0)
    if (['layout', 'remove'].includes(step.kind)) {
      if (
        !step.beforeBackupRef ||
        !addedBackups.some(
          (v) =>
            v.slideId === step.targetSlideId &&
            same(v.packageRef, step.beforeBackupRef) &&
            v.digest === b.pages.find((o) => o.slideId === step.targetSlideId)?.digest,
        )
      )
        fail('backup_invalid')
    } else if (addedBackups.length || step.beforeBackupRef || step.recoveryPreimageRef)
      fail('backup_invalid')
    if (step.recoveryPreimageRef) {
      if (
        step.kind !== 'layout' ||
        after.next?.state !== 'recovery_required' ||
        !targetsFor(s, before).every((id) =>
          addedBackups.some(
            (v) => v.slideId === id && v.digest === a.pages.find((o) => o.slideId === id)?.digest,
          ),
        )
      )
        fail('backup_invalid')
    }
    if (
      (p.currentBackups ?? []).some((v, i) => !same(v, n.currentBackups?.[i])) ||
      (p.masterCarriers ?? []).some((v, i) => !same(v, n.masterCarriers?.[i]))
    )
      fail('backup_invalid')
    const addedCarriers = (n.masterCarriers ?? []).slice(p.masterCarriers?.length ?? 0)
    if (
      ['stage', 'stage_master', 'restore_stage', 'restore_stage_master'].includes(after.action!)
    ) {
      const restoring = after.action!.startsWith('restore'),
        inventory = restoring ? s.originalInventory : s.preparedInventory
      const path =
        after.action === 'stage' || after.action === 'restore_stage'
          ? inventory.sourceMasterPath
          : inventory.masters.filter((m) => m.path !== inventory.sourceMasterPath)[step.index]?.path
      if (
        addedCarriers.length !== 1 ||
        addedCarriers[0]!.packageMasterPath !== path ||
        addedCarriers[0]!.mode !== (restoring ? 'restored' : 'imported') ||
        !same(addedCarriers[0]!.identity, step.inserted)
      )
        fail('backup_invalid')
    } else if (addedCarriers.length) fail('backup_invalid')
    for (const key of ['originalMappings', 'importedMappings', 'restoredMappings'] as const)
      if (p[key].length && !same(p[key], n[key])) fail('backup_invalid')
    if (
      p.pageMappings.some((v, i) => v.originalSlideId !== n.pageMappings[i]!.originalSlideId) ||
      (!['delete_source', 'restore_page', 'restore_page_delete'].includes(after.action!) &&
        !same(p.pageMappings, n.pageMappings))
    )
      fail('backup_invalid')
  }
  const chain = async (c: Context, s: Snapshot, ref: PackageBackupRef, p: Proof) => {
    if (!validateProof(p, s, Number(ref.key.slice(8)))) fail('backup_invalid')
    if (verifiedTip && same(ref, verifiedTip)) return
    let current = p,
      depth = 0
    while (current.sequence) {
      if (++depth > 4095) fail('backup_invalid')
      const previous = (await readJson(c, current.previous!)) as Proof
      if (!validateProof(previous, s, current.sequence - 1)) fail('backup_invalid')
      validateEdge(previous, current, s)
      if (verifiedTip && same(current.previous, verifiedTip)) break
      current = previous
    }
    if (!current.sequence && !validateProof(current, s, 0)) fail('backup_invalid')
    verifiedTip = structuredClone(ref)
  }
  const load = async (c: Context, r: PresentationMasterXmlChange) => {
    const s = (await readJson(c, r.snapshotRef)) as Snapshot
    if (
      !s ||
      s.version !== 1 ||
      s.documentId !== r.documentId ||
      s.changeId !== r.changeId ||
      s.sourceSlideId !== r.sourceSlideId ||
      s.packageSourceSlideId !== r.packageSourceSlideId ||
      !validHost(s.original) ||
      !Array.isArray(s.pages) ||
      s.pages.length !== s.original.slideIds.length ||
      s.pages.some(
        (p, i) =>
          p.slideId !== s.original.slideIds[i] ||
          !validPackageBackupRef(p.packageRef) ||
          p.masterId !== s.original.dependencies[i]!.masterId ||
          p.layoutId !== s.original.dependencies[i]!.layoutId,
      ) ||
      !same(s.preparedRef, r.preparedRef) ||
      !validPackageBackupRef(s.carrierRef)
    )
      fail('backup_invalid')
    const original = encode(await read(c, s.carrierRef)),
      prepared = encode(await read(c, s.preparedRef))
    const rebuilt = await awaited(c, () =>
      editPowerPointPackage(original, 'master', structuredClone(s.replacements), c.signal),
    )
    const prep = await awaited(c, () =>
      assertMasterXmlPreparation(original, prepared, rebuilt.changedPaths, c.signal),
    )
    if (
      rebuilt.base64 !== prepared &&
      (await awaited(c, () => presentationPackageDigest(rebuilt.base64, c.signal))) !==
        prep.prepared.packageDigest
    )
      fail('backup_invalid')
    if (
      !same(prep.original, s.originalInventory) ||
      !same(prep.prepared, s.preparedInventory) ||
      !same(prep.affectedMasterPaths, s.affectedMasterPaths)
    )
      fail('backup_invalid')
    const p = (await readJson(c, r.currentProofRef)) as Proof
    if (
      !p ||
      p.version !== 1 ||
      p.documentId !== r.documentId ||
      p.changeId !== r.changeId ||
      p.sequence !== r.receiptCount ||
      !validHost(p.host) ||
      !p.progress
    )
      fail('backup_invalid')
    try {
      await chain(c, s, r.currentProofRef, p)
    } catch (error) {
      verifiedTip = undefined
      throw error
    }
    if (
      Object.keys(p.progress.roles).some(
        (k) =>
          !same(
            p.progress.roles[k as keyof typeof p.progress.roles],
            r[k as keyof PresentationMasterXmlChange],
          ),
      ) ||
      r.introducedMasterCount !== p.progress.introducedMasterIds.length
    )
      fail('backup_invalid')
    return { s, p, original, prepared }
  }
  const checked = async (c: Context, r: PresentationMasterXmlChange) => {
    const data = await load(c, r)
    if (!same(host(await inspect(c)), data.p.host)) fail(r.pending ? 'unknown' : 'drift')
    return data
  }
  const mapMaster = async (
    c: Context,
    expected: string,
    inventory: MasterXmlPackageInventory,
    reps: Progress['originalRepresentatives'],
    native: Host['masters'],
    anchor?: { path: string; masterId: string },
  ) => {
    const result: Progress['originalMappings'] = [],
      used = new Set<string>()
    const paths = inventory.masters
      .map((m) => m.path)
      .sort((a, b) => (a === anchor?.path ? -1 : b === anchor?.path ? 1 : 0))
    for (const path of paths) {
      const candidates = native.filter(
        (m) => !used.has(m.masterId) && (path !== anchor?.path || m.masterId === anchor.masterId),
      )
      const fully = candidates.filter((m) =>
        m.layouts.every((l) =>
          reps.some((r) => r.masterId === m.masterId && r.layoutId === l.layoutId),
        ),
      )
      const masters = [] as {
        masterId: string
        layouts: { layoutId: string; representativeBase64: string }[]
      }[]
      for (const m of fully) {
        const layouts = [] as { layoutId: string; representativeBase64: string }[]
        for (const l of m.layouts) {
          const rep = reps.find((r) => r.masterId === m.masterId && r.layoutId === l.layoutId)!
          layouts.push({
            layoutId: l.layoutId,
            representativeBase64: encode(await read(c, rep.packageRef)),
          })
        }
        masters.push({ masterId: m.masterId, layouts })
      }
      const mapping = await awaited(c, () =>
        proveMasterXmlLayoutMapping(expected, { masters }, c.signal, path),
      )
      used.add(mapping.masterId)
      result.push({ packageMasterPath: path, mapping })
    }
    return result
  }
  const pageBytes = async (c: Context, s: Snapshot, id: string) => {
    const page = s.pages.find((p) => p.slideId === id)
    if (!page) fail('backup_invalid')
    const bytes = await read(c, page!.packageRef)
    const base64 = encode(bytes)
    if (
      (await awaited(c, () => presentationPackageDigest(base64, c.signal))) !==
      s.original.pages.find((p) => p.slideId === id)!.digest
    )
      fail('backup_invalid')
    return base64
  }
  const provePage = async (
    c: Context,
    s: Snapshot,
    p: Proof,
    originalId: string,
    actualId: string,
    prepared: boolean,
    actual: Host,
  ) => {
    const originalPage = s.pages.find((v) => v.slideId === originalId)!
    const oldMap = p.progress.originalMappings.find(
      (v) => v.mapping.masterId === originalPage.masterId,
    )
    if (!oldMap) fail('mapping_unproven')
    const slot = oldMap!.mapping.layouts.find((v) => v.nativeLayoutId === originalPage.layoutId)!
    if (!slot) fail('mapping_unproven')
    const dependency = actual.dependencies.find((v) => v.slideId === actualId)
    const maps = prepared
      ? p.progress.importedMappings
      : p.progress.restoredMappings.length
        ? p.progress.restoredMappings
        : p.progress.originalMappings
    const expectedMap = maps.find((v) => v.packageMasterPath === oldMap!.packageMasterPath)
    if (
      !dependency ||
      !expectedMap ||
      dependency.masterId !== expectedMap.mapping.masterId ||
      dependency.layoutId !==
        expectedMap.mapping.layouts.find((v) => v.packageLayoutPath === slot.packageLayoutPath)
          ?.nativeLayoutId
    )
      fail('mapping_unproven')
    const original = await pageBytes(c, s, originalId),
      carrier = encode(await read(c, prepared ? s.preparedRef : s.carrierRef)),
      actualBase64 = actual.pages.find((v) => v.slideId === actualId)?.base64
    if (!actualBase64) fail('invalid')
    await awaited(c, () =>
      assertMasterXmlPagePreserved(
        original,
        actualBase64!,
        {
          expectedMasterBase64: carrier,
          targetMasterPath: oldMap!.packageMasterPath,
          packageLayoutPath: slot.packageLayoutPath,
        },
        c.signal,
      ),
    )
  }
  const unchanged = (before: Host, after: Host, exclude: string[]) =>
    before.slideIds
      .filter((id) => !exclude.includes(id))
      .every(
        (id) =>
          same(
            before.pages.find((p) => p.slideId === id),
            after.pages.find((p) => p.slideId === id) && {
              slideId: id,
              digest: after.pages.find((p) => p.slideId === id)!.digest,
            },
          ) &&
          same(
            before.dependencies.find((d) => d.slideId === id),
            after.dependencies.find((d) => d.slideId === id),
          ),
      )
  const missingOriginal = (s: Snapshot, p: Proof) =>
    s.original.masters.flatMap((m) =>
      m.layouts
        .filter(
          (l) =>
            !p.progress.originalRepresentatives.some(
              (v) => v.masterId === m.masterId && v.layoutId === l.layoutId,
            ),
        )
        .map((l) => ({ masterId: m.masterId, layoutId: l.layoutId })),
    )
  const candidateImported = (_s: Snapshot, p: Proof, _r: PresentationMasterXmlChange) =>
    p.host.masters.filter((m) => p.progress.importedMasterIds.includes(m.masterId))
  const candidateRestored = (_s: Snapshot, p: Proof, _r: PresentationMasterXmlChange) =>
    p.host.masters.filter(
      (m) =>
        p.progress.restoredMasterIds.includes(m.masterId) ||
        p.progress.originalMappings.some((v) => v.mapping.masterId === m.masterId),
    )
  const layouts = (masters: Host['masters']) =>
    masters.flatMap((m) => m.layouts.map((l) => ({ masterId: m.masterId, layoutId: l.layoutId })))
  const affected = (s: Snapshot, p: Proof) =>
    s.pages.filter(
      (page) =>
        page.slideId === s.sourceSlideId ||
        p.progress.originalMappings.some(
          (m) =>
            m.mapping.masterId === page.masterId &&
            s.affectedMasterPaths.includes(m.packageMasterPath),
        ),
    )
  const targetLayout = (p: Proof, page: OriginalPage, maps: Progress['originalMappings']) => {
    const old = p.progress.originalMappings.find((m) => m.mapping.masterId === page.masterId),
      slot = old?.mapping.layouts.find((l) => l.nativeLayoutId === page.layoutId),
      next = maps.find((m) => m.packageMasterPath === old?.packageMasterPath),
      target = next?.mapping.layouts.find((l) => l.packageLayoutPath === slot?.packageLayoutPath)
    if (!old || !slot || !next || !target) fail('mapping_unproven')
    return { masterId: next!.mapping.masterId, layoutId: target!.nativeLayoutId }
  }
  type Command = {
    action: MasterXmlAction
    kind: CommandKind
    targetSlideId?: string
    masterId?: string
    layoutId?: string
    base64?: string
    sourceId?: string
    originalId?: string
    nextState: MasterXmlState
    nextPhase: MasterXmlAction
    nextIndex: number
    carrierPath?: string
    carrierRef?: PackageBackupRef
    capture?: 'original' | 'imported' | 'restored'
  }
  const command = async (
    c: Context,
    r: PresentationMasterXmlChange,
    s: Snapshot,
    p: Proof,
  ): Promise<Command> => {
    const phase = r.cursor.phase,
      i = r.cursor.index,
      original = encode(await read(c, s.carrierRef)),
      prepared = encode(await read(c, s.preparedRef))
    if (phase === 'original_probe_stage')
      return {
        action: phase,
        kind: 'stage',
        base64: original,
        sourceId: s.sourceSlideId,
        nextState: 'probing_original',
        nextPhase: missingOriginal(s, p).length ? 'original_probe' : 'original_probe_delete',
        nextIndex: 0,
      }
    if (phase === 'original_probe') {
      const all = layouts(s.original.masters).filter(
          (l) => !s.pages.some((v) => v.masterId === l.masterId && v.layoutId === l.layoutId),
        ),
        target = all[i]
      if (!target) fail('state_invalid')
      return {
        action: phase,
        kind: 'layout',
        targetSlideId: r.originalProbeSource!.slideId,
        ...target,
        nextState: 'probing_original',
        nextPhase: i + 1 < all.length ? 'original_probe' : 'original_probe_delete',
        nextIndex: i + 1 < all.length ? i + 1 : 0,
        capture: 'original',
      }
    }
    if (phase === 'original_probe_delete') {
      p.progress.originalMappings = await mapMaster(
        c,
        original,
        s.originalInventory,
        p.progress.originalRepresentatives,
        s.original.masters,
        {
          path: s.originalInventory.sourceMasterPath,
          masterId: s.original.dependencies.find((d) => d.slideId === s.sourceSlideId)!.masterId,
        },
      )
      return {
        action: phase,
        kind: 'remove',
        targetSlideId: r.originalProbeSource!.slideId,
        nextState: 'prepared',
        nextPhase: 'stage',
        nextIndex: 0,
      }
    }
    if (phase === 'stage')
      return {
        action: phase,
        kind: 'stage',
        base64: prepared,
        carrierPath: s.preparedInventory.sourceMasterPath,
        carrierRef: s.preparedRef,
        sourceId: s.sourceSlideId,
        nextState: 'staged',
        nextPhase: s.preparedInventory.masters.length > 1 ? 'stage_master' : 'import_probe_stage',
        nextIndex: 0,
      }
    if (phase === 'stage_master' || phase === 'restore_stage_master') {
      const restoring = phase === 'restore_stage_master',
        inventory = restoring ? s.originalInventory : s.preparedInventory
      const others = inventory.masters.filter((m) => m.path !== inventory.sourceMasterPath),
        m = others[i]
      if (!m) fail('state_invalid')
      const derived = await awaited(c, () =>
        deriveMasterXmlCarrier(restoring ? original : prepared, m.path, undefined, c.signal),
      )
      const bytes = decode(derived.base64),
        hash = Array.from(
          new Uint8Array(
            await awaited(c, () => crypto.subtle.digest('SHA-256', bytes.buffer as ArrayBuffer)),
          ),
        )
          .map((v) => v.toString(16).padStart(2, '0'))
          .join('')
      const carrierRef = await save(
        c,
        `page-7${BigInt('0x' + hash)
          .toString()
          .padStart(78, '0')}`,
        bytes,
      )
      return {
        action: phase,
        kind: 'stage',
        base64: derived.base64,
        carrierPath: m.path,
        carrierRef,
        sourceId: restoring ? r.restoredSource!.slideId : r.stagedSource!.slideId,
        nextState: restoring ? 'restoring' : 'staged',
        nextPhase:
          i + 1 < others.length ? phase : restoring ? 'restore_probe_stage' : 'import_probe_stage',
        nextIndex: i + 1 < others.length ? i + 1 : 0,
      }
    }
    if (phase === 'import_probe_stage')
      return {
        action: phase,
        kind: 'stage',
        base64: prepared,
        sourceId: r.stagedSource!.slideId,
        nextState: 'probing_imported',
        nextPhase: 'import_probe',
        nextIndex: 0,
      }
    if (phase === 'import_probe') {
      const all = layouts(candidateImported(s, p, r)).filter(
          (l) =>
            l.masterId !== r.importedProbeSource?.masterId ||
            l.masterId === r.stagedSource?.masterId,
        ),
        target = all[i]
      if (!target) fail('state_invalid')
      return {
        action: phase,
        kind: 'layout',
        targetSlideId: r.importedProbeSource!.slideId,
        ...target,
        nextState: 'probing_imported',
        nextPhase: i + 1 < all.length ? 'import_probe' : 'import_probe_delete',
        nextIndex: i + 1 < all.length ? i + 1 : 0,
        capture: 'imported',
      }
    }
    if (phase === 'import_probe_delete') {
      const candidates = candidateImported(s, p, r).filter(
        (m) =>
          m.masterId !== r.importedProbeSource?.masterId || m.masterId === r.stagedSource?.masterId,
      )
      p.progress.importedMappings = []
      for (const carrier of p.progress.masterCarriers!.filter((v) => v.mode === 'imported')) {
        const m = s.preparedInventory.masters.find((m) => m.path === carrier.packageMasterPath)!
        p.progress.importedMappings.push(
          ...(await mapMaster(
            c,
            prepared,
            { ...s.preparedInventory, masters: [m] },
            p.progress.importedRepresentatives,
            candidates.filter((n) => n.masterId === carrier.identity.masterId),
            { path: m.path, masterId: carrier.identity.masterId },
          )),
        )
      }
      return {
        action: phase,
        kind: 'remove',
        targetSlideId: r.importedProbeSource!.slideId,
        nextState: 'staged',
        nextPhase: affected(s, p).some((v) => v.slideId !== s.sourceSlideId)
          ? 'forward_page'
          : 'delete_stage_master',
        nextIndex: 0,
      }
    }
    if (phase === 'forward_page') {
      const all = affected(s, p).filter((v) => v.slideId !== s.sourceSlideId),
        page = all[i]
      if (!page) fail('state_invalid')
      return {
        action: phase,
        kind: 'layout',
        targetSlideId: page.slideId,
        ...targetLayout(p, page, p.progress.importedMappings),
        originalId: page.slideId,
        nextState: 'applying',
        nextPhase: i + 1 < all.length ? 'forward_page' : 'delete_stage_master',
        nextIndex: i + 1 < all.length ? i + 1 : 0,
      }
    }
    if (phase === 'delete_stage_master') {
      const extra = p.progress.masterCarriers?.find(
        (v) =>
          v.mode === 'imported' &&
          v.identity.slideId !== r.stagedSource!.slideId &&
          p.host.slideIds.includes(v.identity.slideId),
      )
      return extra
        ? {
            action: phase,
            kind: 'remove',
            targetSlideId: extra.identity.slideId,
            nextState: 'applying',
            nextPhase: phase,
            nextIndex: 0,
          }
        : {
            action: phase,
            kind: 'checkpoint',
            targetSlideId: r.stagedSource!.slideId,
            nextState: 'applying',
            nextPhase: 'delete_source',
            nextIndex: 0,
          }
    }
    if (phase === 'delete_source')
      return {
        action: phase,
        kind: 'remove',
        targetSlideId: s.sourceSlideId,
        nextState: 'applied',
        nextPhase: 'verify',
        nextIndex: 0,
      }
    if (phase === 'restore_stage')
      return {
        action: phase,
        kind: 'stage',
        base64: original,
        carrierPath: s.originalInventory.sourceMasterPath,
        carrierRef: s.carrierRef,
        sourceId: r.stagedSource?.slideId ?? s.sourceSlideId,
        nextState: 'restoring',
        nextPhase:
          s.originalInventory.masters.length > 1 ? 'restore_stage_master' : 'restore_probe_stage',
        nextIndex: 0,
      }
    if (phase === 'restore_probe_stage')
      return {
        action: phase,
        kind: 'stage',
        base64: original,
        sourceId: r.restoredSource!.slideId,
        nextState: 'probing_restored',
        nextPhase: 'restore_probe',
        nextIndex: 0,
      }
    if (phase === 'restore_probe') {
      const all = layouts(candidateRestored(s, p, r)),
        target = all[i]
      if (!target) fail('state_invalid')
      return {
        action: phase,
        kind: 'layout',
        targetSlideId: r.restoredProbeSource!.slideId,
        ...target,
        nextState: 'probing_restored',
        nextPhase: i + 1 < all.length ? 'restore_probe' : 'restore_probe_delete',
        nextIndex: i + 1 < all.length ? i + 1 : 0,
        capture: 'restored',
      }
    }
    if (phase === 'restore_probe_delete') {
      const native = candidateRestored(s, p, r),
        preferred = p.progress.originalMappings.filter((m) =>
          native.some((n) => n.masterId === m.mapping.masterId),
        ),
        maps: Progress['originalMappings'] = []
      for (const m of s.originalInventory.masters) {
        let mapping: MasterXmlLayoutMapping | undefined
        const old = preferred.find((v) => v.packageMasterPath === m.path)
        if (old) {
          try {
            mapping = (
              await mapMaster(
                c,
                original,
                {
                  ...s.originalInventory,
                  masters: [m],
                  sourceMasterPath: m.path,
                  sourceLayoutPath: m.orderedLayouts[0]!.path,
                },
                p.progress.restoredRepresentatives,
                native.filter((n) => n.masterId === old.mapping.masterId),
              )
            )[0]!.mapping
          } catch (error) {
            if (c.signal?.aborted || c.originalSignal?.aborted || c.epoch !== epoch) throw error
          }
        }
        if (!mapping) {
          const carrier = p.progress.masterCarriers?.find(
            (v) => v.mode === 'restored' && v.packageMasterPath === m.path,
          )
          const candidates = native.filter((n) => n.masterId === carrier?.identity.masterId)
          mapping = (
            await mapMaster(
              c,
              original,
              {
                ...s.originalInventory,
                masters: [m],
                sourceMasterPath: m.path,
                sourceLayoutPath: m.orderedLayouts[0]!.path,
              },
              p.progress.restoredRepresentatives,
              candidates,
              m.path === s.originalInventory.sourceMasterPath
                ? { path: m.path, masterId: r.restoredSource!.masterId }
                : undefined,
            )
          )[0]!.mapping
        }
        maps.push({ packageMasterPath: m.path, mapping })
      }
      p.progress.restoredMappings = maps
      return {
        action: phase,
        kind: 'remove',
        targetSlideId: r.restoredProbeSource!.slideId,
        nextState: 'restoring',
        nextPhase: 'restore_page',
        nextIndex: 0,
      }
    }
    if (
      phase === 'restore_page' ||
      phase === 'restore_page_stage' ||
      phase === 'restore_page_associate' ||
      phase === 'restore_page_delete'
    ) {
      const all = affected(s, p),
        page = all[i]
      if (!page) fail('state_invalid')
      const current = p.progress.pageMappings.find(
          (m) => m.originalSlideId === page.slideId,
        )!.currentSlideId,
        layout = targetLayout(p, page, p.progress.restoredMappings)
      if (phase === 'restore_page_stage')
        return {
          action: phase,
          kind: 'stage',
          base64: await pageBytes(c, s, page.slideId),
          sourceId: current,
          originalId: page.slideId,
          nextState: 'restoring_pages',
          nextPhase: 'restore_page_associate',
          nextIndex: i,
        }
      if (phase === 'restore_page_associate')
        return {
          action: phase,
          kind: 'layout',
          targetSlideId: p.progress.restorationStaged!.identity.slideId,
          ...layout,
          originalId: page.slideId,
          nextState: 'restoring_pages',
          nextPhase: 'restore_page_delete',
          nextIndex: i,
        }
      if (phase === 'restore_page_delete')
        return {
          action: phase,
          kind: 'remove',
          targetSlideId: current,
          originalId: page.slideId,
          nextState: 'restoring_pages',
          nextPhase: i + 1 < all.length ? 'restore_page' : 'delete_applied_source',
          nextIndex: i + 1 < all.length ? i + 1 : 0,
        }
      if (p.progress.fallbackOriginalIds.includes(page.slideId))
        return {
          action: phase,
          kind: 'checkpoint',
          originalId: page.slideId,
          targetSlideId: current,
          ...layout,
          nextState: 'restoring_pages',
          nextPhase: 'restore_page_stage',
          nextIndex: i,
        }
      const target =
        page.slideId === s.sourceSlideId && !p.host.slideIds.includes(s.sourceSlideId)
          ? r.restoredSource!.slideId
          : current
      const dep = p.host.dependencies.find((d) => d.slideId === target)!,
        kind =
          dep.masterId === layout.masterId && dep.layoutId === layout.layoutId
            ? 'checkpoint'
            : 'layout'
      return {
        action: phase,
        kind,
        targetSlideId: target,
        ...layout,
        originalId: page.slideId,
        nextState: 'restoring_pages',
        nextPhase: i + 1 < all.length ? 'restore_page' : 'delete_applied_source',
        nextIndex: i + 1 < all.length ? i + 1 : 0,
      }
    }
    if (phase === 'delete_applied_source') {
      const business = p.progress.pageMappings.map((m) => m.currentSlideId),
        target = [...p.progress.ownedSources]
          .reverse()
          .find((v) => p.host.slideIds.includes(v.slideId) && !business.includes(v.slideId))
      if (!target)
        return {
          action: phase,
          kind: 'checkpoint',
          targetSlideId: p.progress.pageMappings[0]!.currentSlideId,
          nextState: 'undone',
          nextPhase: 'verify',
          nextIndex: 0,
        }
      const remaining = p.progress.ownedSources.filter(
        (v) =>
          v.slideId !== target.slideId &&
          p.host.slideIds.includes(v.slideId) &&
          !business.includes(v.slideId),
      )
      return {
        action: phase,
        kind: 'remove',
        targetSlideId: target.slideId,
        nextState: remaining.length ? 'restoring_pages' : 'undone',
        nextPhase: remaining.length ? phase : 'verify',
        nextIndex: 0,
      }
    }
    if (phase === 'discard') {
      const live = p.progress.ownedSources.filter((v) => p.host.slideIds.includes(v.slideId)),
        target = live.at(-1)
      if (!target)
        return {
          action: phase,
          kind: 'checkpoint',
          targetSlideId: s.sourceSlideId,
          nextState: 'discarded',
          nextPhase: 'verify',
          nextIndex: 0,
        }
      return {
        action: phase,
        kind: 'remove',
        targetSlideId: target.slideId,
        nextState: live.length > 1 ? r.state : 'discarded',
        nextPhase: live.length > 1 ? 'discard' : 'verify',
        nextIndex: 0,
      }
    }
    fail('state_invalid')
  }
  const finish = async (
    c: Context,
    r: PresentationMasterXmlChange,
    nextState: MasterXmlState,
    nextPhase: MasterXmlAction,
    nextIndex: number,
    introducedMasterCount = r.introducedMasterCount,
  ) => {
    if (!r.pending?.afterProofRef) fail('unknown')
    const ref = r.pending!.afterProofRef!,
      pending = r.pending!,
      slots: Partial<PresentationMasterXmlChange> = {}
    const name = (
      {
        original_probe_stage: 'originalProbeSource',
        stage: 'stagedSource',
        import_probe_stage: 'importedProbeSource',
        restore_stage: 'restoredSource',
        restore_probe_stage: 'restoredProbeSource',
      } as const
    )[pending.action as 'stage']
    if (name && !r[name]) Object.assign(slots, { [name]: pending.inserted })
    return store(c, {
      ...r,
      ...slots,
      pending: undefined,
      currentProofRef: ref,
      receiptCount: r.receiptCount + 1,
      introducedMasterCount,
      state: nextState,
      cursor: { phase: nextPhase, index: nextIndex, substep: 0 },
      reviews: [],
    })
  }
  const captureCurrent = async (c: Context, s: Snapshot, p: Proof, id: string, expected: Host) => {
    const page = await awaited(c, async () =>
      structuredClone(await options.adapter.readPage(id, c.signal)),
    )
    const dep = expected.dependencies.find((d) => d.slideId === id)
    if (
      !dep ||
      !same({ slideId: page.slideId, masterId: page.masterId, layoutId: page.layoutId }, dep) ||
      page.digest !== expected.pages.find((v) => v.slideId === id)?.digest
    )
      fail('drift')
    const bytes = decode(page.base64)
    const hash = Array.from(
      new Uint8Array(
        await awaited(c, () => crypto.subtle.digest('SHA-256', bytes.buffer as ArrayBuffer)),
      ),
    )
      .map((b) => b.toString(16).padStart(2, '0'))
      .join('')
    const existing = [
      ...s.pages.map((v) => v.packageRef),
      ...(p.progress.currentBackups ?? []).map((v) => v.packageRef),
    ].find((ref) => ref.sha256 === hash && ref.sizeBytes === bytes.length)
    const packageRef = existing
      ? structuredClone(existing)
      : await save(
          c,
          `page-9${BigInt('0x' + hash)
            .toString()
            .padStart(78, '0')}`,
          bytes,
        )
    if (existing) await read(c, existing)
    if (!same(host(await inspect(c)), expected)) fail('drift')
    const originalSlideId = p.progress.pageMappings.find(
      (m) => m.currentSlideId === id,
    )?.originalSlideId
    return {
      slideId: id,
      ...(originalSlideId ? { originalSlideId } : {}),
      packageRef,
      digest: page.digest,
    }
  }
  const perform = async (
    c: Context,
    r: PresentationMasterXmlChange,
    s: Snapshot,
    p: Proof,
    cmd: Command,
  ) => {
    if (r.pending) fail('unknown')
    const before = host(await inspect(c))
    if (!same(before, p.host)) fail('drift')
    // Verify every original before the first host write; later steps reread their original targets and mappings.
    if (r.receiptCount === 0) for (const page of s.pages) await pageBytes(c, s, page.slideId)
    else if (cmd.originalId) await pageBytes(c, s, cmd.originalId)
    if (!same(host(await inspect(c)), before)) fail('drift')
    // Reject projected proof budget overflow before persisting a host intent; unknown host growth remains unresolved.
    if (json(p).length + 16384 > 8 * 1024 * 1024) fail('backup_budget_exceeded')
    const currentBackup =
      cmd.targetSlideId && ['layout', 'remove'].includes(cmd.kind)
        ? await captureCurrent(c, s, p, cmd.targetSlideId, before)
        : undefined
    const sourceInventory =
      cmd.kind === 'stage'
        ? await awaited(c, () => inspectMasterXmlPackage(cmd.base64!, c.signal))
        : undefined
    r = await store(c, {
      ...r,
      pending: {
        ...(currentBackup ? { currentBackupRef: currentBackup.packageRef } : {}),
        action: cmd.action,
        index: r.cursor.index,
        beforeProofRef: r.currentProofRef,
        ...(cmd.targetSlideId ? { targetSlideId: cmd.targetSlideId } : {}),
        ...(cmd.masterId ? { targetMasterId: cmd.masterId } : {}),
        ...(cmd.layoutId ? { targetLayoutId: cmd.layoutId } : {}),
      },
      reviews: [],
    })
    invalidateVisual()
    if (cmd.kind === 'stage')
      await options.adapter.stage(
        {
          base64: cmd.base64!,
          sourceSlideId: cmd.sourceId!,
          packageSourceSlideId: sourceInventory!.sourceSlideId,
          preimage: structuredClone(before),
        },
        async (inserted) => {
          const ownedInserted = structuredClone(inserted)
          await guard(c)
          r = await store(c, {
            ...r,
            pending: { ...r.pending!, inserted: ownedInserted },
          })
        },
        () => guard(c),
        () => syncGuard(c),
        c.signal,
      )
    else if (cmd.kind === 'layout')
      await options.adapter.applyLayout(
        {
          slideId: cmd.targetSlideId!,
          masterId: cmd.masterId!,
          layoutId: cmd.layoutId!,
          preimage: structuredClone(before),
        },
        () => guard(c),
        () => syncGuard(c),
        c.signal,
      )
    else if (cmd.kind === 'remove')
      await options.adapter.remove(
        { slideId: cmd.targetSlideId!, preimage: structuredClone(before) },
        () => guard(c),
        () => syncGuard(c),
        c.signal,
      )
    await guard(c)
    return receive(c, r, s, p, cmd, before)
  }
  const receive = async (
    c: Context,
    r: PresentationMasterXmlChange,
    s: Snapshot,
    p: Proof,
    cmd: Command,
    before: Host,
    allowUnproven = false,
  ) => {
    const nextProgress = structuredClone(p.progress)
    if (r.pending!.currentBackupRef && cmd.targetSlideId) {
      await read(c, r.pending!.currentBackupRef)
      nextProgress.currentBackups = [
        ...(nextProgress.currentBackups ?? []),
        {
          slideId: cmd.targetSlideId,
          ...(cmd.originalId ? { originalSlideId: cmd.originalId } : {}),
          packageRef: r.pending!.currentBackupRef,
          digest: before.pages.find((v) => v.slideId === cmd.targetSlideId)!.digest,
        },
      ]
    }
    const target = r.pending!.inserted?.slideId ?? cmd.targetSlideId,
      actual = await inspect(c, target && cmd.kind !== 'remove' ? [target] : []),
      after = host(actual)
    let recoveryBackups: NonNullable<Progress['currentBackups']> | undefined
    if (r.pending!.recoveryPreimageRef) {
      const manifest = await readJson(c, r.pending!.recoveryPreimageRef)
      if (
        !exact(manifest, ['version', 'documentId', 'changeId', 'host', 'currentBackups']) ||
        manifest.version !== 1 ||
        manifest.documentId !== c.documentId ||
        manifest.changeId !== c.changeId ||
        !same(manifest.host, after) ||
        !Array.isArray(manifest.currentBackups) ||
        !same(
          manifest.currentBackups.map((v: any) => v.slideId),
          targetsFor(s, p),
        )
      )
        fail('backup_invalid')
      recoveryBackups = manifest.currentBackups
      for (const row of recoveryBackups!) {
        const bytes = await read(c, row.packageRef)
        if (
          (await awaited(c, () => presentationPackageDigest(encode(bytes), c.signal))) !==
            row.digest ||
          row.digest !== after.pages.find((v) => v.slideId === row.slideId)?.digest
        )
          fail('backup_invalid')
      }
      nextProgress.currentBackups = [...(nextProgress.currentBackups ?? []), ...recoveryBackups!]
      nextProgress.fallbackOriginalIds = [
        ...new Set([
          ...nextProgress.fallbackOriginalIds,
          ...recoveryBackups!
            .filter(
              (v) =>
                v.originalSlideId &&
                v.digest !== before.pages.find((o) => o.slideId === v.slideId)?.digest,
            )
            .map((v) => v.originalSlideId!),
        ]),
      ]
      cmd = { ...cmd, nextState: 'recovery_required', nextPhase: 'verify', nextIndex: 0 }
    }
    if (
      !unchanged(
        before,
        after,
        recoveryBackups
          ? recoveryBackups.map((v) => v.slideId)
          : target && cmd.kind !== 'stage'
            ? [target]
            : [],
      )
    )
      fail('unknown')
    if (cmd.kind === 'stage') {
      const inserted = r.pending!.inserted
      if (!inserted) fail('unknown')
      const order = [...before.slideIds]
      order.splice(order.indexOf(cmd.sourceId!) + 1, 0, inserted!.slideId)
      if (
        !same(order, after.slideIds) ||
        after.dependencies.find((d) => d.slideId === inserted!.slideId)?.masterId !==
          inserted!.masterId ||
        after.dependencies.find((d) => d.slideId === inserted!.slideId)?.layoutId !==
          inserted!.layoutId ||
        before.masters.some(
          (m) =>
            !same(
              m,
              after.masters.find((v) => v.masterId === m.masterId),
            ),
        )
      )
        fail('unknown')
      const expected = await awaited(c, () => inspectMasterXmlPackage(cmd.base64!, c.signal)),
        base64 = actual.pages.find((v) => v.slideId === inserted!.slideId)!.base64!
      await awaited(c, () =>
        assertMasterXmlPagePreserved(
          cmd.base64!,
          base64,
          {
            expectedMasterBase64: cmd.base64!,
            targetMasterPath: expected.sourceMasterPath,
            packageLayoutPath: expected.sourceLayoutPath,
          },
          c.signal,
        ),
      )
      nextProgress.ownedSources.push(structuredClone(inserted!))
    } else if (cmd.kind === 'layout') {
      if (
        !same(before.slideIds, after.slideIds) ||
        !same(before.masters, after.masters) ||
        (!recoveryBackups &&
          !same(
            actual.dependencies.find((v) => v.slideId === target),
            { slideId: target, masterId: cmd.masterId, layoutId: cmd.layoutId },
          ))
      )
        fail('unknown')
      if (cmd.capture) {
        const base64 = actual.pages.find((v) => v.slideId === target)!.base64!
        const number = s.pages.length + 1 + p.sequence
        const packageRef = await save(c, `page-${number}`, decode(base64))
        const reps = nextProgress[`${cmd.capture}Representatives`]
        const existing = reps.find(
          (v) => v.masterId === cmd.masterId && v.layoutId === cmd.layoutId,
        )
        if (existing) existing.packageRef = packageRef
        else reps.push({ masterId: cmd.masterId!, layoutId: cmd.layoutId!, packageRef })
      } else if (cmd.originalId && !recoveryBackups)
        try {
          await provePage(
            c,
            s,
            { ...p, progress: nextProgress },
            cmd.originalId,
            target!,
            !cmd.action.startsWith('restore'),
            actual,
          )
        } catch (error) {
          if (!allowUnproven || c.signal?.aborted || c.originalSignal?.aborted || c.epoch !== epoch)
            throw error
          nextProgress.fallbackOriginalIds = [
            ...new Set([...nextProgress.fallbackOriginalIds, cmd.originalId]),
          ]
          cmd = { ...cmd, nextState: 'recovery_required', nextPhase: 'verify', nextIndex: 0 }
        }
    } else if (cmd.kind === 'remove') {
      if (
        !same(
          before.slideIds.filter((id) => id !== target),
          after.slideIds,
        ) ||
        after.masters.some(
          (m) =>
            !same(
              m,
              before.masters.find((v) => v.masterId === m.masterId),
            ),
        ) ||
        before.masters.some(
          (m) =>
            !after.masters.some((v) => v.masterId === m.masterId) &&
            after.dependencies.some((d) => d.masterId === m.masterId),
        )
      )
        fail('unknown')
      if (cmd.action === 'delete_source')
        nextProgress.pageMappings = nextProgress.pageMappings.map((v) =>
          v.originalSlideId === s.sourceSlideId
            ? { ...v, currentSlideId: r.stagedSource!.slideId }
            : v,
        )
    }
    nextProgress.introducedMasterIds = [
      ...new Set([
        ...nextProgress.introducedMasterIds,
        ...after.masters
          .filter((m) => !s.original.masters.some((v) => v.masterId === m.masterId))
          .map((m) => m.masterId),
      ]),
    ]
    if (cmd.carrierPath && cmd.carrierRef)
      nextProgress.masterCarriers = [
        ...(nextProgress.masterCarriers ?? []),
        {
          mode: cmd.action.startsWith('restore') ? 'restored' : 'imported',
          packageMasterPath: cmd.carrierPath,
          identity: r.pending!.inserted!,
          packageRef: cmd.carrierRef,
        },
      ]
    if (cmd.action === 'restore_stage' || cmd.action === 'restore_stage_master')
      nextProgress.restoredMasterIds = [
        ...new Set([...nextProgress.restoredMasterIds, r.pending!.inserted!.masterId]),
      ]
    if (cmd.action === 'restore_page_stage')
      nextProgress.restorationStaged = {
        originalSlideId: cmd.originalId!,
        identity: r.pending!.inserted!,
      }
    if (cmd.action === 'restore_page_delete') {
      nextProgress.pageMappings = nextProgress.pageMappings.map((m) =>
        m.originalSlideId === cmd.originalId
          ? { ...m, currentSlideId: nextProgress.restorationStaged!.identity.slideId }
          : m,
      )
      nextProgress.restorationStaged = undefined
    }
    if (
      cmd.action === 'restore_page' &&
      cmd.originalId === s.sourceSlideId &&
      !s.original.slideIds.includes(cmd.targetSlideId!)
    )
      nextProgress.pageMappings = nextProgress.pageMappings.map((m) =>
        m.originalSlideId === s.sourceSlideId ? { ...m, currentSlideId: cmd.targetSlideId! } : m,
      )
    if (cmd.action === 'stage' || cmd.action === 'stage_master')
      nextProgress.importedMasterIds = [
        ...new Set([...nextProgress.importedMasterIds, r.pending!.inserted!.masterId]),
      ]
    if (cmd.kind === 'checkpoint' && cmd.originalId && cmd.nextPhase !== 'restore_page_stage')
      await provePage(
        c,
        s,
        { ...p, progress: nextProgress },
        cmd.originalId,
        cmd.targetSlideId!,
        false,
        actual,
      )
    if (cmd.nextState === 'undone' || cmd.nextState === 'discarded') {
      for (const page of s.pages) await pageBytes(c, s, page.slideId)
      const wanted = s.original.slideIds.map(
        (id) => nextProgress.pageMappings.find((m) => m.originalSlideId === id)!.currentSlideId,
      )
      if (!same(wanted, after.slideIds)) fail('unknown')
      if (cmd.nextState === 'discarded' && !nextProgress.originalMappings.length) {
        if (!unchanged(s.original, after, []) || !same(after.slideIds, s.original.slideIds))
          fail('unknown')
      } else {
        const all = await inspect(
          c,
          affected(s, { ...p, progress: nextProgress }).map(
            (page) =>
              nextProgress.pageMappings.find((m) => m.originalSlideId === page.slideId)!
                .currentSlideId,
          ),
        )
        for (const page of affected(s, { ...p, progress: nextProgress }))
          await provePage(
            c,
            s,
            { ...p, progress: nextProgress },
            page.slideId,
            nextProgress.pageMappings.find((m) => m.originalSlideId === page.slideId)!
              .currentSlideId,
            false,
            all,
          )
        if (!same(host(all), after)) fail('unknown')
      }
    }
    const slot = (
      {
        original_probe_stage: 'originalProbeSource',
        stage: 'stagedSource',
        import_probe_stage: 'importedProbeSource',
        restore_stage: 'restoredSource',
        restore_probe_stage: 'restoredProbeSource',
      } as const
    )[cmd.action as 'stage']
    if (slot && !nextProgress.roles[slot]) nextProgress.roles[slot] = r.pending!.inserted!
    const nextProof: Proof = {
      version: 1,
      documentId: c.documentId,
      changeId: c.changeId,
      sequence: p.sequence + 1,
      previous: r.currentProofRef,
      host: after,
      progress: nextProgress,
      action: cmd.action,
      step: {
        kind: cmd.kind,
        index: r.cursor.index,
        ...(cmd.targetSlideId ? { targetSlideId: cmd.targetSlideId } : {}),
        ...(cmd.masterId ? { masterId: cmd.masterId } : {}),
        ...(cmd.layoutId ? { layoutId: cmd.layoutId } : {}),
        ...(cmd.sourceId ? { sourceId: cmd.sourceId } : {}),
        ...(cmd.originalId ? { originalId: cmd.originalId } : {}),
        ...(r.pending!.inserted ? { inserted: r.pending!.inserted } : {}),
        ...(r.pending!.currentBackupRef ? { beforeBackupRef: r.pending!.currentBackupRef } : {}),
        ...(r.pending!.recoveryPreimageRef
          ? { recoveryPreimageRef: r.pending!.recoveryPreimageRef }
          : {}),
      },
    }
    const afterProofRef = await save(
      c,
      `receipt-${nextProof.sequence}`,
      json({
        ...nextProof,
        next: { state: cmd.nextState, phase: cmd.nextPhase, index: cmd.nextIndex },
      }),
    )
    if (!same(host(await inspect(c)), after)) fail('unknown')
    r = await store(c, { ...r, pending: { ...r.pending!, afterProofRef } })
    return finish(
      c,
      r,
      cmd.nextState,
      cmd.nextPhase,
      cmd.nextIndex,
      nextProgress.introducedMasterIds.length,
    )
  }
  const drive = async (c: Context, r: PresentationMasterXmlChange) => {
    while (
      !(r.state === 'applied' && r.cursor.phase === 'verify') &&
      r.state !== 'undone' &&
      r.state !== 'discarded'
    ) {
      if (
        r.pending ||
        (r.state === 'recovery_required' &&
          r.cursor.phase !== 'restore_stage' &&
          r.cursor.phase !== 'discard' &&
          !r.cursor.phase.startsWith('restore_page'))
      )
        fail('unknown')
      const { s, p } = await checked(c, r)
      const cmd = await command(c, r, s, p)
      r = await perform(c, r, s, p, cmd)
    }
    return r
  }
  const beginMode = async (
    c: Context,
    r: PresentationMasterXmlChange,
    action: 'undo' | 'discard',
  ) => {
    const { s, p } = await checked(c, r)
    if (action === 'undo') for (const page of s.pages) await pageBytes(c, s, page.slideId)
    r = await store(c, { ...r, cursor: { phase: 'verify', index: 0, substep: 0 }, reviews: [] })
    p.progress.mode = action === 'undo' && !r.stagedSource ? 'discard' : action
    const phase =
      p.progress.mode === 'discard'
        ? 'discard'
        : p.progress.restoredMappings.length
          ? 'restore_page'
          : 'restore_stage'
    return perform(c, r, s, p, {
      action: 'verify',
      kind: 'checkpoint',
      targetSlideId: p.progress.pageMappings[0]!.currentSlideId,
      nextState: r.state,
      nextPhase: phase,
      nextIndex: 0,
    })
  }
  const observed = async (c: Context, r: PresentationMasterXmlChange) => {
    const { s, p } = await load(c, r),
      actual = await inspect(c)
    if (!r.pending)
      return { status: same(host(actual), p.host) ? 'ready' : 'unknown', s, p, actual }
    if (r.pending.afterProofRef) {
      const after = (await readJson(c, r.pending.afterProofRef)) as Proof & {
        next: { state: MasterXmlState; phase: MasterXmlAction; index: number }
      }
      if (
        !validateProof(after, s, r.receiptCount + 1) ||
        after.action !== r.pending.action ||
        !same(after.step?.beforeBackupRef, r.pending.currentBackupRef) ||
        !same(after.step?.recoveryPreimageRef, r.pending.recoveryPreimageRef) ||
        after.step?.index !== r.pending.index ||
        after.sequence !== r.receiptCount + 1 ||
        !same(after.previous, r.currentProofRef) ||
        !validHost(after.host) ||
        !after.next
      )
        fail('backup_invalid')
      validateEdge(p, after, s)
      return { status: same(host(actual), after.host) ? 'after' : 'unknown', s, p, actual, after }
    }
    if (same(host(actual), p.host)) return { status: 'before', s, p, actual }
    const cmd = await command(c, r, s, p)
    let inserted = r.pending.inserted
    if (cmd.kind === 'stage' && !inserted) {
      const index = p.host.slideIds.indexOf(cmd.sourceId!),
        id = actual.slideIds[index + 1],
        newIds = actual.slideIds.filter((id) => !p.host.slideIds.includes(id))
      if (
        id &&
        newIds.length === 1 &&
        newIds[0] === id &&
        same(
          actual.slideIds.filter((v) => v !== id),
          p.host.slideIds,
        ) &&
        unchanged(p.host, host(actual), [])
      ) {
        const dep = actual.dependencies.find((d) => d.slideId === id)
        if (dep) {
          const candidate = await inspect(c, [id]),
            inventory = await awaited(c, () => inspectMasterXmlPackage(cmd.base64!, c.signal))
          try {
            await awaited(c, () =>
              assertMasterXmlPagePreserved(
                cmd.base64!,
                candidate.pages.find((v) => v.slideId === id)!.base64!,
                {
                  expectedMasterBase64: cmd.base64!,
                  targetMasterPath: inventory.sourceMasterPath,
                  packageLayoutPath: inventory.sourceLayoutPath,
                },
                c.signal,
              ),
            )
            inserted = structuredClone(dep)
          } catch (error) {
            if (c.signal?.aborted || c.originalSignal?.aborted || c.epoch !== epoch) throw error
          }
        }
      }
    }
    if (cmd.kind === 'stage')
      return { status: inserted ? 'candidate' : 'unknown', s, p, actual, cmd, inserted }
    if (cmd.kind === 'remove') {
      const exact =
        same(
          actual.slideIds,
          p.host.slideIds.filter((id) => id !== cmd.targetSlideId),
        ) && unchanged(p.host, host(actual), [cmd.targetSlideId!])
      return { status: exact ? 'candidate' : 'unknown', s, p, actual, cmd }
    }
    if (cmd.kind === 'layout') {
      const exact =
        same(actual.slideIds, p.host.slideIds) &&
        same(actual.masters, p.host.masters) &&
        unchanged(p.host, host(actual), [cmd.targetSlideId!]) &&
        same(
          actual.dependencies.find((d) => d.slideId === cmd.targetSlideId),
          { slideId: cmd.targetSlideId, masterId: cmd.masterId, layoutId: cmd.layoutId },
        )
      const owned = targetsFor(s, p)
      const broad =
        !exact &&
        same(actual.slideIds, p.host.slideIds) &&
        same(actual.masters, p.host.masters) &&
        unchanged(p.host, host(actual), owned) &&
        actual.dependencies.every((d) =>
          actual.masters.some(
            (m) => m.masterId === d.masterId && m.layouts.some((l) => l.layoutId === d.layoutId),
          ),
        )
      return {
        status: exact ? 'candidate' : broad ? 'recovery_candidate' : 'unknown',
        s,
        p,
        actual,
        cmd,
      }
    }
    return { status: 'unknown', s, p, actual, cmd }
  }
  const targetsFor = (s: Snapshot, p: Proof) => [
    ...new Set(
      [
        ...affected(s, p).map(
          (v) =>
            p.progress.pageMappings.find((m) => m.originalSlideId === v.slideId)!.currentSlideId,
        ),
        ...p.progress.ownedSources.map((v) => v.slideId),
      ].filter((id) => p.host.slideIds.includes(id)),
    ),
  ]
  const recoveryProposal = async (
    r: PresentationMasterXmlChange,
    action: 'resume' | 'undo' | 'discard' | 'reconcile',
    signal?: AbortSignal,
  ) => {
    const c = context(r, signal),
      value = await observed(c, r)
    if (
      action === 'reconcile'
        ? !r.pending || ['ready', 'unknown'].includes(value.status)
        : value.status !== 'ready'
    )
      fail('unknown')
    if (
      action === 'undo' &&
      ![
        'applied',
        'applying',
        'restoring',
        'probing_restored',
        'restoring_pages',
        'recovery_required',
      ].includes(r.state)
    )
      fail('state_invalid')
    if (
      action === 'discard' &&
      !['prepared', 'probing_original', 'staged', 'probing_imported'].includes(r.state)
    )
      fail('state_invalid')
    if (
      action === 'resume' &&
      ['applied', 'undone', 'discarded', 'recovery_required'].includes(r.state)
    )
      fail('state_invalid')
    let recoveryPreimageRef: PackageBackupRef | undefined
    if (action === 'reconcile' && value.status === 'recovery_candidate') {
      const currentBackups: NonNullable<Progress['currentBackups']> = []
      for (const id of targetsFor(value.s, value.p))
        currentBackups.push(await captureCurrent(c, value.s, value.p, id, host(value.actual)))
      recoveryPreimageRef = await save(
        c,
        `page-8${BigInt(
          '0x' +
            Array.from(
              new Uint8Array(
                await awaited(c, () =>
                  crypto.subtle.digest(
                    'SHA-256',
                    json({ host: host(value.actual), currentBackups }).buffer as ArrayBuffer,
                  ),
                ),
              ),
            )
              .map((v) => v.toString(16).padStart(2, '0'))
              .join(''),
        )
          .toString()
          .padStart(78, '0')}`,
        json({
          version: 1,
          documentId: c.documentId,
          changeId: c.changeId,
          host: host(value.actual),
          currentBackups,
        }),
      )
    }
    const targetIds = targetsFor(value.s, value.p),
      metadata =
        action === 'reconcile' ||
        (action === 'discard' &&
          r.state === 'prepared' &&
          !value.p.progress.ownedSources.some((v) => value.p.host.slideIds.includes(v.slideId)))
    const operation = `${action}_master_xml_change`,
      fingerprint = JSON.stringify({
        record: r,
        observed: host(value.actual),
        status: value.status,
      })
    return options.proposals.propose({
      operation,
      toolName: operation,
      title:
        action === 'undo' && r.state === 'recovery_required'
          ? 'Restore saved original pages over observed unproven owned content'
          : `${action} master XML change`,
      preview: {
        ...summary(r),
        ...(recoveryPreimageRef
          ? {
              warning:
                'Explicit recovery preserves all observed affected page bytes and authorizes original-package restoration over changed owned content.',
              currentPreimageRef: recoveryPreimageRef,
              coveredPageCount: targetIds.length,
              changedOwnedPageCount: value.actual.pages.filter(
                (v) =>
                  targetIds.includes(v.slideId) &&
                  v.digest !== value.p.host.pages.find((o) => o.slideId === v.slideId)?.digest,
              ).length,
            }
          : {}),
        ...(r.state === 'recovery_required'
          ? {
              warning:
                'Confirmed original restoration overwrites the current content on the proven owned pages. Introduced unused master inventory cleanup remains unverified.',
              originalSnapshotRef: r.snapshotRef,
              currentObservedProofRef: r.currentProofRef,
            }
          : {}),
        qaScope: { basis: 'master_xml_savepoint', hostSlideIds: targetIds },
      },
      impact: {
        host: metadata ? 'local_checkpoint' : 'powerpoint',
        targets: targetIds,
        count: targetIds.length,
      },
      fingerprint: selectionFingerprint(fingerprint),
      validate: async (nextSignal) => {
        try {
          const now = await observed({ ...c, signal: nextSignal }, r)
          return (
            JSON.stringify({ record: r, observed: host(now.actual), status: now.status }) ===
            fingerprint
          )
        } catch {
          return false
        }
      },
      execute: async (nextSignal) => {
        const ctx = { ...c, signal: nextSignal }
        const now = await observed(ctx, r)
        if (
          JSON.stringify({ record: r, observed: host(now.actual), status: now.status }) !==
          fingerprint
        )
          fail('stale')
        if (action === 'reconcile') {
          if (now.status === 'before') {
            await store(ctx, {
              ...r,
              pending: undefined,
              state: r.state === 'prepared' ? 'discarded' : 'recovery_required',
            })
            return
          }
          if (now.after) {
            await finish(
              ctx,
              r,
              now.after.next.state,
              now.after.next.phase,
              now.after.next.index,
              now.after.progress.introducedMasterIds.length,
            )
            return
          }
          let current = r
          if (now.inserted && !current.pending!.inserted)
            current = await store(ctx, {
              ...current,
              pending: { ...current.pending!, inserted: now.inserted },
            })
          if (recoveryPreimageRef)
            current = await store(ctx, {
              ...current,
              pending: { ...current.pending!, recoveryPreimageRef },
            })
          await receive(ctx, current, now.s, now.p, now.cmd!, now.p.host, true)
          return
        }
        if (action === 'discard' && metadata) {
          await store(ctx, { ...r, state: 'discarded' })
          return
        }
        let current = r
        if (action === 'undo' || action === 'discard')
          current = await beginMode(ctx, current, action)
        await drive(ctx, current)
      },
      verify: async (nextSignal) => {
        const current = record(r.changeId)
        await checked(context(current, nextSignal), current)
      },
    })
  }
  const initialProposal = async (
    replacements: XmlReplacement[],
    explanation?: string,
    signal?: AbortSignal,
  ) => {
    const owned = structuredClone(replacements),
      intent = explanation ?? 'Edit PowerPoint master XML'
    if (
      !Array.isArray(owned) ||
      owned.length < 1 ||
      owned.length > 32 ||
      json({ version: 1, operations: owned.map((r) => ({ op: 'replace_xml', ...r })) }).length >
        32 * 1024 ||
      typeof intent !== 'string' ||
      intent.length > 8192
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
    const originalHost = host(await inspect(c)),
      sourceSlideId = originalHost.slideIds[0]!
    const source = await awaited(c, async () =>
      structuredClone(await options.adapter.readPage(sourceSlideId, c.signal)),
    )
    if (
      source.slideId !== sourceSlideId ||
      source.digest !== originalHost.pages[0]!.digest ||
      !same(
        { slideId: source.slideId, masterId: source.masterId, layoutId: source.layoutId },
        originalHost.dependencies[0],
      )
    )
      fail('drift')
    const originalBase64 = source.base64,
      edited = await awaited(c, () =>
        editPowerPointPackage(originalBase64, 'master', owned, c.signal),
      ),
      preparation = await awaited(c, () =>
        assertMasterXmlPreparation(originalBase64, edited.base64, edited.changedPaths, c.signal),
      )
    const { base64: preparedBase64, ...expected } = edited
    const pageIdentity: { slideId: string; masterId: string; layoutId: string }[] = []
    const nativeGraph = new Map<string, string>()
    for (let i = 0; i < originalHost.slideIds.length; i++) {
      const slideId = originalHost.slideIds[i]!,
        actual =
          i === 0
            ? source
            : await awaited(c, async () =>
                structuredClone(await options.adapter.readPage(slideId, c.signal)),
              )
      if (
        actual.slideId !== slideId ||
        actual.digest !== originalHost.pages[i]!.digest ||
        !same(
          { slideId, masterId: actual.masterId, layoutId: actual.layoutId },
          originalHost.dependencies[i],
        )
      )
        fail('drift')
      const inventory = await awaited(c, () => inspectMasterXmlPackage(actual.base64, c.signal)),
        selected = inventory.masters.find((m) => m.path === inventory.sourceMasterPath)!
      if (
        nativeGraph.has(actual.masterId) &&
        nativeGraph.get(actual.masterId) !== selected.contentDigest
      )
        fail('mapping_unproven')
      nativeGraph.set(actual.masterId, selected.contentDigest)
      pageIdentity.push({ slideId, masterId: actual.masterId, layoutId: actual.layoutId })
    }
    if (!same(host(await inspect(c)), originalHost)) fail('drift')
    const affectedDigests = preparation.original.masters
        .filter((m) => preparation.affectedMasterPaths.includes(m.path))
        .map((m) => m.contentDigest),
      affectedPageCount = pageIdentity.filter(
        (p) =>
          p.slideId === sourceSlideId || affectedDigests.includes(nativeGraph.get(p.masterId)!),
      ).length
    const targets = pageIdentity
      .filter(
        (p) =>
          p.slideId === sourceSlideId || affectedDigests.includes(nativeGraph.get(p.masterId)!),
      )
      .map((p) => p.slideId)
    const attemptedRefs: PackageBackupRef[] = []
    const trackedSave = async (ctx: Context, key: string, bytes: Uint8Array) => {
      const expectedRef = await packageBackupRefForBytes(key, bytes)
      attemptedRefs.push(expectedRef)
      const saved = await save(ctx, key, bytes)
      if (!same(saved, expectedRef)) fail('backup_invalid')
      return saved
    }
    const prepareConfirmed = async (ctx: Context) => {
      if (!same(host(await inspect(ctx)), originalHost)) fail('drift')
      const preparedRef = await trackedSave(ctx, 'page-0', decode(preparedBase64)),
        pages: OriginalPage[] = [],
        representatives: Progress['originalRepresentatives'] = []
      for (let i = 0; i < pageIdentity.length; i++) {
        const identity = pageIdentity[i]!,
          actual = await awaited(ctx, async () =>
            structuredClone(await options.adapter.readPage(identity.slideId, ctx.signal)),
          )
        if (
          actual.slideId !== identity.slideId ||
          actual.digest !== originalHost.pages[i]!.digest ||
          actual.masterId !== identity.masterId ||
          actual.layoutId !== identity.layoutId
        )
          fail('drift')
        const packageRef = await trackedSave(ctx, `page-${i + 1}`, decode(actual.base64))
        pages.push({ ...identity, packageRef })
        if (
          !representatives.some(
            (v) => v.masterId === identity.masterId && v.layoutId === identity.layoutId,
          )
        )
          representatives.push({
            masterId: identity.masterId,
            layoutId: identity.layoutId,
            packageRef,
          })
      }
      if (!same(host(await inspect(ctx)), originalHost)) fail('drift')
      const s: Snapshot = {
        version: 1,
        documentId,
        changeId: c.changeId,
        sourceSlideId,
        packageSourceSlideId: preparation.original.sourceSlideId,
        original: originalHost,
        pages,
        carrierRef: pages[0]!.packageRef,
        preparedRef,
        originalInventory: preparation.original,
        preparedInventory: preparation.prepared,
        affectedMasterPaths: preparation.affectedMasterPaths,
        replacements: owned,
        expected,
      }
      const snapshotRef = await trackedSave(ctx, 'snapshot', json(s)),
        progress: Progress = {
          originalRepresentatives: representatives,
          originalMappings: [],
          importedRepresentatives: [],
          importedMappings: [],
          restoredRepresentatives: [],
          restoredMappings: [],
          ownedSources: [],
          roles: {},
          mode: 'forward',
          fallbackOriginalIds: [],
          pageMappings: pages.map((p) => ({
            originalSlideId: p.slideId,
            currentSlideId: p.slideId,
          })),
          importedMasterIds: [],
          restoredMasterIds: [],
          introducedMasterIds: [],
        }
      const initial: Proof = {
          version: 1,
          documentId,
          changeId: c.changeId,
          sequence: 0,
          host: originalHost,
          progress,
        },
        currentProofRef = await trackedSave(ctx, 'receipt-0', json(initial))
      const r: PresentationMasterXmlChange = {
        version: 1,
        kind: 'master_xml',
        documentId,
        changeId: c.changeId,
        intent,
        sourceSlideId,
        packageSourceSlideId: s.packageSourceSlideId,
        snapshotRef,
        preparedRef,
        currentProofRef,
        state: 'prepared',
        cursor: { phase: 'original_probe_stage', index: 0, substep: 0 },
        receiptCount: 0,
        scope: {
          originalMasterId: source.masterId,
          affectedPageCount,
          originalLayoutCount: originalHost.masters.find((m) => m.masterId === source.masterId)!
            .layouts.length,
          ...(s.affectedMasterPaths.length
            ? { affectedMasterCount: s.affectedMasterPaths.length }
            : {}),
        },
        introducedMasterCount: 0,
        inventoryCleanupVerified: false,
        reviews: [],
      }
      if (!validatePresentationMasterXmlChange(r)) fail('state_invalid')
      await load(ctx, r)
      if (!same(host(await inspect(ctx)), originalHost)) fail('drift')
      return r
    }
    return options.proposals.propose({
      operation: 'edit_slide_master_xml',
      toolName: 'edit_slide_master_xml',
      title: intent,
      preview: {
        changeId: c.changeId,
        state: 'prepared',
        sourceSlideId,
        qaPassed: false,
        changedPaths: edited.changedPaths,
        qaScope: { basis: 'master_xml_savepoint', hostSlideIds: targets },
      },
      impact: { host: 'powerpoint', targets, count: targets.length },
      fingerprint: selectionFingerprint(JSON.stringify(originalHost)),
      validate: async (nextSignal) => {
        try {
          const ctx = { ...c, signal: nextSignal }
          return same(host(await inspect(ctx)), originalHost)
        } catch (error) {
          if (nextSignal?.aborted || c.originalSignal?.aborted || c.epoch !== epoch) throw error
          return false
        }
      },
      execute: async (nextSignal) => {
        const ctx = { ...c, signal: nextSignal }
        let persistenceAttempted = false
        try {
          const r = await prepareConfirmed(ctx)
          persistenceAttempted = true
          await store(ctx, r)
          await drive(ctx, r)
        } catch (error) {
          if (!persistenceAttempted)
            await Promise.allSettled(
              attemptedRefs.map((backup) =>
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
      verify: async (nextSignal) => {
        const current = record(c.changeId)
        await checked(context(current, nextSignal), current)
      },
    })
  }
  return {
    id: 'presentation-master-xml',
    systemPrompt:
      'Inspect durable master XML recovery before writes. Never replay unresolved native writes, infer layouts from names or order, or certify whole-deck QA from screenshots.',
    tools,
    clear() {
      epoch++
      verifiedTip = undefined
      invalidateVisual()
    },
    beginMutation: invalidateVisual,
    endMutation: invalidateVisual,
    propose: initialProposal,
    async executeTool(inputCall: any, signal?: AbortSignal): Promise<ToolExecutionOutcome> {
      try {
        const call = structuredClone(inputCall),
          action = actions.find((a) => call.name === `${a}_master_xml_change`),
          fields =
            action === 'review'
              ? ['change_id', 'slide_id', 'screenshot_digest', 'status', 'notes']
              : action === 'capture'
                ? ['change_id', 'slide_id']
                : ['change_id']
        if (
          !action ||
          !call.input ||
          typeof call.input !== 'object' ||
          Object.keys(call.input).some((k) => !fields.includes(k)) ||
          fields.some((k) => !Object.hasOwn(call.input, k)) ||
          typeof call.input.change_id !== 'string'
        )
          throw Error('invalid_tool_input')
        const r = record(call.input.change_id),
          c = context(r, signal)
        if (action === 'inspect') {
          const value = await observed(c, r)
          return {
            output: JSON.stringify({
              ...summary(r),
              observed: value.status,
              scope: targetsFor(value.s, value.p),
            }),
            mutated: false,
            summary: 'Inspected durable master XML change',
          }
        }
        if (action === 'capture' || action === 'review') {
          const { s, p } = await load(c, r),
            id = call.input.slide_id,
            before = host(await inspect(c))
          if (
            !validPackageHostId(id) ||
            !before.slideIds.includes(id) ||
            ![
              ...s.original.slideIds,
              ...p.progress.ownedSources.map((v) => v.slideId),
              r.pending?.inserted?.slideId,
            ].includes(id) ||
            !options.adapter.screenshotSlide
          )
            throw Error('invalid_tool_input')
          const capturedEpoch = visualEpoch,
            index = before.slideIds.indexOf(id),
            screenshot = await awaited(c, async () =>
              structuredClone(await options.adapter.screenshotSlide!(index, c.signal)),
            ),
            base64 = validatePowerPointPageScreenshot(screenshot.base64)
          if (screenshot.slideId !== id || screenshot.mime !== 'image/png') fail('qa_stale')
          const digest = Array.from(
            new Uint8Array(
              await awaited(c, () =>
                crypto.subtle.digest('SHA-256', decode(base64).buffer as ArrayBuffer),
              ),
            ),
            (b) => b.toString(16).padStart(2, '0'),
          ).join('')
          if (!same(host(await inspect(c)), before) || capturedEpoch !== visualEpoch)
            fail('qa_stale')
          const key = `${r.changeId}:${id}`
          if (action === 'capture') {
            captures.set(key, { digest, record: JSON.stringify(r), epoch: visualEpoch })
            return {
              output: JSON.stringify({
                slideId: id,
                screenshot: base64,
                screenshotDigest: digest,
                qaPassed: false,
              }),
              mutated: false,
              summary: 'Captured fresh native master XML page screenshot',
            }
          }
          const captured = captures.get(key)
          if (
            r.pending ||
            !['applied', 'undone'].includes(r.state) ||
            !captured ||
            captured.digest !== digest ||
            captured.digest !== call.input.screenshot_digest ||
            captured.record !== JSON.stringify(r) ||
            captured.epoch !== visualEpoch ||
            !['pass', 'fail'].includes(call.input.status) ||
            typeof call.input.notes !== 'string' ||
            call.input.notes.length > 8192
          )
            fail('qa_stale')
          const keyNumber = (r.reviewSequence ?? 0) + 1 + 1000000,
            reviewRef = await save(
              c,
              `image-${keyNumber}`,
              json({
                slideId: id,
                screenshotDigest: digest,
                status: call.input.status,
                notes: call.input.notes,
                proofRef: r.currentProofRef,
                capturedAt: new Date().toISOString(),
                qaPassed: false,
              }),
            )
          if (!same(host(await inspect(c)), before) || capturedEpoch !== visualEpoch)
            fail('qa_stale')
          await store(c, {
            ...r,
            reviewSequence: (r.reviewSequence ?? 0) + 1,
            reviews: [...r.reviews.filter((v) => v.slideId !== id), { slideId: id, reviewRef }],
          })
          captures.delete(key)
          return {
            output: JSON.stringify({ slideId: id, status: call.input.status, qaPassed: false }),
            mutated: false,
            summary: 'Recorded historical master XML page review',
          }
        }
        const proposal = await recoveryProposal(r, action, signal)
        return {
          output: JSON.stringify(proposal),
          mutated: false,
          summary: `Proposed master XML ${action}`,
        }
      } catch (error) {
        return {
          output: JSON.stringify({ error: safeError(error) }),
          mutated: false,
          isError: true,
          summary: 'Master XML operation refused',
        }
      }
    },
  } satisfies AgentSkill & {
    propose: typeof initialProposal
    clear(): void
    beginMutation(): void
    endMutation(): void
  }
}
