import { createPresentationResearchSkill } from '../skills/powerpoint/presentation-research.js'
import {
  createPresentationResearchController,
  type PresentationResearchController,
  type PresentationResearchDeleteAttempt,
} from './presentation-research.js'
import { createPresentationHostBundleSkill } from '../skills/powerpoint/presentation-host-bundle.js'
import {
  supportsPowerPointDocumentExport,
  exportPowerPointDocument,
} from '../skills/powerpoint/presentation-document-export.js'
import type { PresentationAcquisitionHistory } from '@wiswork/project-store/presentation-acquisition'
import { presentationMutationScope } from '../skills/powerpoint/presentation-mutation-scope.js'
import { readPresentationNativeLocks } from '../skills/powerpoint/presentation-native-locks.js'
import { createPresentationExistingEditingSkill } from '../skills/powerpoint/presentation-existing-editing.js'
import { createPresentationExistingBatchEditingSkill } from '../skills/powerpoint/presentation-existing-batch-editing.js'
import { createPresentationExistingImageEditingSkill } from '../skills/powerpoint/presentation-existing-image-editing.js'
import { createPresentationExistingPageEditingSkill } from '../skills/powerpoint/presentation-existing-page-editing.js'
import type { PresentationExistingImageChange } from '../skills/powerpoint/presentation-existing-image.js'
import type { PresentationExistingPageChange } from '../skills/powerpoint/presentation-existing-page.js'
import type { PresentationExistingChartChange } from '../skills/powerpoint/presentation-existing-chart.js'
import type { PresentationExistingBatch } from '../skills/powerpoint/presentation-existing-batch.js'
import type { PresentationExistingChange } from '../skills/powerpoint/presentation-existing-change.js'
import { createPresentationBaselineSkill } from '../skills/powerpoint/presentation-baseline.js'
import { BrowserPresentationBaselineAdapter } from '../skills/powerpoint/browser-presentation-baseline-adapter.js'
import { createBrowserPresentationDocumentBinding } from '../skills/powerpoint/presentation-document.js'
import { createPresentationHistorySkill } from '../skills/powerpoint/presentation-history.js'
import type { PresentationHistoryEntry } from '../skills/powerpoint/presentation-change-history.js'
import { createPresentationImageBackup } from '../skills/powerpoint/presentation-image-backup.js'
import {
  createPresentationChangesController,
  type PresentationChangesController,
} from './presentation-changes.js'
import type { PresentationTextChange } from '../skills/powerpoint/presentation-text-change.js'
import { createPresentationEvidenceDeliverySkill } from '../skills/powerpoint/presentation-evidence-delivery.js'
import { createPresentationJobsSkill } from '../skills/powerpoint/presentation-jobs.js'
import type { PresentationPageReplacement } from '../skills/powerpoint/presentation-page-replacement-record.js'
import { createPresentationPageReplacementSkill } from '../skills/powerpoint/presentation-page-replacement.js'
import { BrowserPresentationPageReplacementAdapter } from '../skills/powerpoint/browser-presentation-page-replacement-adapter.js'
import { createPresentationPageBackupSkill } from '../skills/powerpoint/presentation-page-backup.js'
import { createPresentationPageScreenshotFallback } from '../skills/powerpoint/presentation-page-screenshot-fallback.js'
import { presentationPackageDigest } from '../skills/powerpoint/powerpoint-package.js'
import type { PresentationGeometryChange } from '../skills/powerpoint/presentation-geometry-change.js'
import { createPresentationProductionSkill } from '../skills/powerpoint/presentation-production.js'
import { BrowserPresentationImageAdapter } from '../skills/powerpoint/browser-presentation-image-adapter.js'
import type { ImageReplacementRecord } from '../skills/powerpoint/presentation-image-replacement-record.js'
import { createPresentationPageEditingSkill } from '../skills/powerpoint/presentation-page-editing.js'
import {
  createPresentationQaSkill,
  type PresentationQaRecord,
} from '../skills/powerpoint/presentation-qa.js'
import type { PresentationQaController } from './presentation-qa-card.js'
import {
  createPresentationProductionDeliverySkill,
  presentationImportKey,
  summarizePresentationImport,
} from '../skills/powerpoint/presentation-page-delivery.js'
import type { PresentationImportProgressController } from './presentation-import-progress.js'
import {
  createPresentationAttachmentSkill,
  supportsPresentationAttachment,
  type PresentationAttachmentMetadata,
} from '../skills/powerpoint/presentation-attachments.js'
import { createPresentationPlanningSkill } from '../skills/powerpoint/presentation-planning.js'
import { createPresentationCommentsSkill } from '../skills/powerpoint/presentation-comments.js'
import {
  createPresentationProjectController,
  type PresentationProjectController,
} from '../skills/powerpoint/presentation-project.js'
import {
  createPresentationDeliverySkill,
  type PresentationImportRecord,
} from '../skills/powerpoint/presentation-delivery.js'
import { createBrowserPresentationImportAdapter } from '../skills/powerpoint/presentation-import.js'
import type { AgentSkill } from '@wiswork/agent-core'
import type { OfficeDiagnostics } from '../diagnostics/office-diagnostics.js'
import {
  createOfficeDocumentClient,
  createBrowserOfficeRuntime,
  type OfficeDocumentClient,
  type OfficeHost,
} from '../office-document.js'
import { BrowserExcelAdapter } from '../skills/excel/browser-excel-adapter.js'
import {
  BrowserExcelImportMediaAdapter,
  supportsExcelImportMedia,
} from '../skills/excel/browser-excel-import-media-adapter.js'
import { createExcelImportMediaSkill } from '../skills/excel/excel-import-media.js'
import { createExcelSkill } from '../skills/excel/excel-skill.js'
import { BrowserPowerPointAdapter } from '../skills/powerpoint/browser-powerpoint-adapter.js'
import {
  BrowserPowerPointImportMediaAdapter,
  supportsPowerPointImportMedia,
} from '../skills/powerpoint/browser-powerpoint-import-media-adapter.js'
import { createPowerPointImportMediaSkill } from '../skills/powerpoint/powerpoint-import-media.js'
import {
  createPresentationGenerationSkill,
  type PresentationGenerationOptions,
} from '../skills/powerpoint/presentation-generation.js'
import { createPowerPointSkill } from '../skills/powerpoint/powerpoint-skill.js'
import { createSharedBrowserSkill } from '../skills/shared/shared-skill.js'
import { supportsBrowserMediaValidation } from '../skills/shared/import-media.js'
import { MAX_SKILL_BYTES, SkillRegistry } from '../skills/shared/skill-registry.js'
import { SkillPackageWorkerRuntime } from '../skills/shared/skill-package-runtime.js'
import { InMemoryVfs, MAX_VFS_FILE_BYTES } from '../skills/shared/vfs.js'
import { BrowserWordAdapter } from '../skills/word/browser-word-adapter.js'
import { createWordSkill } from '../skills/word/word-skill.js'
import { createOfficeSkill } from './office-skill.js'
import {
  createProposalController,
  createStructuredProposalController,
  type ProposalController,
  type StructuredProposalController,
} from './proposal-controller.js'
import { composeOfficeSkills } from './skill-registry.js'

export interface OfficeHostRuntime {
  readonly research?: PresentationResearchController
  readonly presentation?: PresentationProjectController
  readonly importProgress?: PresentationImportProgressController
  readonly qa?: PresentationQaController
  readonly changes?: PresentationChangesController
  durableAttachmentsAvailable?(): boolean
  durableImagesAvailable?(): boolean
  remoteImagesAvailable?(): boolean
  webpagesAvailable?(): boolean
  rightsAvailable?(): boolean
  animationFrameAvailable?(): boolean
  readPresentationAcquisitionHistory?(): Promise<PresentationAcquisitionHistory | undefined>
  listDurableAttachments?(): Promise<PresentationAttachmentMetadata[]>
  deleteDurableAttachment?(attachmentId: string): Promise<void>
  importPresentationImageUrl?(url: string | string[]): Promise<void>
  importPresentationWebpageUrl?(url: string): Promise<void>
  attestPresentationImageLicense?(
    imageId: string,
    license: 'owned' | 'licensed' | 'public_domain',
    evidenceId: string,
  ): Promise<void>
  revokePresentationImageLicense?(imageId: string): Promise<void>
  extractPresentationImageFirstFrame?(imageId: string): Promise<void>
  skill: AgentSkill
  proposals: ProposalController | StructuredProposalController
  vfs: InMemoryVfs
  skills: SkillRegistry
  readonly skillPackagesEnabled: boolean
  uploadFile(name: string, content: Promise<ArrayBuffer>): Promise<void>
  installSkill(source: Promise<string>): Promise<void>
  installSkillPackage(source: Promise<ArrayBuffer>, signal?: AbortSignal): Promise<void>
  uninstallSkill(name: string): void
  clearSession(): void
  dispose(): void
}

function currentOfficePlatform(): string | undefined {
  try {
    return typeof Office.context.platform === 'string' ? Office.context.platform : undefined
  } catch {
    return undefined
  }
}

function supportsNativePowerPointMasterEditing(): boolean {
  try {
    return Office.context.requirements.isSetSupported('PowerPointApi', '1.10')
  } catch {
    return false
  }
}

export function createOfficeHostRuntime(
  host: OfficeHost,
  options: {
    presentation?: Omit<PresentationGenerationOptions, 'vfs'> & {
      readResearchDeleteAttempt?(documentId: string): unknown
      writeResearchDeleteAttempt?(
        documentId: string,
        value: PresentationResearchDeleteAttempt | undefined,
      ): void
      readExistingChange?(changeId: string): PresentationExistingChange | undefined
      readExistingBatch?(changeId: string): PresentationExistingBatch | undefined
      readExistingImageChange?(changeId: string): PresentationExistingImageChange | undefined
      readExistingPageChange?(changeId: string): PresentationExistingPageChange | undefined
      readExistingChartChange?(changeId: string): PresentationExistingChartChange | undefined
      writeExistingChartChange?(
        record: PresentationExistingChartChange,
        expected: PresentationExistingChartChange | undefined,
      ): Promise<void>
      writeExistingPageChange?(
        record: PresentationExistingPageChange,
        expected: PresentationExistingPageChange | undefined,
      ): Promise<void>
      writeExistingImageChange?(
        record: PresentationExistingImageChange,
        expected: PresentationExistingImageChange | undefined,
      ): Promise<void>
      writeExistingBatch?(
        record: PresentationExistingBatch,
        expected: PresentationExistingBatch | undefined,
      ): Promise<void>
      writeExistingChange?(
        record: PresentationExistingChange,
        expected: PresentationExistingChange | undefined,
      ): Promise<void>
      readPageReplacement?(): PresentationPageReplacement | undefined
      writePageReplacement?(
        record: PresentationPageReplacement,
        expected: PresentationPageReplacement | undefined,
      ): Promise<void>
      listChangeHistory?(): PresentationHistoryEntry[]
      readTextChange?(changeId?: string): PresentationTextChange | undefined
      writeTextChange?(
        record: PresentationTextChange,
        expected: PresentationTextChange | undefined,
      ): Promise<void>
      listImageReplacements?(): ImageReplacementRecord[]
      readGeometryChange?(changeId?: string): PresentationGeometryChange | undefined
      writeGeometryChange?(
        record: PresentationGeometryChange,
        expected: PresentationGeometryChange | undefined,
      ): Promise<void>
      readImageReplacement?(key: string): ImageReplacementRecord | undefined
      writeImageReplacement?(key: string, record: ImageReplacementRecord): Promise<void>
      invalidateQa?(hostSlideIds?: readonly string[]): Promise<void>
      readQa?(key: string): PresentationQaRecord | undefined
      writeQa?(key: string, record: PresentationQaRecord): Promise<void>
      readReceipt?(key: string): PresentationImportRecord | undefined
      listReceipts?(): { key: string; record: PresentationImportRecord }[]
      writeReceipt?(key: string, record: PresentationImportRecord | undefined): Promise<void>
    }
    enableHostSkills?: boolean
    document?: OfficeDocumentClient
    packageRuntime?: Pick<SkillPackageWorkerRuntime, 'parse' | 'cancelAll'>
    enableConversions?: boolean
    enableSkillPackages?: boolean
    enableImportMedia?: boolean
    /** Internal test seam for the legacy image adapter; production callers leave this unset. */
    imageAdapterOverrideForTests?: boolean
    platform?: string
    diagnostics?: Pick<OfficeDiagnostics, 'setTool' | 'record'>
  } = {},
): OfficeHostRuntime {
  if (host === 'unknown') throw new Error('office_host_unsupported')
  let changes: PresentationChangesController | undefined
  const vfs = new InMemoryVfs()
  const skills = new SkillRegistry(vfs)
  if (options.enableHostSkills === false) {
    const document = options.document ?? createOfficeDocumentClient(createBrowserOfficeRuntime())
    const proposals = createProposalController(document, options.diagnostics)
    return lifecycle(
      createOfficeSkill(document, proposals),
      proposals,
      vfs,
      skills,
      options.packageRuntime,
    )
  }
  const localBinding =
    host === 'powerpoint'
      ? (options.presentation ?? createBrowserPresentationDocumentBinding())
      : undefined
  let existingEditing: ReturnType<typeof createPresentationExistingEditingSkill> | undefined
  let existingBatchEditing:
    ReturnType<typeof createPresentationExistingBatchEditingSkill> | undefined
  let existingImageEditing:
    ReturnType<typeof createPresentationExistingImageEditingSkill> | undefined
  let existingPageEditing: ReturnType<typeof createPresentationExistingPageEditingSkill> | undefined
  let mutationStarted = false
  const readNativeLocks = (
    proposal: Parameters<typeof readPresentationNativeLocks>[1],
    signal: AbortSignal,
  ) =>
    readPresentationNativeLocks(
      {
        ...localBinding!,
        available: options.presentation?.available ?? (() => false),
        request:
          options.presentation?.request ??
          (async () => {
            throw new Error('presentation_lock_review_unavailable')
          }),
        hostSlideIds: async (readSignal) =>
          (await createBrowserPresentationImportAdapter().snapshot(readSignal)).slideIds,
      },
      proposal,
      signal,
    )
  const proposals = createStructuredProposalController(
    options.diagnostics,
    host === 'powerpoint'
      ? {
          review: readNativeLocks,
          beforeWrite: async (proposal, signal) => {
            // Local preference and review writes do not touch the host deck or its QA state.
            if (
              proposal.toolName === proposal.operation &&
              ((['save_presentation_preference', 'delete_presentation_preference'].includes(
                proposal.operation,
              ) &&
                proposal.impact.host === 'local_preference') ||
                ([
                  'add_presentation_review_comment',
                  'resolve_presentation_review_comment',
                ].includes(proposal.operation) &&
                  proposal.impact.host === 'local_review'))
            )
              return
            // Only these internally constructed operations resolve a stable host page before
            // proposing. Generic script/index-based impact labels cannot prove their write scope.
            const hostSlideIds = presentationMutationScope(proposal)
            existingEditing?.beginMutation()
            existingBatchEditing?.beginMutation()
            qaSkill?.beginMutation(hostSlideIds)
            mutationStarted = Boolean(qaSkill)
            await localBinding?.invalidateQa?.(hostSlideIds)
            notifyQa()
            const fresh = await readNativeLocks(proposal, signal)
            if (
              fresh?.token !==
              (proposal.lockReview?.state === 'ready' ? proposal.lockReview.token : undefined)
            )
              throw new Error('presentation_lock_review_stale')
          },
          afterWrite: () => {
            existingEditing?.endMutation()
            existingBatchEditing?.endMutation()
            if (mutationStarted) {
              mutationStarted = false
              qaSkill!.endMutation()
            }
            notifyQa()
          },
        }
      : undefined,
  )
  const shared = createSharedBrowserSkill({
    vfs,
    skills,
    enableConversions: options.enableConversions,
  })
  const powerPointAdapter = host === 'powerpoint' ? new BrowserPowerPointAdapter() : undefined
  const inspectQaPage =
    powerPointAdapter && options.presentation
      ? createPresentationPageScreenshotFallback({
          adapter: powerPointAdapter,
          request: options.presentation.request,
          documentId: options.presentation.documentId,
        })
      : undefined
  const hostSkill = {
    word: () => createWordSkill({ adapter: new BrowserWordAdapter(), vfs, proposals }),
    excel: () => createExcelSkill({ adapter: new BrowserExcelAdapter(), proposals }),
    powerpoint: () =>
      createPowerPointSkill({
        adapter: powerPointAdapter!,
        proposals,
        vfs,
        nativeMasterEditingSupported: supportsNativePowerPointMasterEditing(),
        platform: options.platform ?? currentOfficePlatform(),
        screenshotFallback:
          powerPointAdapter && inspectQaPage
            ? async (index, signal, expectedSlideId) => {
                const before = await powerPointAdapter.exportSlidePackage(index, signal)
                if (expectedSlideId && before.slideId !== expectedSlideId)
                  throw new Error('office_concurrent_change')
                const inspected = await inspectQaPage(before.slideId, signal)
                const after = await powerPointAdapter.exportSlidePackage(index, signal)
                if (
                  after.slideId !== before.slideId ||
                  (await presentationPackageDigest(after.base64, signal)) !==
                    (await presentationPackageDigest(before.base64, signal))
                )
                  throw new Error('office_concurrent_change')
                return {
                  slideId: before.slideId,
                  mime: 'image/png' as const,
                  base64: inspected.screenshot.base64,
                  ...(inspected.screenshot.renderer
                    ? { renderer: inspected.screenshot.renderer }
                    : {}),
                }
              }
            : undefined,
        chartSavepoint:
          localBinding?.readExistingChartChange &&
          localBinding.writeExistingChartChange &&
          options.presentation?.request
            ? {
                documentId: localBinding.documentId,
                request: options.presentation.request,
                readExistingChartChange: localBinding.readExistingChartChange,
                writeExistingChartChange: localBinding.writeExistingChartChange,
              }
            : undefined,
      }),
  }[host]()
  const extensions =
    options.enableImportMedia === false
      ? []
      : host === 'excel' && supportsExcelImportMedia()
        ? [
            createExcelImportMediaSkill({
              adapter: new BrowserExcelImportMediaAdapter(),
              proposals,
              vfs,
              enableImage: supportsBrowserMediaValidation(),
            }),
          ]
        : host === 'powerpoint' &&
            powerPointAdapter &&
            supportsPowerPointImportMedia() &&
            supportsBrowserMediaValidation()
          ? [
              createPowerPointImportMediaSkill({
                adapter: new BrowserPowerPointImportMediaAdapter(powerPointAdapter),
                proposals,
                vfs,
              }),
            ]
          : []
  const baselineSkill = powerPointAdapter
    ? createPresentationBaselineSkill({
        adapter: new BrowserPresentationBaselineAdapter(),
        documentId: localBinding!.documentId,
        inspectPage: (slideId, signal) =>
          powerPointAdapter.inspectPresentationPage(slideId, signal),
        exportPagePackage: (slideId, signal) =>
          powerPointAdapter.exportPresentationPagePackage(slideId, signal),
        readMasters: (signal) => powerPointAdapter.inspectSlideMasters(signal),
      })
    : undefined
  const imageBackup = options.presentation
    ? createPresentationImageBackup({
        available: () =>
          Boolean(
            options.presentation?.attachmentsAvailable?.() &&
            options.presentation?.assetsAvailable?.(),
          ),
        request: options.presentation.attachmentsRequest ?? options.presentation.request,
      })
    : undefined
  if (
    baselineSkill &&
    powerPointAdapter &&
    localBinding?.readExistingChange &&
    localBinding.writeExistingChange &&
    localBinding.listChangeHistory
  )
    existingEditing = createPresentationExistingEditingSkill({
      baseline: baselineSkill,
      baselineAdapter: new BrowserPresentationBaselineAdapter(),
      adapter: powerPointAdapter,
      proposals,
      documentId: localBinding.documentId,
      request:
        options.presentation?.request ??
        (async () => {
          throw new Error('presentation_page_backup_unavailable')
        }),
      readExistingChange: localBinding.readExistingChange,
      listChangeHistory: localBinding.listChangeHistory,
      writeExistingChange: async (record, expected) => {
        try {
          await localBinding.writeExistingChange!(record, expected)
        } finally {
          void changes?.refresh()
        }
      },
    })
  if (
    baselineSkill &&
    powerPointAdapter &&
    imageBackup &&
    import.meta.env.MODE === 'test' &&
    options.imageAdapterOverrideForTests === true &&
    localBinding?.readExistingImageChange &&
    localBinding.writeExistingImageChange
  )
    existingImageEditing = createPresentationExistingImageEditingSkill({
      baseline: baselineSkill,
      imageAdapter: new BrowserPresentationImageAdapter(),
      inspectPage: (slideId, signal) => powerPointAdapter.inspectPresentationPage(slideId, signal),
      imageBackup,
      vfs,
      proposals,
      documentId: localBinding.documentId,
      readExistingImageChange: localBinding.readExistingImageChange,
      writeExistingImageChange: async (record, expected) => {
        try {
          await localBinding.writeExistingImageChange!(record, expected)
        } finally {
          void changes?.refresh()
        }
      },
    })
  if (
    baselineSkill &&
    powerPointAdapter &&
    options.presentation &&
    localBinding?.readExistingPageChange &&
    localBinding.writeExistingPageChange
  )
    existingPageEditing = createPresentationExistingPageEditingSkill({
      baseline: baselineSkill,
      adapter: new BrowserPresentationPageReplacementAdapter(),
      exportAdapter: powerPointAdapter,
      inspectPage: (slideId, signal) => powerPointAdapter.inspectPresentationPage(slideId, signal),
      vfs,
      request: options.presentation.request,
      proposals,
      documentId: localBinding.documentId,
      readExistingPageChange: localBinding.readExistingPageChange,
      readExistingChange: localBinding.readExistingChange,
      readExistingBatch: localBinding.readExistingBatch,
      writeExistingPageChange: async (record, expected) => {
        try {
          await localBinding.writeExistingPageChange!(record, expected)
        } finally {
          void changes?.refresh()
        }
      },
      available: options.presentation.available,
    })
  if (
    baselineSkill &&
    powerPointAdapter &&
    localBinding?.readExistingBatch &&
    localBinding.writeExistingBatch
  )
    existingBatchEditing = createPresentationExistingBatchEditingSkill({
      baseline: baselineSkill,
      baselineAdapter: new BrowserPresentationBaselineAdapter(),
      adapter: powerPointAdapter,
      proposals,
      documentId: localBinding.documentId,
      request:
        options.presentation?.request ??
        (async () => {
          throw new Error('presentation_page_backup_unavailable')
        }),
      readExistingBatch: localBinding.readExistingBatch,
      writeExistingBatch: async (record, expected) => {
        try {
          await localBinding.writeExistingBatch!(record, expected)
        } finally {
          void changes?.refresh()
        }
      },
    })
  const base = composeOfficeSkills(hostSkill, shared, [
    ...extensions,
    ...(baselineSkill ? [baselineSkill] : []),
    ...(existingEditing ? [existingEditing] : []),
    ...(existingBatchEditing ? [existingBatchEditing] : []),
    ...(existingImageEditing ? [existingImageEditing] : []),
    ...(existingPageEditing ? [existingPageEditing] : []),
  ])
  const generation =
    host === 'powerpoint' && options.presentation
      ? createPresentationGenerationSkill({ ...options.presentation, vfs })
      : undefined
  const attachments =
    generation && options.presentation
      ? createPresentationAttachmentSkill({
          vfs,
          documentId: options.presentation.documentId,
          available: options.presentation.attachmentsAvailable ?? (() => false),
          request: options.presentation.attachmentsRequest ?? options.presentation.request,
          imagesAvailable: () =>
            Boolean(
              options.presentation?.attachmentsAvailable?.() &&
              options.presentation?.assetsAvailable?.(),
            ),
          remoteImagesAvailable: () => Boolean(options.presentation?.remoteImagesAvailable?.()),
          webpagesAvailable: () => Boolean(options.presentation?.webpagesAvailable?.()),
          rightsAvailable: () => Boolean(options.presentation?.rightsAvailable?.()),
        })
      : undefined
  const planning =
    generation && options.presentation
      ? createPresentationPlanningSkill({ ...options.presentation, vfs, proposals })
      : undefined
  const comments =
    generation && options.presentation
      ? createPresentationCommentsSkill({ ...options.presentation, proposals })
      : undefined
  const production =
    generation && options.presentation
      ? createPresentationProductionSkill({ ...options.presentation, vfs })
      : undefined
  const productionJobs =
    generation && options.presentation
      ? createPresentationJobsSkill(options.presentation)
      : undefined
  const evidenceDelivery =
    generation && options.presentation
      ? createPresentationEvidenceDeliverySkill({ ...options.presentation, vfs })
      : undefined
  let importSource: 'generation' | 'production' = 'generation'
  let importSelectionEpoch = 0
  const activeArtifact = (projectId?: string) =>
    importSource === 'production'
      ? production?.artifact(projectId)
      : generation?.artifact(projectId)
  // A cached source can outlive a mapping switch; hide it from current views and QA.
  const visibleArtifact = (projectId?: string) => {
    const artifact = activeArtifact(projectId)
    if (!artifact) return undefined
    const change = options.presentation?.readPageReplacement?.()
    if (
      change?.documentId === artifact.documentId &&
      change.projectId === artifact.projectId &&
      [change.parentRequestId, change.requestId].includes(artifact.requestId) &&
      ['commit_pending', 'undo_pending', 'restore_inserted'].includes(change.state)
    )
      return undefined
    try {
      options.presentation?.readReceipt?.(presentationImportKey(artifact))
    } catch (error) {
      if (error instanceof Error && error.message === 'presentation_import_superseded')
        return undefined
      throw error
    }
    return artifact
  }
  const pageBackup =
    production && powerPointAdapter && options.presentation?.readReceipt
      ? createPresentationPageBackupSkill({
          available: options.presentation.available,
          request: options.presentation.request,
          documentId: options.presentation.documentId,
          vfs,
          artifact: activeArtifact,
          readReceipt: options.presentation.readReceipt,
          adapter: powerPointAdapter,
        })
      : undefined
  const pageReplacement =
    pageBackup &&
    options.presentation?.readReceipt &&
    options.presentation.readPageReplacement &&
    options.presentation.writePageReplacement
      ? createPresentationPageReplacementSkill({
          available: options.presentation.available,
          request: options.presentation.request,
          documentId: options.presentation.documentId,
          artifact: activeArtifact,
          readReceipt: options.presentation.readReceipt,
          readPageReplacement: options.presentation.readPageReplacement,
          writePageReplacement: async (record, expected) => {
            try {
              await options.presentation!.writePageReplacement!(record, expected)
            } finally {
              notifyImport()
              notifyQa()
            }
          },
          loadBackup: pageBackup.loadBackup,
          adapter: new BrowserPresentationPageReplacementAdapter(),
          proposals,
        })
      : undefined
  const executeGeneration: AgentSkill['executeTool'] = async (call, signal) => {
    const selection = ++importSelectionEpoch
    try {
      const result = await generation!.executeTool(call, signal)
      if (!result.isError && !signal?.aborted && selection === importSelectionEpoch)
        importSource = 'generation'
      return result
    } finally {
      notifyImport()
      notifyQa()
    }
  }
  const researchSkill =
    generation && options.presentation
      ? createPresentationResearchSkill({
          ...options.presentation,
          vfs,
          onChanged: (projectId) => {
            void research?.selectProject(projectId)
          },
        })
      : undefined
  const research =
    researchSkill && options.presentation
      ? createPresentationResearchController({
          available: options.presentation.available,
          request: options.presentation.request,
          documentId: options.presentation.documentId,
          lastProject: options.presentation.lastProject,
          executeTool: researchSkill.executeTool,
          readDeleteAttempt: options.presentation.readResearchDeleteAttempt,
          writeDeleteAttempt: options.presentation.writeResearchDeleteAttempt,
        })
      : undefined
  const hostBundle =
    generation && options.presentation
      ? createPresentationHostBundleSkill({
          available: options.presentation.available,
          request: options.presentation.request,
          documentId: options.presentation.documentId,
          vfs,
          nativeAvailable: supportsPowerPointDocumentExport,
          exportDocument: exportPowerPointDocument,
          readQuality: (projectId, requestId) => {
            const artifact = production?.artifact(projectId)
            if (!artifact || artifact.requestId !== requestId) return null
            const record = options.presentation?.readQa?.(presentationImportKey(artifact))
            return record?.source === 'production' &&
              record.documentId === artifact.documentId &&
              record.projectId === projectId &&
              record.requestId === requestId
              ? record
              : null
          },
          readCheckpoints: () => options.presentation?.listChangeHistory?.() ?? [],
          readResearch: researchSkill?.readLatestCompleted,
        })
      : undefined
  const presentation =
    generation && options.presentation
      ? createPresentationProjectController({
          ...options.presentation,
          nativeDocumentExportAvailable: () => supportsPowerPointDocumentExport('pptx'),
          readResearchRecord: research
            ? async (projectId, ledgerId, signal) => {
                const binding = presentation?.snapshot().project?.plan?.value.research
                const documentId = await options.presentation!.documentId()
                const abort = () => research.clear()
                const check = async () => {
                  if (signal?.aborted) {
                    research.clear()
                    throw new Error('cancelled')
                  }
                  if ((await options.presentation!.documentId()) !== documentId) {
                    research.clear()
                    throw new Error('presentation_document_changed')
                  }
                  if (signal?.aborted) {
                    research.clear()
                    throw new Error('cancelled')
                  }
                }
                signal?.addEventListener('abort', abort, { once: true })
                try {
                  await check()
                  await research.selectProject(projectId)
                  await check()
                  if (research.snapshot().available === false)
                    throw new Error('presentation_upgrade_required')
                  await research.read(ledgerId)
                  await check()
                  const snapshot = research.snapshot()
                  const record = snapshot.record
                  if (snapshot.error || !record)
                    throw new Error('presentation_research_unavailable')
                  if (
                    !binding ||
                    record.state !== 'completed' ||
                    record.documentId !== documentId ||
                    record.projectId !== projectId ||
                    record.id !== ledgerId ||
                    record.sequence !== binding.sequence ||
                    record.draftDigest !== binding.draftDigest
                  ) {
                    research.clear()
                    throw new Error('presentation_research_binding_invalid')
                  }
                } finally {
                  signal?.removeEventListener('abort', abort)
                }
              }
            : undefined,

          ...(options.presentation.listReceipts
            ? {
                hostSlideIds: async (signal?: AbortSignal) =>
                  (await createBrowserPresentationImportAdapter().snapshot(signal)).slideIds,
              }
            : {}),
          executeTool: (call, signal) =>
            hostBundle?.tools.some((tool) => tool.name === call.name)
              ? hostBundle.executeTool(call, signal)
              : evidenceDelivery?.tools.some((tool) => tool.name === call.name)
                ? executeEvidenceDelivery(call, signal)
                : productionJobs?.tools.some((tool) => tool.name === call.name)
                  ? executeProductionJob(call, signal)
                  : production?.tools.some((tool) => tool.name === call.name)
                    ? executeProduction(call, signal)
                    : planning?.tools.some((tool) => tool.name === call.name)
                      ? planning.executeTool(call, signal)
                      : executeGeneration(call, signal),
        })
      : undefined
  let productionEpoch = 0
  const executeProduction: AgentSkill['executeTool'] = async (call, signal) => {
    const captured = productionEpoch
    const selection =
      call.name === 'prepare_presentation_production_import' ? ++importSelectionEpoch : undefined
    const result = await production!.executeTool(call, signal)
    if (
      captured === productionEpoch &&
      !signal?.aborted &&
      !result.isError &&
      [
        'start_presentation_production',
        'run_presentation_production',
        'rebuild_presentation_page',
      ].includes(call.name)
    )
      await presentation?.refresh()
    if (captured !== productionEpoch || signal?.aborted)
      return {
        output: 'cancelled',
        isError: true,
        mutated: false,
        summary: '已停止等待；PC可能已保存成果，可刷新查看',
      }
    if (!result.isError && selection !== undefined && selection === importSelectionEpoch) {
      importSource = 'production'
      notifyImport()
      notifyQa()
    }
    return result
  }
  const executeProductionJob: AgentSkill['executeTool'] = async (call, signal) => {
    const captured = productionEpoch
    const result = await productionJobs!.executeTool(call, signal)
    if (
      !result.isError &&
      !signal?.aborted &&
      captured === productionEpoch &&
      call.name !== 'read_presentation_production_job'
    )
      await presentation?.refresh()
    if (captured !== productionEpoch || signal?.aborted)
      return {
        output: 'presentation_aborted',
        isError: true,
        mutated: false,
        summary: '已停止等待；后台任务可能继续，可重连后查看',
      }
    return result
  }
  const executeEvidenceDelivery: AgentSkill['executeTool'] = async (call, signal) => {
    const captured = productionEpoch
    const result = await evidenceDelivery!.executeTool(call, signal)
    const selected = presentation?.snapshot().project?.production
    if (
      !result.isError &&
      !signal?.aborted &&
      captured === productionEpoch &&
      call.name !== 'read_presentation_delivery_report' &&
      selected?.projectId === call.input.project_id &&
      selected?.requestId === call.input.request_id
    )
      await presentation?.readDeliveryReport()
    if (captured !== productionEpoch || signal?.aborted)
      return {
        output: 'presentation_aborted',
        isError: true,
        mutated: false,
        summary: '已停止等待；保存结果可重新读取确认',
      }
    return result
  }
  let importRevision = 0
  const importListeners = new Set<() => void>()
  const notifyImport = () => {
    importRevision++
    for (const listener of importListeners) listener()
  }
  const importProgress: PresentationImportProgressController | undefined =
    generation && options.presentation?.readReceipt
      ? {
          read: () => {
            const artifact = visibleArtifact()
            return artifact
              ? summarizePresentationImport(
                  artifact,
                  options.presentation!.readReceipt!(presentationImportKey(artifact)),
                )
              : undefined
          },
          revision: () => importRevision,
          subscribe: (listener) => {
            importListeners.add(listener)
            return () => {
              importListeners.delete(listener)
            }
          },
        }
      : undefined
  let qaRevision = 0
  const qaListeners = new Set<() => void>()
  const notifyQa = () => {
    qaRevision++
    for (const listener of qaListeners) listener()
    void changes?.refresh()
  }
  const qaSkill =
    generation &&
    powerPointAdapter &&
    options.presentation?.readReceipt &&
    options.presentation.readQa &&
    options.presentation.writeQa
      ? createPresentationQaSkill({
          vfs,
          available: () =>
            options.presentation!.available() && supportsNativePowerPointMasterEditing(),
          artifact: visibleArtifact,
          documentId: options.presentation.documentId,
          readReceipt: options.presentation.readReceipt,
          inspectPage: inspectQaPage!,
          exportPage: (id, signal) => powerPointAdapter.exportPresentationPagePackage(id, signal),
          readQa: options.presentation.readQa,
          writeQa: async (key, record) => {
            try {
              await options.presentation!.writeQa!(key, record)
            } finally {
              notifyQa()
            }
          },
        })
      : undefined
  const pageEditing =
    generation && powerPointAdapter && options.presentation?.readReceipt
      ? createPresentationPageEditingSkill({
          available: () =>
            options.presentation!.available() && supportsNativePowerPointMasterEditing(),
          artifact: activeArtifact,
          documentId: options.presentation.documentId,
          readReceipt: options.presentation.readReceipt,
          adapter: powerPointAdapter,
          vfs,
          imageAdapter:
            import.meta.env.MODE === 'test' && options.imageAdapterOverrideForTests === true
              ? new BrowserPresentationImageAdapter()
              : undefined,
          imageBackup,
          readTextChange: options.presentation.readTextChange,
          writeTextChange: options.presentation.writeTextChange
            ? async (record, expected) => {
                try {
                  await options.presentation!.writeTextChange!(record, expected)
                } finally {
                  void changes?.refresh()
                }
              }
            : undefined,
          readGeometryChange: options.presentation.readGeometryChange,
          writeGeometryChange: options.presentation.writeGeometryChange
            ? async (record, expected) => {
                try {
                  await options.presentation!.writeGeometryChange!(record, expected)
                } finally {
                  void changes?.refresh()
                }
              }
            : undefined,
          readImageReplacement: options.presentation.readImageReplacement,
          writeImageReplacement: options.presentation.writeImageReplacement
            ? async (key, record) => {
                try {
                  await options.presentation!.writeImageReplacement!(key, record)
                } finally {
                  void changes?.refresh()
                }
              }
            : undefined,
          proposals,
        })
      : undefined
  const dispatchChangeTool: AgentSkill['executeTool'] = (call, signal) => {
    const owner = existingEditing?.tools.some((tool) => tool.name === call.name)
      ? existingEditing
      : existingBatchEditing?.tools.some((tool) => tool.name === call.name)
        ? existingBatchEditing
        : existingImageEditing?.tools.some((tool) => tool.name === call.name)
          ? existingImageEditing
          : existingPageEditing?.tools.some((tool) => tool.name === call.name)
            ? existingPageEditing
            : pageEditing?.tools.some((tool) => tool.name === call.name)
              ? pageEditing
              : pageReplacement?.tools.some((tool) => tool.name === call.name)
                ? pageReplacement
                : undefined
    return owner
      ? owner.executeTool(call, signal)
      : Promise.resolve({
          output: 'presentation_unavailable',
          isError: true,
          mutated: false,
          summary: '当前操作不可用',
        })
  }
  const executeChangeTool: AgentSkill['executeTool'] = async (call, signal) => {
    try {
      return await dispatchChangeTool(call, signal)
    } finally {
      await changes?.refresh()
    }
  }
  if (
    existingEditing ||
    existingBatchEditing ||
    existingImageEditing ||
    existingPageEditing ||
    (generation &&
      options.presentation &&
      (options.presentation.listChangeHistory ||
        options.presentation.readTextChange ||
        options.presentation.readGeometryChange ||
        options.presentation.readPageReplacement ||
        options.presentation.listImageReplacements))
  )
    changes = createPresentationChangesController({
      available: options.presentation?.available ?? (() => false),
      existingAvailable: () =>
        Boolean(
          existingEditing || existingBatchEditing || existingImageEditing || existingPageEditing,
        ),
      artifact: activeArtifact,
      documentId: localBinding!.documentId,
      listChangeHistory: localBinding!.listChangeHistory,
      listExistingPageBackups: options.presentation?.request
        ? async (documentId) => {
            const response = await options.presentation!.request!({
              operation: 'existing_page_backup_list',
              documentId,
            })
            if (!response.ok) throw new Error('backup_inventory_unavailable')
            const value = (await response.json()) as { documentId?: unknown; backups?: unknown }
            if (value.documentId !== documentId || !Array.isArray(value.backups))
              throw new Error('backup_inventory_invalid')
            return value.backups as {
              backupId: string
              status: string
              hostSlideId: string
              slideIds: string[]
              sha256: string
              sizeBytes: number
            }[]
          }
        : undefined,
      readTextChange: options.presentation?.readTextChange,
      readGeometryChange: options.presentation?.readGeometryChange,
      readPageReplacement: options.presentation?.readPageReplacement,
      listImageReplacements: options.presentation?.listImageReplacements,
      executeTool: dispatchChangeTool,
    })
  const historySkill =
    generation && options.presentation?.listChangeHistory
      ? createPresentationHistorySkill({
          available: options.presentation.available,
          artifact: activeArtifact,
          documentId: options.presentation.documentId,
          listChangeHistory: options.presentation.listChangeHistory,
        })
      : undefined
  const qa: PresentationQaController | undefined =
    qaSkill && generation
      ? {
          read: () => {
            const artifact = visibleArtifact()
            return artifact
              ? options.presentation!.readQa!(presentationImportKey(artifact))
              : undefined
          },
          revision: () => qaRevision,
          subscribe: (listener) => {
            qaListeners.add(listener)
            return () => {
              qaListeners.delete(listener)
            }
          },
        }
      : undefined
  const delivery =
    generation && options.presentation?.readReceipt && options.presentation.writeReceipt
      ? createPresentationDeliverySkill({
          adapter: createBrowserPresentationImportAdapter(),
          proposals,
          artifact: generation.artifact,
          available: options.presentation.available,
          documentId: options.presentation.documentId,
          readReceipt: options.presentation.readReceipt,
          writeReceipt: async (key, record) => {
            try {
              await options.presentation!.writeReceipt!(key, record)
            } finally {
              notifyImport()
            }
          },
        })
      : undefined
  const productionDelivery =
    production && options.presentation?.readReceipt && options.presentation.writeReceipt
      ? createPresentationProductionDeliverySkill({
          adapter: createBrowserPresentationImportAdapter(),
          proposals,
          artifact: production.artifact,
          available: options.presentation.available,
          documentId: options.presentation.documentId,
          readReceipt: options.presentation.readReceipt,
          writeReceipt: async (key, record) => {
            try {
              await options.presentation!.writeReceipt!(key, record)
            } finally {
              notifyImport()
            }
          },
        })
      : undefined
  const executeDelivery: AgentSkill['executeTool'] = async (call, signal) => {
    const selection = ++importSelectionEpoch
    const captured = productionEpoch
    const isProduction = [
      'import_presentation_production',
      'read_presentation_production_import_status',
      'reconcile_presentation_production_import',
    ].includes(call.name)
    const result = await (isProduction ? productionDelivery! : delivery!).executeTool(call, signal)
    if (captured !== productionEpoch || signal?.aborted)
      return {
        output: 'cancelled',
        isError: true,
        mutated: false,
        summary: '已停止等待；请重新读取导入记录',
      }
    if (!result.isError && selection === importSelectionEpoch) {
      importSource = isProduction ? 'production' : 'generation'
      notifyImport()
      notifyQa()
    }
    return result
  }
  const skill: AgentSkill = generation
    ? {
        ...base,
        get tools() {
          return [
            ...base.tools,
            ...generation.tools,
            ...(planning?.tools ?? []),
            ...(comments?.tools ?? []),
            ...(attachments?.tools ?? []),
            ...(delivery?.tools ?? []),
            ...(productionDelivery?.tools ?? []),
            ...(qaSkill?.tools ?? []),
            ...(pageEditing?.tools ?? []),
            ...(historySkill?.tools ?? []),
            ...(production?.tools ?? []),
            ...(productionJobs?.tools ?? []),
            ...(evidenceDelivery?.tools ?? []),
            ...(hostBundle?.tools ?? []),
            ...(researchSkill?.tools ?? []),
            ...(pageBackup?.tools ?? []),
            ...(pageReplacement?.tools ?? []),
          ]
        },
        get systemPrompt() {
          return `${base.systemPrompt}\n\n${generation.tools.length ? generation.systemPrompt : ''}\n${delivery?.tools.length ? delivery.systemPrompt : ''}\n${productionDelivery?.tools.length ? productionDelivery.systemPrompt : ''}\n${planning?.tools.length ? planning.systemPrompt : ''}\n${comments?.tools.length ? comments.systemPrompt : ''}\n${attachments?.tools.length ? attachments.systemPrompt : ''}\n${qaSkill?.tools.length ? qaSkill.systemPrompt : ''}\n${pageEditing?.tools.length ? pageEditing.systemPrompt : ''}\n${historySkill?.tools.length ? historySkill.systemPrompt : ''}\n${production?.tools.length ? production.systemPrompt : ''}\n${productionJobs?.tools.length ? productionJobs.systemPrompt : ''}\n${evidenceDelivery?.tools.length ? evidenceDelivery.systemPrompt : ''}\n${hostBundle?.tools.length ? hostBundle.systemPrompt : ''}\n${researchSkill?.tools.length ? researchSkill.systemPrompt : ''}\n${pageBackup?.tools.length ? pageBackup.systemPrompt : ''}\n${pageReplacement?.tools.length ? pageReplacement.systemPrompt : ''}\nQA and stable page editing use the currently selected artifact: a successfully prepared production task or explicitly compiled/restored whole deck. Select the intended source before acting; do not substitute another task with the same IDs.`
        },
        buildContext: () =>
          [base.buildContext?.(), generation.buildContext?.()].filter(Boolean).join('\n\n'),
        executeTool: (call, signal) =>
          researchSkill?.tools.some((tool) => tool.name === call.name)
            ? researchSkill.executeTool(call, signal)
            : hostBundle?.tools.some((tool) => tool.name === call.name)
              ? hostBundle.executeTool(call, signal)
              : call.name === 'list_presentation_changes' && historySkill
                ? historySkill.executeTool(call, signal)
                : evidenceDelivery?.tools.some((tool) => tool.name === call.name)
                  ? executeEvidenceDelivery(call, signal)
                  : productionJobs?.tools.some((tool) => tool.name === call.name)
                    ? executeProductionJob(call, signal)
                    : pageReplacement?.tools.some((tool) => tool.name === call.name)
                      ? executeChangeTool(call, signal)
                      : ['save_presentation_page_backup', 'read_presentation_page_backup'].includes(
                            call.name,
                          ) && pageBackup
                        ? pageBackup.executeTool(call, signal)
                        : [
                              'start_presentation_production',
                              'rebuild_presentation_page',
                              'run_presentation_production',
                              'read_presentation_production',
                              'read_presentation_page_artifact',
                              'check_presentation_page_content',
                              'read_presentation_claim_evidence',
                              'record_presentation_claim_review',
                              'read_presentation_claim_review',
                              'read_presentation_page_reviews',
                              'prepare_presentation_production_import',
                            ].includes(call.name) && production
                          ? executeProduction(call, signal)
                          : pageEditing?.tools.some((tool) => tool.name === call.name)
                            ? executeChangeTool(call, signal)
                            : [
                                  'capture_presentation_page_qa',
                                  'read_presentation_qa',
                                  'record_presentation_page_review',
                                ].includes(call.name) && qaSkill
                              ? qaSkill.executeTool(call, signal)
                              : [
                                    'list_presentation_attachments',
                                    'read_presentation_attachment',
                                  ].includes(call.name) && attachments
                                ? attachments.executeTool(call, signal)
                                : comments?.tools.some((tool) => tool.name === call.name)
                                  ? comments.executeTool(call, signal)
                                  : planning?.tools.some((tool) => tool.name === call.name)
                                    ? planning.executeTool(call, signal)
                                    : [
                                          'import_presentation_production',
                                          'read_presentation_production_import_status',
                                          'reconcile_presentation_production_import',
                                        ].includes(call.name) && productionDelivery
                                      ? executeDelivery(call, signal)
                                      : [
                                            'import_generated_presentation',
                                            'read_presentation_import_status',
                                          ].includes(call.name) && delivery
                                        ? executeDelivery(call, signal)
                                        : [
                                              'compile_deck_with_pptxgenjs',
                                              'restore_presentation_project',
                                              'resume_presentation_project',
                                            ].includes(call.name)
                                          ? executeGeneration(call, signal)
                                          : base.executeTool(call, signal),
      }
    : base
  return {
    ...lifecycle(
      skill,
      proposals,
      vfs,
      skills,
      options.packageRuntime,
      options.enableSkillPackages !== false,
      () => {
        importSource = 'generation'
        importSelectionEpoch++
        productionEpoch++
        production?.clear()
        productionJobs?.clear()
        evidenceDelivery?.clear()
        pageBackup?.clear()
        pageReplacement?.clear()
        researchSkill?.clear()
        research?.clear()
        hostBundle?.clear()
        historySkill?.clear()
        baselineSkill?.clear()
        existingEditing?.clear()
        existingBatchEditing?.clear()
        existingImageEditing?.clear()
        existingPageEditing?.clear()
        pageEditing?.clear()
        changes?.clear()
        qaSkill?.clear()
        attachments?.clear()
        generation?.clear()
        planning?.clear()
        comments?.clear()
        presentation?.clear()
        notifyImport()
        notifyQa()
      },
      attachments && options.presentation
        ? {
            available: options.presentation.attachmentsAvailable ?? (() => false),
            upload: attachments.upload,
            list: attachments.list,
            acquisitionHistory: attachments.acquisitionHistory,
            remove: attachments.remove,
            importUrl: attachments.importUrl,
            importUrls: attachments.importUrls,
            importWebpage: attachments.importWebpage,
            attestLicense: attachments.attestLicense,
            revokeLicense: attachments.revokeLicense,
            extractFirstFrame: attachments.extractFirstFrame,
            imagesAvailable: () =>
              Boolean(
                options.presentation?.attachmentsAvailable?.() &&
                options.presentation?.assetsAvailable?.(),
              ),
            remoteImagesAvailable: () => Boolean(options.presentation?.remoteImagesAvailable?.()),
            webpagesAvailable: () => Boolean(options.presentation?.webpagesAvailable?.()),
            rightsAvailable: () => Boolean(options.presentation?.rightsAvailable?.()),
            animationFrameAvailable: () =>
              Boolean(options.presentation?.animationFrameAvailable?.()),
          }
        : undefined,
    ),
    ...(research ? { research } : {}),
    ...(presentation ? { presentation } : {}),
    ...(importProgress ? { importProgress } : {}),
    ...(qa ? { qa } : {}),
    ...(changes ? { changes } : {}),
  }
}

function lifecycle(
  skill: AgentSkill,
  proposals: ProposalController | StructuredProposalController,
  vfs: InMemoryVfs,
  skills: SkillRegistry,
  suppliedPackageRuntime?: Pick<SkillPackageWorkerRuntime, 'parse' | 'cancelAll'>,
  skillPackagesEnabled = true,
  onClear?: () => void,
  attachments?: {
    available(): boolean
    imagesAvailable(): boolean
    remoteImagesAvailable(): boolean
    webpagesAvailable(): boolean
    rightsAvailable(): boolean
    animationFrameAvailable(): boolean
    upload(name: string, content: Promise<ArrayBuffer>): Promise<void>
    acquisitionHistory?(): Promise<PresentationAcquisitionHistory | undefined>
    list(): Promise<PresentationAttachmentMetadata[]>
    remove(attachmentId: string): Promise<void>
    importUrl(url: string): Promise<unknown>
    importUrls(urls: string[]): Promise<unknown>
    importWebpage(url: string): Promise<unknown>
    attestLicense(
      imageId: string,
      license: 'owned' | 'licensed' | 'public_domain',
      evidenceId: string,
    ): Promise<unknown>
    revokeLicense(imageId: string): Promise<unknown>
    extractFirstFrame(imageId: string): Promise<unknown>
  },
): OfficeHostRuntime {
  const packageRuntime = suppliedPackageRuntime ?? new SkillPackageWorkerRuntime()
  let epoch = 0
  let disposed = false
  const check = (captured: number) => {
    if (disposed || captured !== epoch) throw new Error('upload_cancelled')
  }
  const clearSession = () => {
    epoch += 1
    packageRuntime.cancelAll()
    proposals.logout()
    skills.clear()
    vfs.clear()
    onClear?.()
  }
  return {
    skill,
    proposals,
    vfs,
    skills,
    skillPackagesEnabled,
    durableAttachmentsAvailable: () => attachments?.available() ?? false,
    durableImagesAvailable: () => attachments?.imagesAvailable() ?? false,
    remoteImagesAvailable: () => attachments?.remoteImagesAvailable() ?? false,
    webpagesAvailable: () => attachments?.webpagesAvailable() ?? false,
    rightsAvailable: () => attachments?.rightsAvailable() ?? false,
    animationFrameAvailable: () => attachments?.animationFrameAvailable() ?? false,
    readPresentationAcquisitionHistory: () =>
      attachments?.acquisitionHistory?.() ?? Promise.resolve(undefined),
    listDurableAttachments: () => attachments?.list() ?? Promise.resolve([]),
    deleteDurableAttachment: (attachmentId) =>
      attachments?.remove(attachmentId) ?? Promise.reject(new Error('presentation_unavailable')),
    importPresentationImageUrl: async (url) => {
      if (!attachments?.remoteImagesAvailable()) throw new Error('presentation_assets_unavailable')
      await attachments.importUrls(Array.isArray(url) ? url : [url])
    },
    importPresentationWebpageUrl: async (url) => {
      if (!attachments?.webpagesAvailable()) throw new Error('presentation_webpages_unavailable')
      await attachments.importWebpage(url)
    },
    attestPresentationImageLicense: async (imageId, license, evidenceId) => {
      if (!attachments?.rightsAvailable()) throw new Error('presentation_assets_unavailable')
      await attachments.attestLicense(imageId, license, evidenceId)
    },
    revokePresentationImageLicense: async (imageId) => {
      if (!attachments?.rightsAvailable()) throw new Error('presentation_assets_unavailable')
      await attachments.revokeLicense(imageId)
    },
    extractPresentationImageFirstFrame: async (imageId) => {
      if (!attachments?.animationFrameAvailable())
        throw new Error('presentation_assets_unavailable')
      await attachments.extractFirstFrame(imageId)
    },
    async uploadFile(name, content) {
      if (disposed) throw new Error('upload_cancelled')
      if (
        attachments?.available() &&
        supportsPresentationAttachment(name, attachments.imagesAvailable())
      ) {
        return attachments.upload(name, content)
      }
      const captured = epoch
      if (!name || name.length > 128 || name.includes('/') || name.includes('\\'))
        throw new Error('vfs_path_denied')
      const buffer = await content
      check(captured)
      if (buffer.byteLength > MAX_VFS_FILE_BYTES) throw new Error('vfs_limit')
      const bytes = new Uint8Array(buffer)
      vfs.writeFile(`/home/user/${name}`, bytes)
    },
    async installSkill(source) {
      if (!skillPackagesEnabled) throw new Error('office_capability_disabled')
      const captured = epoch
      const value = await source
      check(captured)
      if (new TextEncoder().encode(value).byteLength > MAX_SKILL_BYTES)
        throw new Error('invalid_skill_package')
      skills.install(value)
    },
    async installSkillPackage(source, signal) {
      if (!skillPackagesEnabled) throw new Error('office_capability_disabled')
      const captured = epoch
      const value = await source
      check(captured)
      const pkg = await packageRuntime.parse(value, signal)
      check(captured)
      skills.installPackage(pkg)
    },
    uninstallSkill(name) {
      skills.uninstall(name)
    },
    clearSession,
    dispose() {
      clearSession()
      disposed = true
    },
  }
}
