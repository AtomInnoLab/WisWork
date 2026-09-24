import type { PresentationPageReplacement } from '../skills/powerpoint/presentation-page-replacement-record.js'
import { createPresentationPageReplacementSkill } from '../skills/powerpoint/presentation-page-replacement.js'
import { BrowserPresentationPageReplacementAdapter } from '../skills/powerpoint/browser-presentation-page-replacement-adapter.js'
import { createPresentationPageBackupSkill } from '../skills/powerpoint/presentation-page-backup.js'
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
} from '../skills/powerpoint/presentation-attachments.js'
import { createPresentationPlanningSkill } from '../skills/powerpoint/presentation-planning.js'
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
  readonly presentation?: PresentationProjectController
  readonly importProgress?: PresentationImportProgressController
  readonly qa?: PresentationQaController
  durableAttachmentsAvailable?(): boolean
  durableImagesAvailable?(): boolean
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
      readPageReplacement?(): PresentationPageReplacement | undefined
      writePageReplacement?(
        record: PresentationPageReplacement,
        expected: PresentationPageReplacement | undefined,
      ): Promise<void>
      readGeometryChange?(): PresentationGeometryChange | undefined
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
      writeReceipt?(key: string, record: PresentationImportRecord | undefined): Promise<void>
    }
    enableHostSkills?: boolean
    document?: OfficeDocumentClient
    packageRuntime?: Pick<SkillPackageWorkerRuntime, 'parse' | 'cancelAll'>
    enableConversions?: boolean
    enableSkillPackages?: boolean
    enableImportMedia?: boolean
    platform?: string
    diagnostics?: Pick<OfficeDiagnostics, 'setTool' | 'record'>
  } = {},
): OfficeHostRuntime {
  if (host === 'unknown') throw new Error('office_host_unsupported')
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
  let mutationStarted = false
  const proposals = createStructuredProposalController(
    options.diagnostics,
    host === 'powerpoint' && options.presentation?.invalidateQa
      ? {
          beforeWrite: async (proposal) => {
            // Only these internally constructed operations resolve a stable host page before
            // proposing. Generic script/index-based impact labels cannot prove their write scope.
            const target = proposal.impact.targets[0]
            const hostSlideIds =
              [
                'stage_presentation_page_replacement',
                'resume_presentation_page_replacement',
                'discard_presentation_page_replacement',
                'commit_presentation_page_replacement',
                'undo_presentation_page_replacement',
                'edit_presentation_page_text',
                'edit_presentation_page_geometry',
                'undo_presentation_geometry_change',
                'resume_presentation_geometry_change',
                'replace_presentation_page_image',
                'resume_presentation_image_replacement',
              ].includes(proposal.operation) &&
              proposal.operation === proposal.toolName &&
              proposal.impact.host === 'powerpoint' &&
              proposal.impact.count === 1 &&
              proposal.impact.targets.length === 1 &&
              typeof target === 'string' &&
              target.length > 0 &&
              target.length <= 256 &&
              !Array.from(target).some(
                (char) =>
                  char.charCodeAt(0) < 32 ||
                  (char.charCodeAt(0) >= 127 && char.charCodeAt(0) <= 159),
              )
                ? [target]
                : undefined
            qaSkill?.beginMutation(hostSlideIds)
            mutationStarted = Boolean(qaSkill)
            await options.presentation!.invalidateQa!(hostSlideIds)
            notifyQa()
          },
          afterWrite: () => {
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
  const base = composeOfficeSkills(hostSkill, shared, extensions)
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
        })
      : undefined
  const planning =
    generation && options.presentation
      ? createPresentationPlanningSkill({ ...options.presentation, vfs })
      : undefined
  const production =
    generation && options.presentation
      ? createPresentationProductionSkill({ ...options.presentation, vfs })
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
  const presentation =
    generation && options.presentation
      ? createPresentationProjectController({
          ...options.presentation,
          executeTool: (call, signal) =>
            call.name === 'run_presentation_production' && production
              ? production.executeTool(call, signal)
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
          inspectPage: (id, signal) => powerPointAdapter.inspectPresentationPage(id, signal),
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
          imageAdapter: new BrowserPresentationImageAdapter(),
          readGeometryChange: options.presentation.readGeometryChange,
          writeGeometryChange: options.presentation.writeGeometryChange,
          readImageReplacement: options.presentation.readImageReplacement,
          writeImageReplacement: options.presentation.writeImageReplacement,
          proposals,
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
            ...(attachments?.tools ?? []),
            ...(delivery?.tools ?? []),
            ...(productionDelivery?.tools ?? []),
            ...(qaSkill?.tools ?? []),
            ...(pageEditing?.tools ?? []),
            ...(production?.tools ?? []),
            ...(pageBackup?.tools ?? []),
            ...(pageReplacement?.tools ?? []),
          ]
        },
        get systemPrompt() {
          return `${base.systemPrompt}\n\n${generation.tools.length ? generation.systemPrompt : ''}\n${delivery?.tools.length ? delivery.systemPrompt : ''}\n${productionDelivery?.tools.length ? productionDelivery.systemPrompt : ''}\n${planning?.tools.length ? planning.systemPrompt : ''}\n${attachments?.tools.length ? attachments.systemPrompt : ''}\n${qaSkill?.tools.length ? qaSkill.systemPrompt : ''}\n${pageEditing?.tools.length ? pageEditing.systemPrompt : ''}\n${production?.tools.length ? production.systemPrompt : ''}\n${pageBackup?.tools.length ? pageBackup.systemPrompt : ''}\n${pageReplacement?.tools.length ? pageReplacement.systemPrompt : ''}\nQA and stable page editing use the currently selected artifact: a successfully prepared production task or explicitly compiled/restored whole deck. Select the intended source before acting; do not substitute another task with the same IDs.`
        },
        buildContext: () =>
          [base.buildContext?.(), generation.buildContext?.()].filter(Boolean).join('\n\n'),
        executeTool: (call, signal) =>
          [
            'stage_presentation_page_replacement',
            'inspect_presentation_page_replacement',
            'resume_presentation_page_replacement',
            'discard_presentation_page_replacement',
            'commit_presentation_page_replacement',
            'undo_presentation_page_replacement',
          ].includes(call.name) && pageReplacement
            ? pageReplacement.executeTool(call, signal)
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
                    'prepare_presentation_production_import',
                  ].includes(call.name) && production
                ? executeProduction(call, signal)
                : [
                      'read_presentation_page',
                      'edit_presentation_page_text',
                      'read_presentation_page_geometry',
                      'read_presentation_geometry_change',
                      'inspect_presentation_geometry_change',
                      'resume_presentation_geometry_change',
                      'undo_presentation_geometry_change',
                      'edit_presentation_page_geometry',
                      'replace_presentation_page_image',
                      'read_presentation_image_replacement',
                      'inspect_presentation_image_replacement',
                      'resume_presentation_image_replacement',
                    ].includes(call.name) && pageEditing
                  ? pageEditing.executeTool(call, signal)
                  : [
                        'capture_presentation_page_qa',
                        'read_presentation_qa',
                        'record_presentation_page_review',
                      ].includes(call.name) && qaSkill
                    ? qaSkill.executeTool(call, signal)
                    : ['list_presentation_attachments', 'read_presentation_attachment'].includes(
                          call.name,
                        ) && attachments
                      ? attachments.executeTool(call, signal)
                      : ['save_presentation_plan', 'read_presentation_plan'].includes(call.name) &&
                          planning
                        ? planning.executeTool(call, signal)
                        : [
                              'import_presentation_production',
                              'read_presentation_production_import_status',
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
        pageBackup?.clear()
        pageReplacement?.clear()
        pageEditing?.clear()
        qaSkill?.clear()
        attachments?.clear()
        generation?.clear()
        planning?.clear()
        presentation?.clear()
        notifyImport()
        notifyQa()
      },
      attachments && options.presentation
        ? {
            available: options.presentation.attachmentsAvailable ?? (() => false),
            upload: attachments.upload,
            imagesAvailable: () =>
              Boolean(
                options.presentation?.attachmentsAvailable?.() &&
                options.presentation?.assetsAvailable?.(),
              ),
          }
        : undefined,
    ),
    ...(presentation ? { presentation } : {}),
    ...(importProgress ? { importProgress } : {}),
    ...(qa ? { qa } : {}),
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
    upload(name: string, content: Promise<ArrayBuffer>): Promise<void>
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
