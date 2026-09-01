import type { AgentSkill } from '@wiswork/agent-core'
import type {
  PresentationTelemetryEvent,
  PresentationVerificationFlags,
} from '@wiswork/presentation-verification'
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
import { createPowerPointSkill } from '../skills/powerpoint/powerpoint-skill.js'
import { createBrowserPowerPointVerificationAuthority } from '../skills/powerpoint/powerpoint-verification.js'
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
    enableHostSkills?: boolean
    document?: OfficeDocumentClient
    packageRuntime?: Pick<SkillPackageWorkerRuntime, 'parse' | 'cancelAll'>
    enableConversions?: boolean
    enableSkillPackages?: boolean
    enableImportMedia?: boolean
    platform?: string
    diagnostics?: Pick<OfficeDiagnostics, 'setTool' | 'record'>
    presentationVerification?: PresentationVerificationFlags
    presentationTelemetry?: (event: PresentationTelemetryEvent) => void
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
  const proposals = createStructuredProposalController(options.diagnostics)
  const shared = createSharedBrowserSkill({
    vfs,
    skills,
    enableConversions: options.enableConversions,
  })
  const powerPointAdapter = host === 'powerpoint' ? new BrowserPowerPointAdapter() : undefined
  const presentationFlags = options.presentationVerification ?? {
    planning: true,
    verifiedCompletion: true,
    visualReview: true,
    autoCorrection: false,
  }
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
        verificationAuthority: presentationFlags.verifiedCompletion
          ? createBrowserPowerPointVerificationAuthority(powerPointAdapter!)
          : undefined,
        presentationFlags,
        presentationTelemetry: options.presentationTelemetry,
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
  return lifecycle(
    composeOfficeSkills(hostSkill, shared, extensions),
    proposals,
    vfs,
    skills,
    options.packageRuntime,
    options.enableSkillPackages !== false,
  )
}

function lifecycle(
  skill: AgentSkill,
  proposals: ProposalController | StructuredProposalController,
  vfs: InMemoryVfs,
  skills: SkillRegistry,
  suppliedPackageRuntime?: Pick<SkillPackageWorkerRuntime, 'parse' | 'cancelAll'>,
  skillPackagesEnabled = true,
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
  }
  return {
    skill,
    proposals,
    vfs,
    skills,
    skillPackagesEnabled,
    async uploadFile(name, content) {
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
