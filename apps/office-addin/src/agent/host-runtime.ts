import type { AgentSkill } from '@wiswork/agent-core'
import {
  createOfficeDocumentClient,
  createBrowserOfficeRuntime,
  type OfficeDocumentClient,
  type OfficeHost,
} from '../office-document.js'
import { BrowserExcelAdapter } from '../skills/excel/browser-excel-adapter.js'
import { createExcelSkill } from '../skills/excel/excel-skill.js'
import { BrowserPowerPointAdapter } from '../skills/powerpoint/browser-powerpoint-adapter.js'
import { createPowerPointSkill } from '../skills/powerpoint/powerpoint-skill.js'
import { createSharedBrowserSkill } from '../skills/shared/shared-skill.js'
import { MAX_SKILL_BYTES, SkillRegistry } from '../skills/shared/skill-registry.js'
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
  uploadFile(name: string, content: Promise<ArrayBuffer>): Promise<void>
  installSkill(source: Promise<string>): Promise<void>
  clearSession(): void
  dispose(): void
}

export function createOfficeHostRuntime(
  host: OfficeHost,
  options: { enableHostSkills?: boolean; document?: OfficeDocumentClient } = {},
): OfficeHostRuntime {
  if (host === 'unknown') throw new Error('office_host_unsupported')
  const vfs = new InMemoryVfs()
  const skills = new SkillRegistry(vfs)
  if (options.enableHostSkills === false) {
    const document = options.document ?? createOfficeDocumentClient(createBrowserOfficeRuntime())
    const proposals = createProposalController(document)
    return lifecycle(createOfficeSkill(document, proposals), proposals, vfs, skills)
  }
  const proposals = createStructuredProposalController()
  const shared = createSharedBrowserSkill({ vfs, skills })
  const hostSkill = {
    word: () => createWordSkill({ adapter: new BrowserWordAdapter(), vfs, proposals }),
    excel: () => createExcelSkill({ adapter: new BrowserExcelAdapter(), proposals }),
    powerpoint: () => createPowerPointSkill({ adapter: new BrowserPowerPointAdapter(), proposals }),
  }[host]()
  return lifecycle(composeOfficeSkills(hostSkill, shared), proposals, vfs, skills)
}

function lifecycle(
  skill: AgentSkill,
  proposals: ProposalController | StructuredProposalController,
  vfs: InMemoryVfs,
  skills: SkillRegistry,
): OfficeHostRuntime {
  let epoch = 0
  let disposed = false
  const check = (captured: number) => {
    if (disposed || captured !== epoch) throw new Error('upload_cancelled')
  }
  const clearSession = () => {
    epoch += 1
    proposals.logout()
    skills.clear()
    vfs.clear()
  }
  return {
    skill,
    proposals,
    vfs,
    skills,
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
      const captured = epoch
      const value = await source
      check(captured)
      if (new TextEncoder().encode(value).byteLength > MAX_SKILL_BYTES)
        throw new Error('invalid_skill_package')
      skills.install(value)
    },
    clearSession,
    dispose() {
      clearSession()
      disposed = true
    },
  }
}
