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
import { SkillRegistry } from '../skills/shared/skill-registry.js'
import { InMemoryVfs } from '../skills/shared/vfs.js'
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
  const shared = createSharedBrowserSkill({ vfs })
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
  return {
    skill,
    proposals,
    vfs,
    skills,
    dispose() {
      proposals.logout()
      skills.clear()
      vfs.clear()
    },
  }
}
