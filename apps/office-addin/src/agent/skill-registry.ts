import { composeSkills, type AgentSkill } from '@wiswork/agent-core'

export function composeOfficeSkills(
  host: AgentSkill,
  shared: AgentSkill,
  extensions: AgentSkill[] = [],
): AgentSkill {
  return composeSkills(
    'office',
    'Office tools are host-scoped. Read tools do not mutate. PowerPoint ordinary writes follow the PC-managed session policy; raw Office writes always require confirmation.',
    [shared, host, ...extensions],
  )
}
