import { composeSkills, type AgentSkill } from '@wiswork/agent-core'

export function composeOfficeSkills(
  host: AgentSkill,
  shared: AgentSkill,
  extensions: AgentSkill[] = [],
): AgentSkill {
  return composeSkills(
    'office',
    'Office tools are host-scoped. Read tools do not mutate; document writes require confirmation.',
    [shared, host, ...extensions],
  )
}
