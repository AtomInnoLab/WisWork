/** Update a tool chip while retaining the host's display-specific fields. */
export function upsertToolActivity<T extends { callId?: string; running?: boolean }>(
  current: readonly T[] | undefined,
  activity: T,
): T[] {
  const tools = [...(current ?? [])]
  const index = activity.callId ? tools.findIndex((tool) => tool.callId === activity.callId) : -1
  if (index < 0) tools.push(activity)
  else if (!activity.running || tools[index]?.running) tools[index] = activity
  return tools
}
