import type { AgentImage } from '@wiswork/agent-core'

// Leaves room for the MCP envelope under the bridge's 1 MB string budget.
export async function boundedScreenshot(
  render: (ratio: number) => Promise<string | undefined>,
): Promise<AgentImage | null> {
  for (const ratio of [1, 0.75, 0.5]) {
    const png = await render(ratio)
    if (png && png.length <= 800_000) return { mime: 'image/png', base64: png }
  }
  return null
}
