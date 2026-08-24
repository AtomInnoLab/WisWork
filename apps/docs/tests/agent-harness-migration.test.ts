import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

const panel = readFileSync(resolve(__dirname, '../src/renderer/ai/AiPanel.tsx'), 'utf8')

describe('Docs agent harness integration', () => {
  it('maps the interactive lifecycle to AgentHarness', () => {
    expect(panel).toContain("from '@wiswork/agent-harness'")
    expect(panel).toContain('useRef<AgentHarness<PmNode> | null>')
    expect(panel).toContain('createAgentHarness<PmNode>({')
    expect(panel).toContain('harness.stop()')
    expect(panel).toContain('harnessRef.current?.reset()')
    expect(panel).toContain('harnessRef.current?.restore(')
    expect(panel).not.toContain('new AgentLoop')
  })

  it('keeps attachments, snapshots, and transcript persistence host-owned', () => {
    expect(panel).toContain('captureSnapshot: () => editorRef.current.getJSON() as PmNode')
    expect(panel).toContain('snapshot: runSnapshotRef.current ?? undefined')
    expect(panel).toContain('collectImageAttachments(sentAtts)')
    expect(panel).toContain("persistMessage('user', instruction, undefined, sentAtts)")
    expect(panel).toContain("persistMessage('assistant', finalText, runToolsRef.current)")
  })
})
