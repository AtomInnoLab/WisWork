import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

const panel = readFileSync(resolve(__dirname, '../src/renderer/ai/AiPanel.tsx'), 'utf8')

describe('Markdown agent harness integration', () => {
  it('maps the interactive lifecycle to AgentHarness', () => {
    expect(panel).toContain("from '@wiswork/agent-harness'")
    expect(panel).toContain('useRef<AgentHarness<string> | null>')
    expect(panel).toContain('createAgentHarness<string>({')
    expect(panel).toContain('harness.stop()')
    expect(panel).toContain('harnessRef.current?.reset()')
    expect(panel).toContain('harnessRef.current?.restore(')
    expect(panel).not.toContain('new AgentLoop')
  })

  it('keeps snapshots, transcript persistence, and autosave host-owned', () => {
    expect(panel).toContain('captureSnapshot: () => depsRef.current.getSnapshot()')
    expect(panel).toContain('setSnapshots((prev) =>')
    expect(panel).toContain("persistMessage('assistant', final, runToolsRef.current)")
    expect(panel).toContain('depsRef.current.onRunDone(runMutatedRef.current)')
  })
})
