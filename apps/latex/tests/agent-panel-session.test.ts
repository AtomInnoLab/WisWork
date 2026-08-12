import { describe, expect, it } from 'vitest'
import { AgentPanelSession } from '../src/renderer/ai/agent-panel-session.js'

function deferred<T>() {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((done) => {
    resolve = done
  })
  return { promise, resolve }
}

describe('AI panel project and run isolation', () => {
  it('invalidates a cancelled run while keeping its completion eligible to settle busy UI', () => {
    const session = new AgentPanelSession('project-a')
    const run = session.beginRun()
    expect(session.acceptsRun(run)).toBe(true)

    session.cancelRun(run)
    expect(session.acceptsRun(run)).toBe(false)
    expect(session.acceptsRunResult(run)).toBe(false)
    expect(session.acceptsCompletion(run)).toBe(true)
  })

  it('accepts a deferred proposal after normal completion until a newer run starts', async () => {
    const session = new AgentPanelSession('project-a')
    const run = session.beginRun()
    const proposal = deferred<string>()
    const committed: string[] = []
    const completion = proposal.promise.then((value) => {
      if (session.acceptsRunResult(run)) committed.push(value)
    })

    session.finishRun(run)
    proposal.resolve('proposal-1')
    await completion
    expect(committed).toEqual(['proposal-1'])

    session.beginRun()
    expect(session.acceptsRunResult(run)).toBe(false)
  })

  it('rejects deferred work from an old project and resets tool ids per run scope', async () => {
    const session = new AgentPanelSession('project-a')
    const oldProject = session.captureProject()
    const oldRun = session.beginRun()
    const late = deferred<string>()
    const committed: string[] = []
    const completion = late.promise.then((value) => {
      if (session.acceptsProject(oldProject)) committed.push(value)
    })

    session.switchProject('project-b')
    const newRun = session.beginRun()
    expect(session.timelineId(oldRun, 'tool-1')).not.toBe(session.timelineId(newRun, 'tool-1'))
    late.resolve('old chat or proposal')
    await completion
    expect(committed).toEqual([])
  })
})
