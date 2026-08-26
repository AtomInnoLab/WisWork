export type ChatLoadState = 'loading' | 'ready' | 'error'

export interface AgentProjectScope {
  projectId: string
  generation: number
}

export interface AgentRunScope extends AgentProjectScope {
  run: number
}

export class AgentPanelSession {
  private projectId: string
  private generation = 0
  private run = 0
  private currentLoop: object | null = null
  private activeRun: AgentRunScope | null = null
  private validRun: AgentRunScope | null = null
  private cancelledRun: AgentRunScope | null = null
  private runStarted = false

  constructor(projectId: string) {
    this.projectId = projectId
  }

  attachLoop(loop: object, projectId: string): AgentProjectScope {
    this.projectId = projectId
    this.generation += 1
    this.currentLoop = loop
    this.activeRun = null
    this.validRun = null
    this.cancelledRun = null
    this.runStarted = false
    return this.captureProject()
  }

  detachLoop(loop: object): void {
    if (this.currentLoop !== loop) return
    this.generation += 1
    this.currentLoop = null
    this.activeRun = null
    this.validRun = null
    this.cancelledRun = null
  }

  captureProject(): AgentProjectScope {
    return { projectId: this.projectId, generation: this.generation }
  }

  beginRun(loop: object): AgentRunScope {
    if (this.currentLoop !== loop) throw new Error('agent loop is no longer current')
    this.run += 1
    const scope = { ...this.captureProject(), run: this.run }
    this.activeRun = scope
    this.validRun = scope
    this.cancelledRun = null
    this.runStarted = true
    return scope
  }

  cancelRun(loop: object, scope: AgentRunScope): void {
    if (!this.acceptsRun(loop, scope)) return
    this.cancelledRun = scope
    this.activeRun = null
    this.validRun = null
  }

  finishRun(loop: object, scope: AgentRunScope): void {
    if (this.currentLoop !== loop) return
    if (this.sameRun(this.activeRun, scope)) this.activeRun = null
    if (this.sameRun(this.cancelledRun, scope)) this.cancelledRun = null
  }

  acceptsLoopProject(loop: object, scope: AgentProjectScope): boolean {
    return this.currentLoop === loop && this.acceptsProject(scope)
  }

  acceptsProject(scope: AgentProjectScope): boolean {
    return scope.projectId === this.projectId && scope.generation === this.generation
  }

  acceptsRun(loop: object, scope: AgentRunScope): boolean {
    return this.acceptsLoopProject(loop, scope) && this.sameRun(this.activeRun, scope)
  }

  acceptsRunResult(loop: object, scope: AgentRunScope): boolean {
    return this.acceptsLoopProject(loop, scope) && this.sameRun(this.validRun, scope)
  }

  acceptsCompletion(loop: object, scope: AgentRunScope): boolean {
    return (
      this.acceptsLoopProject(loop, scope) &&
      (this.sameRun(this.activeRun, scope) || this.sameRun(this.cancelledRun, scope))
    )
  }

  canRestoreChat(loop: object, scope: AgentProjectScope): boolean {
    return this.acceptsLoopProject(loop, scope) && !this.runStarted
  }

  canSend(loop: object, chatState: ChatLoadState): boolean {
    return this.currentLoop === loop && chatState !== 'loading'
  }

  timelineId(scope: AgentRunScope, toolId: string): string {
    return `${scope.generation}.${scope.run}:${toolId}`
  }

  private sameRun(left: AgentRunScope | null, right: AgentRunScope): boolean {
    return Boolean(
      left &&
      left.projectId === right.projectId &&
      left.generation === right.generation &&
      left.run === right.run,
    )
  }
}
