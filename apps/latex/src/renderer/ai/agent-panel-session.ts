export interface AgentProjectScope {
  projectId: string
  generation: number
}

export interface AgentRunScope extends AgentProjectScope {
  run: number
}

export class AgentPanelSession {
  private projectId: string
  private generation = 1
  private run = 0
  private activeRun: AgentRunScope | null = null
  private validRun: AgentRunScope | null = null
  private cancelledRun: AgentRunScope | null = null

  constructor(projectId: string) {
    this.projectId = projectId
  }

  captureProject(): AgentProjectScope {
    return { projectId: this.projectId, generation: this.generation }
  }

  switchProject(projectId: string): AgentProjectScope {
    this.projectId = projectId
    this.generation += 1
    this.activeRun = null
    this.validRun = null
    this.cancelledRun = null
    return this.captureProject()
  }

  beginRun(): AgentRunScope {
    this.run += 1
    const scope = { ...this.captureProject(), run: this.run }
    this.activeRun = scope
    this.validRun = scope
    this.cancelledRun = null
    return scope
  }

  cancelRun(scope: AgentRunScope): void {
    if (!this.sameRun(this.activeRun, scope)) return
    this.cancelledRun = scope
    this.activeRun = null
    this.validRun = null
  }

  finishRun(scope: AgentRunScope): void {
    if (this.sameRun(this.activeRun, scope)) this.activeRun = null
    if (this.sameRun(this.cancelledRun, scope)) this.cancelledRun = null
  }

  acceptsProject(scope: AgentProjectScope): boolean {
    return scope.projectId === this.projectId && scope.generation === this.generation
  }

  acceptsRun(scope: AgentRunScope): boolean {
    return this.acceptsProject(scope) && this.sameRun(this.activeRun, scope)
  }

  acceptsRunResult(scope: AgentRunScope): boolean {
    return this.acceptsProject(scope) && this.sameRun(this.validRun, scope)
  }

  acceptsCompletion(scope: AgentRunScope): boolean {
    return (
      this.acceptsProject(scope) &&
      (this.sameRun(this.activeRun, scope) || this.sameRun(this.cancelledRun, scope))
    )
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
