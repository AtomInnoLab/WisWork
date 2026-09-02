import type {
  EnhancedModeComponentManager,
  EnhancedModeComponentStatus as ComponentStatus,
} from '@wiswork/codex-bridge'
import {
  ENHANCED_MODE_CHANNELS,
  type EnhancedModeStatus,
  type ProductAgentMode,
} from '../shared/enhanced-mode-api'

interface ComponentOwner {
  isDestroyed(): boolean
  once(event: 'destroyed', listener: () => void): this
  removeListener(event: 'destroyed', listener: () => void): this
}

interface IpcMainLike {
  handle(
    channel: string,
    handler: (event: { sender: ComponentOwner }, ...args: unknown[]) => unknown,
  ): void
}

interface ComponentManagerLike {
  status(): Promise<ComponentStatus>
  install(options?: { readonly signal?: AbortSignal }): Promise<unknown>
  resolveExecutable(options?: { readonly signal?: AbortSignal }): Promise<string>
  remove(options?: { readonly signal?: AbortSignal }): Promise<void>
}

export interface RegisterEnhancedModeComponentIpcOptions {
  readonly ipcMain: IpcMainLike
  readonly component: EnhancedModeComponentManager | ComponentManagerLike
  readonly isTrustedSender: (owner: ComponentOwner) => boolean
  readonly readSavedMode: () => unknown
  readonly writeSavedMode: (mode: ProductAgentMode) => void
  readonly currentMode: () => ProductAgentMode
  readonly runtimeInUse: () => boolean
  readonly authorizeEnhanced: () => Promise<boolean>
  readonly policyAllowed: () => boolean
  readonly enhancedRuntimeAvailable: () => boolean
  readonly metadata?: Readonly<{
    platform: string
    bytes: number
    publisher: string
    license: string
    primaryUrl: string
    fallbackUrl: string
  }>
}

export interface EnhancedModeComponentController {
  close(): Promise<void>
  revoke(): Promise<void>
}

export class EnhancedModePublicError extends Error {
  constructor(readonly code: string) {
    super(code)
    this.name = 'EnhancedModePublicError'
  }
}

function publicFail(code: string): never {
  throw new EnhancedModePublicError(code)
}

function productMode(value: unknown): ProductAgentMode {
  // `codex` is a one-release internal migration value and is never returned or persisted here.
  return value === 'enhanced' || value === 'codex' ? 'enhanced' : 'standard'
}

export function registerEnhancedModeComponentIpc(
  options: RegisterEnhancedModeComponentIpcOptions,
): EnhancedModeComponentController {
  const active = new Set<Promise<unknown>>()
  const controllers = new Set<AbortController>()
  let closed = false

  const trusted = (owner: ComponentOwner) => {
    if (closed || owner.isDestroyed() || !options.isTrustedSender(owner)) {
      publicFail('enhanced_mode_untrusted_request')
    }
  }
  const status = async (): Promise<EnhancedModeStatus> => {
    let component: ComponentStatus
    try {
      component = await options.component.status()
    } catch {
      component = { state: 'invalid', supported: true, version: '0.147.0' }
    }
    const requestedAgentRuntime = productMode(options.readSavedMode())
    const activeAgentRuntime = options.currentMode()
    const lifecycleState = !options.policyAllowed()
      ? 'blocked_by_policy'
      : !component.supported || component.state === 'unsupported'
        ? 'unsupported_platform'
        : component.state === 'missing'
          ? 'not_installed'
          : component.state === 'invalid'
            ? 'failed_safe'
            : requestedAgentRuntime === 'enhanced' && !options.enhancedRuntimeAvailable()
              ? 'failed_safe'
              : requestedAgentRuntime !== activeAgentRuntime
                ? 'installed_restart_required'
                : 'ready'
    return {
      requestedAgentRuntime,
      activeAgentRuntime,
      component: component.state,
      supported: component.supported,
      version: component.version,
      restartRequired: requestedAgentRuntime !== activeAgentRuntime,
      lifecycleState,
      ...options.metadata,
    }
  }
  const operation = async <T>(owner: ComponentOwner, run: (signal: AbortSignal) => Promise<T>) => {
    trusted(owner)
    const controller = new AbortController()
    controllers.add(controller)
    const onDestroyed = () => controller.abort()
    owner.once('destroyed', onDestroyed)
    const promise = run(controller.signal)
    active.add(promise)
    try {
      const value = await promise
      return value
    } catch {
      publicFail(
        controller.signal.aborted ? 'enhanced_mode_cancelled' : 'enhanced_mode_operation_failed',
      )
    } finally {
      active.delete(promise)
      controllers.delete(controller)
      owner.removeListener('destroyed', onDestroyed)
    }
  }

  options.ipcMain.handle(ENHANCED_MODE_CHANNELS.status, async (event, ...args) => {
    trusted(event.sender)
    if (args.length !== 0) publicFail('enhanced_mode_invalid_request')
    return status()
  })
  options.ipcMain.handle(ENHANCED_MODE_CHANNELS.install, async (event, ...args) => {
    if (args.length !== 0) publicFail('enhanced_mode_invalid_request')
    trusted(event.sender)
    if (!options.policyAllowed()) publicFail('enhanced_mode_blocked_by_policy')
    if (!(await options.authorizeEnhanced())) publicFail('auth_required')
    await operation(event.sender, (signal) => options.component.install({ signal }))
    return status()
  })
  options.ipcMain.handle(ENHANCED_MODE_CHANNELS.update, async (event, ...args) => {
    if (args.length !== 0) publicFail('enhanced_mode_invalid_request')
    trusted(event.sender)
    if (!options.policyAllowed()) publicFail('enhanced_mode_blocked_by_policy')
    if (!(await options.authorizeEnhanced())) publicFail('auth_required')
    await operation(event.sender, (signal) => options.component.install({ signal }))
    return status()
  })
  options.ipcMain.handle(ENHANCED_MODE_CHANNELS.cancel, async (event, ...args) => {
    trusted(event.sender)
    if (args.length !== 0) publicFail('enhanced_mode_invalid_request')
    for (const controller of controllers) controller.abort()
    await Promise.allSettled([...active])
    return status()
  })
  options.ipcMain.handle(ENHANCED_MODE_CHANNELS.remove, async (event, ...args) => {
    trusted(event.sender)
    if (args.length !== 0) publicFail('enhanced_mode_invalid_request')
    if (options.runtimeInUse()) publicFail('enhanced_mode_restart_required')
    await operation(event.sender, (signal) => options.component.remove({ signal }))
    return status()
  })
  options.ipcMain.handle(ENHANCED_MODE_CHANNELS.setMode, async (event, ...args) => {
    trusted(event.sender)
    if (args.length !== 1 || (args[0] !== 'standard' && args[0] !== 'enhanced')) {
      publicFail('enhanced_mode_invalid_request')
    }
    const mode = args[0] as ProductAgentMode
    if (mode === 'enhanced') {
      if (!options.policyAllowed()) publicFail('enhanced_mode_blocked_by_policy')
      if (!(await options.authorizeEnhanced())) publicFail('auth_required')
      try {
        await operation(event.sender, (signal) => options.component.resolveExecutable({ signal }))
      } catch (error) {
        if (error instanceof EnhancedModePublicError && error.code === 'enhanced_mode_cancelled') {
          throw error
        }
        publicFail('enhanced_mode_install_required')
      }
    }
    options.writeSavedMode(mode)
    return status()
  })

  return {
    async revoke() {
      for (const controller of controllers) controller.abort()
      await Promise.allSettled([...active])
      options.writeSavedMode('standard')
    },
    async close() {
      if (closed) return
      closed = true
      for (const controller of controllers) controller.abort()
      await Promise.allSettled([...active])
    },
  }
}
