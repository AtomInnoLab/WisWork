import { EventEmitter } from 'node:events'
import { describe, expect, it, vi } from 'vitest'
import {
  registerEnhancedModeComponentIpc,
  type EnhancedModeComponentController,
} from '../src/main/enhanced-mode-component'
import { ENHANCED_MODE_CHANNELS } from '../src/shared/enhanced-mode-api'

class FakeSender extends EventEmitter {
  destroyed = false

  isDestroyed(): boolean {
    return this.destroyed
  }

  destroy(): void {
    this.destroyed = true
    this.emit('destroyed')
  }
}

function fixture(options: { runtimeInUse?: boolean } = {}) {
  const handlers = new Map<string, (event: { sender: FakeSender }, ...args: unknown[]) => unknown>()
  const component = {
    status: vi.fn(async () => ({
      state: 'missing' as const,
      supported: true,
      version: '0.147.0' as const,
    })),
    install: vi.fn(async () => ({
      executablePath: '/private/component/bin/codex',
      version: '0.147.0' as const,
      platform: 'darwin' as const,
      arch: 'arm64' as const,
    })),
    resolveExecutable: vi.fn(async () => '/private/component/bin/codex'),
    remove: vi.fn(async () => undefined),
  }
  let savedMode: unknown = 'standard'
  const writeMode = vi.fn((mode: 'standard' | 'enhanced') => {
    savedMode = mode
  })
  const sender = new FakeSender()
  const controller = registerEnhancedModeComponentIpc({
    ipcMain: { handle: (channel, handler) => handlers.set(channel, handler as never) },
    component,
    isTrustedSender: (owner) => owner === sender,
    readSavedMode: () => savedMode,
    writeSavedMode: writeMode,
    currentMode: () => (options.runtimeInUse === true ? 'enhanced' : 'standard'),
    runtimeInUse: () => options.runtimeInUse === true,
  })
  return {
    handlers,
    component,
    sender,
    controller,
    writeMode,
    get savedMode() {
      return savedMode
    },
  }
}

describe('Enhanced mode optional component IPC', () => {
  it('reports product-safe Standard mode status without installing or enabling anything', async () => {
    const f = fixture()
    await expect(
      f.handlers.get(ENHANCED_MODE_CHANNELS.status)!({ sender: f.sender }),
    ).resolves.toEqual({
      mode: 'standard',
      component: 'missing',
      supported: true,
      version: '0.147.0',
      restartRequired: false,
    })
    expect(f.component.install).not.toHaveBeenCalled()
    expect(f.writeMode).not.toHaveBeenCalled()
    expect(
      JSON.stringify(await f.handlers.get(ENHANCED_MODE_CHANNELS.status)!({ sender: f.sender })),
    ).not.toMatch(/codex/i)
  })

  it('downloads only after an explicit trusted request and does not silently enable it', async () => {
    const f = fixture()
    f.component.status.mockResolvedValue({ state: 'ready', supported: true, version: '0.147.0' })
    await expect(
      f.handlers.get(ENHANCED_MODE_CHANNELS.install)!({ sender: f.sender }),
    ).resolves.toMatchObject({
      mode: 'standard',
      component: 'ready',
    })
    expect(f.component.install).toHaveBeenCalledOnce()
    expect(f.writeMode).not.toHaveBeenCalled()

    const attacker = new FakeSender()
    await expect(
      f.handlers.get(ENHANCED_MODE_CHANNELS.install)!({ sender: attacker }),
    ).rejects.toThrow('enhanced_mode_untrusted_request')
  })

  it('enables only after a fresh integrity/version resolution and persists product mode names', async () => {
    const f = fixture()
    f.component.status.mockResolvedValue({ state: 'ready', supported: true, version: '0.147.0' })
    await expect(
      f.handlers.get(ENHANCED_MODE_CHANNELS.setMode)!({ sender: f.sender }, 'enhanced'),
    ).resolves.toMatchObject({ mode: 'enhanced', component: 'ready', restartRequired: true })
    expect(f.component.resolveExecutable).toHaveBeenCalledOnce()
    expect(f.writeMode).toHaveBeenCalledWith('enhanced')
    expect(f.savedMode).toBe('enhanced')

    f.component.resolveExecutable.mockRejectedValueOnce(new Error('private integrity detail'))
    await expect(
      f.handlers.get(ENHANCED_MODE_CHANNELS.setMode)!({ sender: f.sender }, 'enhanced'),
    ).rejects.toThrow('enhanced_mode_install_required')
    expect(f.writeMode).toHaveBeenCalledTimes(1)
  })

  it('supports flag-only rollback and refuses removal while the current process uses the component', async () => {
    const f = fixture({ runtimeInUse: true })
    await expect(
      f.handlers.get(ENHANCED_MODE_CHANNELS.setMode)!({ sender: f.sender }, 'standard'),
    ).resolves.toMatchObject({ mode: 'standard', restartRequired: true })
    expect(f.writeMode).toHaveBeenCalledWith('standard')
    await expect(
      f.handlers.get(ENHANCED_MODE_CHANNELS.remove)!({ sender: f.sender }),
    ).rejects.toThrow('enhanced_mode_restart_required')
    expect(f.component.remove).not.toHaveBeenCalled()
  })

  it('cancels an in-flight download when its renderer or controller closes', async () => {
    const f = fixture()
    let signal: AbortSignal | undefined
    f.component.install.mockImplementationOnce(async (options) => {
      signal = options?.signal
      await new Promise<void>((_resolve, reject) =>
        options?.signal?.addEventListener('abort', () => reject(new Error('cancelled')), {
          once: true,
        }),
      )
      throw new Error('unreachable')
    })
    const installing = expect(
      f.handlers.get(ENHANCED_MODE_CHANNELS.install)!({ sender: f.sender }),
    ).rejects.toThrow('enhanced_mode_cancelled')
    await Promise.resolve()
    f.sender.destroy()
    await installing
    expect(signal?.aborted).toBe(true)

    await expect((f.controller as EnhancedModeComponentController).close()).resolves.toBeUndefined()
  })

  it('rejects malformed arguments with bounded stable public errors', async () => {
    const f = fixture()
    await expect(
      f.handlers.get(ENHANCED_MODE_CHANNELS.status)!({ sender: f.sender }, 'extra'),
    ).rejects.toThrow('enhanced_mode_invalid_request')
    await expect(
      f.handlers.get(ENHANCED_MODE_CHANNELS.setMode)!({ sender: f.sender }, 'codex'),
    ).rejects.toThrow('enhanced_mode_invalid_request')
  })
})
