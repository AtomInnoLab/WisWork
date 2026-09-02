import { describe, expect, it } from 'vitest'
import { enhancedModeView } from '../src/renderer/src/enhanced-mode-view'

describe('Enhanced mode user-facing copy', () => {
  it('uses only Standard mode and Enhanced mode product naming in English', () => {
    const missingStatus = {
      requestedAgentRuntime: 'standard',
      activeAgentRuntime: 'standard',
      component: 'missing' as const,
      supported: true,
      version: '0.147.0',
      restartRequired: false,
    } as const
    const missing = enhancedModeView(missingStatus, 'en')
    expect(missing).toMatchObject({
      label: 'Enhanced mode',
      detail: 'Optional download required',
      action: 'install',
      actionLabel: 'Download',
    })
    expect(JSON.stringify(missing)).not.toMatch(/codex/i)

    const ready = enhancedModeView({ ...missingStatus, component: 'ready' }, 'en')
    expect(ready).toMatchObject({
      detail: 'Standard mode',
      action: 'enable',
      secondaryAction: 'remove',
      secondaryActionLabel: 'Remove optional component',
    })
    const pendingRestart = enhancedModeView(
      { ...missingStatus, component: 'ready', restartRequired: true },
      'en',
    )
    expect(pendingRestart).toMatchObject({ action: 'enable' })
    expect(pendingRestart.secondaryAction).toBeUndefined()
  })

  it('uses the required Standard/Enhanced mode naming in Chinese', () => {
    const view = enhancedModeView(
      {
        requestedAgentRuntime: 'enhanced',
        activeAgentRuntime: 'enhanced',
        component: 'ready',
        supported: true,
        version: '0.147.0',
        restartRequired: false,
      },
      'zh',
    )
    expect(view).toMatchObject({
      label: '增强模式',
      detail: '增强模式',
      action: 'disable',
      actionLabel: '切换到标准模式',
    })
    expect(JSON.stringify(view)).not.toMatch(/Codex/i)
  })

  it('fails closed in UI for unsupported, invalid, and selected-but-missing states', () => {
    const base = {
      requestedAgentRuntime: 'enhanced' as const,
      activeAgentRuntime: 'enhanced' as const,
      component: 'missing' as const,
      supported: true,
      version: '0.147.0',
      restartRequired: false,
    }
    expect(enhancedModeView(base, 'en')).toMatchObject({
      detail: 'Install required before use',
      action: 'install',
    })
    expect(enhancedModeView({ ...base, component: 'invalid' }, 'en')).toMatchObject({
      detail: 'Verification failed',
      action: 'install',
    })
    expect(
      enhancedModeView({ ...base, component: 'unsupported', supported: false }, 'en'),
    ).toMatchObject({ detail: 'Not supported on this device', action: 'none' })
  })
})
