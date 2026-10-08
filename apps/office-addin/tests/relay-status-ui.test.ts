import { describe, expect, it } from 'vitest'
import {
  relayConnectionPresentation,
  relayPersistenceNotice,
  shouldShowRelayStatusScreen,
  shouldResetOfficeSession,
} from '../src/App.js'

describe('persistent relay status UI', () => {
  it('preserves the active conversation across transient connection states', () => {
    expect(shouldResetOfficeSession('reconnecting')).toBe(false)
    expect(shouldResetOfficeSession('connecting')).toBe(false)
    expect(shouldResetOfficeSession('waiting_for_pc')).toBe(false)
    expect(shouldResetOfficeSession('offline')).toBe(false)
    expect(shouldResetOfficeSession('rejected')).toBe(true)
    expect(shouldResetOfficeSession('expired')).toBe(true)
  })

  it('keeps an existing workspace visible while the relay reconnects', () => {
    expect(shouldShowRelayStatusScreen('reconnecting', true)).toBe(false)
    expect(shouldShowRelayStatusScreen('waiting_for_pc', true)).toBe(false)
    expect(shouldShowRelayStatusScreen('offline', true)).toBe(true)
    expect(shouldShowRelayStatusScreen('reconnecting', false)).toBe(true)
    expect(shouldShowRelayStatusScreen('rejected', true)).toBe(true)
  })

  it('shows bounded reconnect progress without offering a new pairing action', () => {
    expect(relayConnectionPresentation('reconnecting')).toEqual({
      title: 'Reconnecting to WisWork PC…',
      detail: 'Reconnecting to WisWork PC…',
      busy: true,
      actionDisabled: true,
    })
  })

  it('keeps an offline PC in the remembered waiting state without a new code', () => {
    expect(relayConnectionPresentation('waiting_for_pc')).toEqual({
      title: 'Waiting for WisWork PC',
      detail: 'Waiting for a signed-in WisWork PC.',
      busy: false,
      actionDisabled: false,
    })
  })

  it('keeps the six-digit prompt for ordinary one-time pairing', () => {
    expect(relayConnectionPresentation('pending', '123456')).toEqual({
      title: 'Connect to WisWork PC',
      detail: 'Enter code 123456 in WisWork PC, then approve the matching request.',
      busy: true,
      actionDisabled: true,
    })
  })

  it('surfaces an enhanced enrollment that connected without being remembered', () => {
    expect(
      relayPersistenceNotice({
        status: 'connected',
        capabilities: ['agent.v1'],
        remembered: false,
      }),
    ).toBe(
      'Connected, but this Office installation was not remembered. Pair again after reconnecting.',
    )
    expect(relayPersistenceNotice({ status: 'connected', capabilities: ['agent.v1'] })).toBe(
      undefined,
    )
  })
})
