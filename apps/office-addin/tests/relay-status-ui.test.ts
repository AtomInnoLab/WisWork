import { describe, expect, it } from 'vitest'
import { relayConnectionPresentation } from '../src/App.js'

describe('persistent relay status UI', () => {
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
})
