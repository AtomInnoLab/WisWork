import { describe, expect, it, vi } from 'vitest'

import { OFFICE_PAIRING_CHANNELS } from '../src/shared/home-api'
import { registerOfficePairingIpc } from '../src/main/office-pairing-ipc'

describe('Office pairing IPC', () => {
  it('validates sender and payload, checks account before approval, and supports rejection', async () => {
    const handlers = new Map<string, (...args: unknown[]) => unknown>()
    const ipcMain = {
      handle: vi.fn((name: string, fn: (...args: unknown[]) => unknown) => handlers.set(name, fn)),
    }
    const trusted = { id: 1 }
    const bridge = { approve: vi.fn(() => true), reject: vi.fn(() => true) }
    const getValidAccountStatus = vi.fn(async () => ({ loggedIn: true }))
    registerOfficePairingIpc({
      ipcMain,
      bridge,
      getValidAccountStatus,
      isTrustedSender: (sender) => sender === trusted,
    })
    const approve = handlers.get(OFFICE_PAIRING_CHANNELS.approve)!
    await expect(approve({ sender: {} }, { pairingId: 'valid_pairing-id' })).rejects.toThrow(
      'Untrusted IPC sender',
    )
    await expect(approve({ sender: trusted }, { pairingId: '../bad' })).rejects.toThrow(
      'Invalid pairing IPC payload',
    )
    await expect(approve({ sender: trusted }, { pairingId: 'valid_pairing-id' })).resolves.toBe(
      true,
    )
    expect(getValidAccountStatus).toHaveBeenCalledOnce()
    expect(bridge.approve).toHaveBeenCalledWith('valid_pairing-id', true)

    getValidAccountStatus.mockResolvedValue({ loggedIn: false })
    await expect(approve({ sender: trusted }, { pairingId: 'another_pairing' })).resolves.toBe(
      false,
    )
    expect(bridge.approve).not.toHaveBeenCalledWith('another_pairing', true)

    const reject = handlers.get(OFFICE_PAIRING_CHANNELS.reject)!
    expect(await reject({ sender: trusted }, { pairingId: 'valid_pairing-id' })).toBe(true)
  })
})
