import type { OfficeBridge } from '@wiswork/office-bridge'
import { OFFICE_PAIRING_CHANNELS } from '../shared/home-api'

const PAIRING_ID = /^[A-Za-z0-9_-]{8,128}$/

function pairingId(input: unknown): string {
  const value = (input as { pairingId?: unknown } | null)?.pairingId
  if (
    !input ||
    typeof input !== 'object' ||
    Object.keys(input).length !== 1 ||
    typeof value !== 'string' ||
    !PAIRING_ID.test(value)
  )
    throw new Error('Invalid pairing IPC payload.')
  return value
}

export function registerOfficePairingIpc(options: {
  ipcMain: {
    handle(
      name: string,
      listener: (event: { sender: unknown }, ...args: unknown[]) => unknown,
    ): unknown
  }
  bridge: Pick<OfficeBridge, 'approve' | 'reject'>
  listPending(): ReturnType<OfficeBridge['listPending']>
  getValidAccountStatus(): Promise<{ loggedIn: boolean }>
  isTrustedSender(sender: unknown): boolean
}): void {
  const validate = (event: { sender: unknown }, args: unknown[]) => {
    if (!options.isTrustedSender(event.sender)) throw new Error('Untrusted IPC sender.')
    if (args.length !== 1) throw new Error('Invalid pairing IPC payload.')
    return pairingId(args[0])
  }
  options.ipcMain.handle(OFFICE_PAIRING_CHANNELS.list, (event, ...args) => {
    if (!options.isTrustedSender(event.sender)) throw new Error('Untrusted IPC sender.')
    if (args.length !== 0) throw new Error('Invalid pairing IPC payload.')
    return options.listPending()
  })
  options.ipcMain.handle(OFFICE_PAIRING_CHANNELS.approve, async (event, ...args) => {
    const id = validate(event, args)
    const status = await options.getValidAccountStatus()
    return status.loggedIn ? options.bridge.approve(id, true) : false
  })
  options.ipcMain.handle(OFFICE_PAIRING_CHANNELS.reject, (event, ...args) =>
    options.bridge.reject(validate(event, args)),
  )
}
