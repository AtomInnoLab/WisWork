import {
  saveMasterBackup,
  readMasterBackup,
  releaseMasterBackup,
  masterBackupRefForBytes,
  type MasterBackupRef,
} from './presentation-master-backup.js'
export type PackageBackupRef = MasterBackupRef
export const packageBackupRefForBytes = masterBackupRefForBytes
async function packageError<T>(run: () => Promise<T>): Promise<T> {
  try {
    return await run()
  } catch (error) {
    if (error instanceof Error && error.message.startsWith('presentation_master_backup_'))
      throw new Error(
        error.message.replace('presentation_master_backup_', 'presentation_package_backup_'),
        { cause: error },
      )
    throw error
  }
}
export function savePackageBackup(
  input: Omit<Parameters<typeof saveMasterBackup>[0], 'protocol'>,
): Promise<PackageBackupRef> {
  return packageError(() => saveMasterBackup({ ...input, protocol: 'package' }))
}
export function readPackageBackup(
  input: Omit<Parameters<typeof readMasterBackup>[0], 'protocol'>,
): Promise<Uint8Array> {
  return packageError(() => readMasterBackup({ ...input, protocol: 'package' }))
}
export function releasePackageBackup(
  input: Omit<Parameters<typeof releaseMasterBackup>[0], 'protocol'>,
): Promise<void> {
  return packageError(() => releaseMasterBackup({ ...input, protocol: 'package' }))
}
