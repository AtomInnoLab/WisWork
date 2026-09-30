import { createPresentationMasterBackupService } from './presentation-master-backups'
const operations = new Set(
  ['begin', 'chunk', 'finish', 'status', 'read', 'list', 'release', 'inventory'].map(
    (x) => `package_backup_${x}`,
  ),
)
export function createPresentationPackageBackupService(options: { userDataPath: string }) {
  const service = createPresentationMasterBackupService({
    userDataPath: options.userDataPath,
    storageDirectory: 'presentation-package-backups',
  })
  return async (input: Record<string, unknown>, signal: AbortSignal): Promise<unknown> => {
    const body = structuredClone(input)
    if (typeof body.operation !== 'string' || !operations.has(body.operation))
      throw new Error('presentation_package_backup_invalid')
    body.operation = body.operation.replace('package_backup_', 'master_backup_')
    try {
      return await service(body, signal)
    } catch (error) {
      if (error instanceof Error && error.message.startsWith('presentation_master_backup_'))
        throw new Error(
          error.message.replace('presentation_master_backup_', 'presentation_package_backup_'),
          { cause: error },
        )
      throw error
    }
  }
}
