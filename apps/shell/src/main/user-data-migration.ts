import { cpSync, mkdtempSync, readdirSync, renameSync, rmdirSync, rmSync, statSync } from 'node:fs'
import { basename, dirname, join, resolve } from 'node:path'

const LEGACY_PRODUCT_NAMES = ['GenOffice', 'AI Office'] as const

type TargetState = 'missing' | 'empty' | 'occupied'

export interface UserDataMigrationOperations {
  copyDirectory(source: string, destination: string): void
  createTemporaryDirectory(prefix: string): string
  removeEmptyDirectory(path: string): void
  removeTemporaryDirectory(path: string): void
  rename(source: string, destination: string): void
}

const defaultOperations: UserDataMigrationOperations = {
  copyDirectory: (source, destination) => cpSync(source, destination, { recursive: true }),
  createTemporaryDirectory: (prefix) => mkdtempSync(prefix),
  removeEmptyDirectory: (path) => rmdirSync(path),
  removeTemporaryDirectory: (path) => rmSync(path, { recursive: true, force: true }),
  rename: (source, destination) => renameSync(source, destination),
}

function targetState(path: string): TargetState {
  try {
    const stat = statSync(path)
    if (!stat.isDirectory()) return 'occupied'
    return readdirSync(path).length === 0 ? 'empty' : 'occupied'
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return 'missing'
    throw error
  }
}

function existingLegacyDirectory(appDataDir: string): string | null {
  for (const productName of LEGACY_PRODUCT_NAMES) {
    const candidate = join(appDataDir, productName)
    try {
      if (statSync(candidate).isDirectory()) return candidate
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
    }
  }
  return null
}

/**
 * Atomically copy the first available legacy profile into a fresh WisWork profile.
 * The source is retained, and a failed or competing migration never exposes a partial target.
 */
export function migrateLegacyUserData(
  appDataDir: string,
  targetDir: string,
  operationOverrides: Partial<Pick<UserDataMigrationOperations, 'copyDirectory' | 'rename'>> = {},
): void {
  const expectedTarget = resolve(appDataDir, 'WisWork')
  if (
    resolve(targetDir) !== expectedTarget ||
    resolve(dirname(targetDir)) !== resolve(appDataDir)
  ) {
    throw new Error('WisWork userData migration target must be the direct appData/WisWork child')
  }

  if (targetState(targetDir) === 'occupied') return
  const sourceDir = existingLegacyDirectory(appDataDir)
  if (!sourceDir) return

  const operations = { ...defaultOperations, ...operationOverrides }
  const temporaryPrefix = join(dirname(targetDir), `.${basename(targetDir)}-migration-`)
  const temporaryDir = operations.createTemporaryDirectory(temporaryPrefix)
  let promoted = false
  try {
    operations.copyDirectory(sourceDir, temporaryDir)
    const beforePromotion = targetState(targetDir)
    if (beforePromotion === 'occupied') return
    if (beforePromotion === 'empty') {
      try {
        operations.removeEmptyDirectory(targetDir)
      } catch (error) {
        if (targetState(targetDir) === 'occupied') return
        throw error
      }
    }
    try {
      operations.rename(temporaryDir, targetDir)
      promoted = true
    } catch (error) {
      if (targetState(targetDir) === 'occupied') return
      throw error
    }
  } finally {
    if (!promoted) operations.removeTemporaryDirectory(temporaryDir)
  }
}
