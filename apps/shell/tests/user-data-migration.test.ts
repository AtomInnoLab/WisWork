import {
  cpSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { migrateLegacyUserData } from '../src/main/user-data-migration'

const roots: string[] = []

afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})

function fixture(): string {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-user-data-'))
  roots.push(root)
  return root
}

describe('migrateLegacyUserData', () => {
  it('copies GenOffice data into an empty WisWork directory without deleting the source', () => {
    const root = fixture()
    const legacy = join(root, 'GenOffice')
    const target = join(root, 'WisWork')
    writeFileSync(join(root, 'sentinel'), 'outside')
    expect(() => migrateLegacyUserData(root, target)).not.toThrow()

    // A missing legacy directory is a no-op.
    expect(existsSync(target)).toBe(false)

    mkdirSync(legacy, { recursive: true })
    writeFileSync(join(legacy, 'settings.json'), '{"theme":"dark"}')
    migrateLegacyUserData(root, target)

    expect(readFileSync(join(target, 'settings.json'), 'utf8')).toBe('{"theme":"dark"}')
    expect(readFileSync(join(legacy, 'settings.json'), 'utf8')).toBe('{"theme":"dark"}')
  })

  it('falls back to AI Office and never overwrites a non-empty WisWork directory', () => {
    const root = fixture()
    const legacy = join(root, 'AI Office')
    const target = join(root, 'WisWork')
    mkdirSync(legacy, { recursive: true })
    writeFileSync(join(legacy, 'legacy.json'), 'legacy')

    migrateLegacyUserData(root, target)

    expect(readFileSync(join(target, 'legacy.json'), 'utf8')).toBe('legacy')
    expect(readFileSync(join(legacy, 'legacy.json'), 'utf8')).toBe('legacy')

    writeFileSync(join(target, 'current.json'), 'current')
    writeFileSync(join(legacy, 'legacy.json'), 'changed')
    migrateLegacyUserData(root, target)

    expect(readFileSync(join(target, 'current.json'), 'utf8')).toBe('current')
    expect(readFileSync(join(target, 'legacy.json'), 'utf8')).toBe('legacy')
    expect(readFileSync(join(legacy, 'legacy.json'), 'utf8')).toBe('changed')
  })

  it('cleans an interrupted temporary copy, propagates the error, and succeeds on retry', () => {
    const root = fixture()
    const legacy = join(root, 'GenOffice')
    const target = join(root, 'WisWork')
    mkdirSync(legacy, { recursive: true })
    writeFileSync(join(legacy, 'settings.json'), 'source')
    expect(() =>
      migrateLegacyUserData(root, target, {
        copyDirectory(_source, temporary) {
          writeFileSync(join(temporary, 'partial.json'), 'partial')
          throw new Error('injected copy interruption')
        },
      }),
    ).toThrow('injected copy interruption')
    expect(existsSync(target)).toBe(false)
    expect(readFileSync(join(legacy, 'settings.json'), 'utf8')).toBe('source')
    expect(readdirSync(root).filter((entry) => entry.startsWith('.WisWork-migration-'))).toEqual([])
    migrateLegacyUserData(root, target)
    expect(readFileSync(join(target, 'settings.json'), 'utf8')).toBe('source')
  })

  it('promotes over Electron pre-created empty userData without exposing a partial target', () => {
    const root = fixture()
    const legacy = join(root, 'GenOffice')
    const target = join(root, 'WisWork')
    mkdirSync(legacy, { recursive: true })
    mkdirSync(target)
    writeFileSync(join(legacy, 'settings.json'), 'source')
    let targetVisibleDuringCopy = false
    migrateLegacyUserData(root, target, {
      copyDirectory(source, temporary) {
        targetVisibleDuringCopy = existsSync(join(target, 'settings.json'))
        cpSync(source, temporary, { recursive: true })
      },
    })
    expect(targetVisibleDuringCopy).toBe(false)
    expect(readFileSync(join(target, 'settings.json'), 'utf8')).toBe('source')
  })

  it('loses a concurrent rename safely without overwriting the winning profile', () => {
    const root = fixture()
    const legacy = join(root, 'GenOffice')
    const target = join(root, 'WisWork')
    mkdirSync(legacy, { recursive: true })
    writeFileSync(join(legacy, 'settings.json'), 'source')
    expect(() =>
      migrateLegacyUserData(root, target, {
        rename() {
          mkdirSync(target)
          writeFileSync(join(target, 'winner.json'), 'winner')
          const conflict = new Error('rename conflict') as NodeJS.ErrnoException
          conflict.code = 'EEXIST'
          throw conflict
        },
      }),
    ).not.toThrow()
    expect(readFileSync(join(target, 'winner.json'), 'utf8')).toBe('winner')
    expect(existsSync(join(target, 'settings.json'))).toBe(false)
    expect(readFileSync(join(legacy, 'settings.json'), 'utf8')).toBe('source')
    expect(readdirSync(root).filter((entry) => entry.startsWith('.WisWork-migration-'))).toEqual([])
  })

  it('rejects any migration target outside the direct appData/WisWork child', () => {
    const root = fixture()
    mkdirSync(join(root, 'GenOffice'))
    expect(() => migrateLegacyUserData(root, join(root, 'Other'))).toThrow(
      'direct appData/WisWork child',
    )
  })

  it('runs only after the single-instance lock and before userData consumers', () => {
    const main = readFileSync(join(import.meta.dirname, '../src/main/index.ts'), 'utf8')
    const ready = main.slice(main.indexOf('app.whenReady().then'))
    expect(ready.indexOf('requestSingleInstanceLock')).toBeGreaterThanOrEqual(0)
    expect(ready.indexOf('migrateLegacyUserData')).toBeGreaterThan(
      ready.indexOf('requestSingleInstanceLock'),
    )
    expect(ready.indexOf('initializeElectronAuthRuntime')).toBeGreaterThan(
      ready.indexOf('migrateLegacyUserData'),
    )
    expect(ready.indexOf('currentLang()')).toBeGreaterThan(ready.indexOf('migrateLegacyUserData'))
    expect(ready.indexOf('createShellWindow()')).toBeGreaterThan(
      ready.indexOf('migrateLegacyUserData'),
    )
  })
})
