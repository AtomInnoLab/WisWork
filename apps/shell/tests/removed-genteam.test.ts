import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

const appRoot = resolve(import.meta.dirname, '..')

function source(path: string): string {
  return readFileSync(resolve(appRoot, path), 'utf8')
}

describe('removed GenTeam integration', () => {
  it('has no renderer, preload, shared IPC, or main-process entry to the old site', () => {
    const combined = [
      source('src/main/index.ts'),
      source('src/preload/index.ts'),
      source('src/shared/home-api.ts'),
      source('src/renderer/src/Onboarding.tsx'),
    ].join('\n')
    expect(combined).not.toContain('openGenTeam')
    expect(combined).not.toContain('home:open-genteam')
    expect(combined).not.toContain('wiswork.ai/wiswork/join')
  })
})
