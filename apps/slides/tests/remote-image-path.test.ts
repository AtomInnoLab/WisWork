import { describe, expect, it } from 'vitest'
import source from '../src/main/ai-ipc.ts?raw'

describe('Slides remote image download path', () => {
  it('uses the retrying image downloader for insertion and replacement', () => {
    expect(source).toContain("import { fetchRemoteImage } from '@wiswork/electron-utils'")
    expect(source.match(/await fetchRemoteImage\(/g)).toHaveLength(2)
    expect(source).not.toContain('await fetchWithSsrfGuard(')
  })
})
