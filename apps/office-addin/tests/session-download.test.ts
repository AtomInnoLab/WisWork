// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
import { downloadSessionFile } from '../src/agent/session-download.js'

afterEach(() => {
  vi.restoreAllMocks()
  vi.useRealTimers()
})
describe('session file download', () => {
  it('downloads the current VFS artifact locally and releases its blob URL', () => {
    vi.useFakeTimers()
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/deck.pptx', new Uint8Array([80, 75, 3, 4]))
    const create = vi.fn((_blob: Blob) => 'blob:generated')
    const revoke = vi.fn()
    vi.stubGlobal(
      'URL',
      class extends URL {
        static createObjectURL = create
        static revokeObjectURL = revoke
      },
    )
    let downloaded = ''
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (
      this: HTMLAnchorElement,
    ) {
      downloaded = this.download
      expect(this.href).toBe('blob:generated')
    })
    downloadSessionFile(vfs, '/home/user/deck.pptx')
    expect(downloaded).toBe('deck.pptx')
    expect(create.mock.calls[0]?.[0]).toBeInstanceOf(Blob)
    vi.runAllTimers()
    expect(revoke).toHaveBeenCalledWith('blob:generated')
    vi.unstubAllGlobals()
  })
  it('does not export files outside session attachments', () => {
    const vfs = new InMemoryVfs()
    expect(() => downloadSessionFile(vfs, '/etc/passwd')).toThrow('vfs_path_denied')
  })
})
