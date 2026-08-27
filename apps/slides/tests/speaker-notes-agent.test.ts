import { describe, expect, it, vi } from 'vitest'
import { setAndVerifySpeakerNotes } from '../src/renderer/ai/speaker-notes'

describe('setAndVerifySpeakerNotes', () => {
  it('returns unchanged when the host rejects the write', async () => {
    const getNotes = vi.fn(async () => 'old')

    await expect(
      setAndVerifySpeakerNotes({ setNotes: async () => false, getNotes }, 0, 'new'),
    ).resolves.toBe('unchanged')
    expect(getNotes).not.toHaveBeenCalled()
  })

  it('returns applied only after exact readback', async () => {
    await expect(
      setAndVerifySpeakerNotes(
        { setNotes: async () => true, getNotes: async () => 'new' },
        0,
        'new',
      ),
    ).resolves.toBe('applied')
  })

  it('returns uncertain when a successful host write cannot be verified', async () => {
    await expect(
      setAndVerifySpeakerNotes(
        { setNotes: async () => true, getNotes: async () => 'host-normalized' },
        0,
        'new',
      ),
    ).resolves.toBe('uncertain')
  })
})
