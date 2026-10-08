import { describe, expect, it, vi } from 'vitest'
import { createOfficeImageHandoff } from '../src/main/office-image-handoff'

type DecodedImage = ReturnType<Parameters<typeof createOfficeImageHandoff>[0]['createFromBuffer']>
const LIMIT = 180 * 1024

function decoded(overrides: Partial<DecodedImage> = {}): DecodedImage {
  return {
    isEmpty: () => false,
    getSize: () => ({ width: 800, height: 600 }),
    resize: vi.fn(() => decoded()),
    toPNG: vi.fn(() => new Uint8Array(LIMIT + 1)),
    toJPEG: vi.fn(() => new Uint8Array(LIMIT + 1)),
    ...overrides,
  }
}

describe('private Office image handoff', () => {
  it.each(['image/png', 'image/jpeg'] as const)(
    'preserves a fitting %s without re-encoding',
    async (mime) => {
      const image = decoded()
      const input = { mime, bytes: new Uint8Array(LIMIT) }
      const prepare = createOfficeImageHandoff({ createFromBuffer: () => image })
      expect(await prepare(input)).toEqual(input)
      expect(image.resize).not.toHaveBeenCalled()
      expect(image.toPNG).not.toHaveBeenCalled()
      expect(image.toJPEG).not.toHaveBeenCalled()
    },
  )

  it('keeps PNG transparency when a lossless encoding fits', async () => {
    const bytes = new Uint8Array(LIMIT)
    const image = decoded({ toPNG: vi.fn(() => bytes) })
    const prepare = createOfficeImageHandoff({ createFromBuffer: () => image })
    expect(await prepare({ mime: 'image/png', bytes: new Uint8Array(LIMIT + 1) })).toEqual({
      mime: 'image/png',
      bytes,
    })
    expect(image.toJPEG).not.toHaveBeenCalled()
  })

  it('falls back to JPEG when PNG cannot fit the private payload', async () => {
    const bytes = new Uint8Array(LIMIT - 1)
    const image = decoded({ toJPEG: vi.fn(() => bytes) })
    const prepare = createOfficeImageHandoff({ createFromBuffer: () => image })
    expect(await prepare({ mime: 'image/png', bytes: new Uint8Array(LIMIT + 1) })).toEqual({
      mime: 'image/jpeg',
      bytes,
    })
  })

  it('reduces dimensions without stretching and fits below the base64 envelope', async () => {
    const bytes = new Uint8Array(LIMIT)
    const image = decoded({
      getSize: () => ({ width: 3200, height: 1600 }),
      resize: vi.fn(({ width, height }) => {
        expect(width).toBe(height * 2)
        return decoded({ toJPEG: () => (width <= 1024 ? bytes : new Uint8Array(LIMIT + 1)) })
      }),
    })
    const result = await createOfficeImageHandoff({ createFromBuffer: () => image })({
      mime: 'image/jpeg',
      bytes: new Uint8Array(4 * 1024 * 1024),
    })
    expect(result).toEqual({ mime: 'image/jpeg', bytes })
    expect(Buffer.byteLength(Buffer.from(result.bytes).toString('base64'))).toBeLessThan(256 * 1024)
    expect(image.resize).toHaveBeenLastCalledWith({ width: 1024, height: 512, quality: 'best' })
  })

  it.each([
    { width: 0, height: 1 },
    { width: Number.NaN, height: 1 },
    { width: 8193, height: 1 },
    { width: 5000, height: 5000 },
  ])('rejects unsafe dimensions even for fitting input: %o', async (size) => {
    const prepare = createOfficeImageHandoff({
      createFromBuffer: () => decoded({ getSize: () => size }),
    })
    await expect(prepare({ mime: 'image/png', bytes: new Uint8Array(100) })).rejects.toThrow(
      'image_limit',
    )
  })

  it('rejects oversized source bytes before decoding', async () => {
    const createFromBuffer = vi.fn(() => decoded())
    const prepare = createOfficeImageHandoff({ createFromBuffer })
    await expect(
      prepare({ mime: 'image/png', bytes: new Uint8Array(10 * 1024 * 1024 + 1) }),
    ).rejects.toThrow('image_limit')
    expect(createFromBuffer).not.toHaveBeenCalled()
  })

  it('does not shrink a large photo to a blurry thumbnail just to fit transport', async () => {
    const sizes: number[] = []
    const image = decoded({
      getSize: () => ({ width: 3200, height: 1800 }),
      resize: ({ width, height }) => {
        sizes.push(width)
        return decoded({
          getSize: () => ({ width, height }),
          toJPEG: () => new Uint8Array(width < 960 ? LIMIT : LIMIT + 1),
        })
      },
    })
    await expect(
      createOfficeImageHandoff({ createFromBuffer: () => image })({
        mime: 'image/jpeg',
        bytes: new Uint8Array(LIMIT + 1),
      }),
    ).rejects.toThrow('image_limit')
    expect(sizes.every((size) => size >= 960)).toBe(true)
  })

  it('rejects empty or undecodable images without forwarding them', async () => {
    for (const bytes of [new Uint8Array(), new Uint8Array(100)]) {
      const prepare = createOfficeImageHandoff({
        createFromBuffer: () => decoded({ isEmpty: () => true }),
      })
      await expect(prepare({ mime: 'image/png', bytes })).rejects.toThrow('invalid_image')
    }
    const prepare = createOfficeImageHandoff({
      createFromBuffer: () => {
        throw new Error('private decoder details')
      },
    })
    await expect(prepare({ mime: 'image/jpeg', bytes: new Uint8Array(100) })).rejects.toThrow(
      'invalid_image',
    )
  })

  it.each([0, LIMIT + 1])(
    'rejects impossible encoding size %i after bounded attempts',
    async (length) => {
      const toPNG = vi.fn(() => new Uint8Array(length))
      const toJPEG = vi.fn(() => new Uint8Array(length))
      const image = decoded({
        getSize: () => ({ width: 4096, height: 2048 }),
        resize: vi.fn(() => decoded({ toPNG, toJPEG })),
      })
      const prepare = createOfficeImageHandoff({ createFromBuffer: () => image })
      await expect(
        prepare({ mime: 'image/jpeg', bytes: new Uint8Array(LIMIT + 1) }),
      ).rejects.toThrow('image_limit')
      expect(image.resize).toHaveBeenCalledTimes(3)
      expect(toPNG).toHaveBeenCalledTimes(3)
      expect(toJPEG).toHaveBeenCalledTimes(3)
    },
  )
})
