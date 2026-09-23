import { describe, expect, it } from 'vitest'
import { inspectPresentationImage } from '../src/main/presentation-image'
const png = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII=',
  'base64',
)
describe('image admission before native decoding', () => {
  it('recognizes PNG and bounded JPEG SOF dimensions', () => {
    expect(inspectPresentationImage(png)).toEqual({ mime: 'image/png', width: 1, height: 1 })
    const jpeg = Buffer.from([
      255, 216, 255, 224, 0, 2, 255, 192, 0, 11, 8, 0, 2, 0, 3, 1, 1, 17, 0, 255, 217,
    ])
    expect(inspectPresentationImage(jpeg)).toEqual({ mime: 'image/jpeg', width: 3, height: 2 })
  })
  it('rejects unsupported bytes and oversized headers before decoder allocation', () => {
    expect(() => inspectPresentationImage(Buffer.from('GIF89a'))).toThrow('parse_failed')
    const large = Buffer.from(png)
    large.writeUInt32BE(8193, 16)
    expect(() => inspectPresentationImage(large)).toThrow('parse_failed')
    large.writeUInt32BE(5000, 16)
    large.writeUInt32BE(5000, 20)
    expect(() => inspectPresentationImage(large)).toThrow('parse_failed')
    expect(() => inspectPresentationImage(Buffer.alloc(10 * 1024 * 1024 + 1))).toThrow(
      'parse_failed',
    )
    expect(() => inspectPresentationImage(png.subarray(0, 30))).toThrow('parse_failed')
  })
})
