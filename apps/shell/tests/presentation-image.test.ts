import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { inspectPresentationImage } from '../src/main/presentation-image'
const fixture = (name: string) => readFileSync(join(__dirname, 'fixtures/presentation-image', name))
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
  it('accepts bounded static GIF/WebP and rejects animation before decoding', () => {
    expect(inspectPresentationImage(fixture('static.gif'))).toEqual({
      mime: 'image/gif',
      width: 2,
      height: 3,
    })
    expect(inspectPresentationImage(fixture('static.webp'))).toEqual({
      mime: 'image/webp',
      width: 2,
      height: 3,
    })
    expect(inspectPresentationImage(fixture('extended.webp'))).toEqual({
      mime: 'image/webp',
      width: 2,
      height: 3,
    })
    expect(() => inspectPresentationImage(fixture('animated.gif'))).toThrow(
      'animated_image_unsupported',
    )
    expect(() => inspectPresentationImage(fixture('animated.webp'))).toThrow(
      'animated_image_unsupported',
    )
    expect(() => inspectPresentationImage(fixture('animated.png'))).toThrow(
      'animated_image_unsupported',
    )
    expect(inspectPresentationImage(fixture('animated.gif'), true)).toMatchObject({
      mime: 'image/gif',
      animated: true,
    })
    expect(inspectPresentationImage(fixture('animated.webp'), true)).toMatchObject({
      mime: 'image/webp',
      animated: true,
    })
    expect(inspectPresentationImage(fixture('animated.png'), true)).toMatchObject({
      mime: 'image/png',
      animated: true,
    })
    const apngBytes = fixture('animated.png')
    const chunks: Buffer[] = []
    for (let offset = 8; offset < apngBytes.length;) {
      const end = offset + 12 + apngBytes.readUInt32BE(offset)
      chunks.push(apngBytes.subarray(offset, end))
      offset = end
    }
    const posterOnly = Buffer.concat([
      apngBytes.subarray(0, 8),
      chunks[0]!,
      chunks[1]!,
      chunks[3]!,
      chunks[2]!,
      ...chunks.slice(4),
    ])
    expect(() => inspectPresentationImage(posterOnly, true)).toThrow('parse_failed')
    const animationChunk = Buffer.alloc(20)
    animationChunk.writeUInt32BE(8, 0)
    animationChunk.write('acTL', 4, 'ascii')
    animationChunk.writeUInt32BE(2, 8)
    let crc = 0xffffffff
    for (const byte of animationChunk.subarray(4, 16)) {
      crc ^= byte
      for (let bit = 0; bit < 8; bit++) crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1
    }
    animationChunk.writeUInt32BE((crc ^ 0xffffffff) >>> 0, 16)
    const apng = Buffer.concat([png.subarray(0, 33), animationChunk, png.subarray(33)])
    expect(() => inspectPresentationImage(apng)).toThrow('animated_image_unsupported')
    const huge = fixture('static.gif')
    huge.writeUInt16LE(8193, 6)
    expect(() => inspectPresentationImage(huge)).toThrow('parse_failed')
    expect(() => inspectPresentationImage(fixture('static.webp').subarray(0, -4))).toThrow(
      'parse_failed',
    )
  })
})
