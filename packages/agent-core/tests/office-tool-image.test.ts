import { describe, expect, it } from 'vitest'
import {
  decodeOfficeScreenshotResult,
  encodeOfficeScreenshotResult,
  OFFICE_SCREENSHOT_PREVIEW_BYTES,
  OFFICE_SCREENSHOT_WIRE_BYTES,
} from '../src/office-tool-image'

const image = {
  mime: 'image/png',
  base64:
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4AWP4DwQACfsD/c8LaHIAAAAASUVORK5CYII=',
}
const wire = () =>
  encodeOfficeScreenshotResult(
    JSON.stringify({ fingerprint: 'original', visualAvailableToModel: true }),
    [{ type: 'image', image }],
  )

describe('bounded Office screenshot envelope', () => {
  it('keeps the old receiver fail-closed and restores an image only after unpacking', () => {
    const output = wire()
    expect(JSON.parse(output)).toMatchObject({
      visualAvailableToModel: false,
      metadata: { visualAvailableToModel: false },
    })
    expect(Buffer.byteLength(JSON.stringify({ output }))).toBeLessThan(256 * 1024)
    const decoded = decodeOfficeScreenshotResult(output)
    expect(decoded.modelContent).toEqual([{ type: 'image', image }])
    expect(JSON.parse(decoded.output)).toMatchObject({
      mime: 'image/png',
      bytes: 68,
      visualAvailableToModel: false,
    })
    expect(decoded.output).not.toContain(image.base64)
  })

  it.each([
    [
      'signature-only PNG',
      (value: any) => {
        value.image.base64 = Buffer.concat([
          Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
          Buffer.alloc(16),
        ]).toString('base64')
      },
    ],
    [
      'empty JPEG',
      (value: any) => {
        value.image = {
          mime: 'image/jpeg',
          base64: Buffer.from([255, 216, 255, 217]).toString('base64'),
        }
      },
    ],
    [
      'version',
      (value: any) => {
        value.schema = 'wiswork.office-screenshot/2'
      },
    ],
    [
      'extra field',
      (value: any) => {
        value.url = 'https://private.example'
      },
    ],
    [
      'multiple images',
      (value: any) => {
        value.image = [image, image]
      },
    ],
    [
      'MIME mismatch',
      (value: any) => {
        value.image.mime = 'image/jpeg'
      },
    ],
    [
      'unsupported MIME',
      (value: any) => {
        value.image.mime = 'image/svg+xml'
      },
    ],
    [
      'URL',
      (value: any) => {
        value.image.base64 = 'https://private.example'
      },
    ],
    [
      'noncanonical base64',
      (value: any) => {
        value.image.base64 += '\n'
      },
    ],
    [
      'invalid padding',
      (value: any) => {
        value.image.base64 = value.image.base64.slice(0, -2) + 'J='
      },
    ],
    [
      'empty',
      (value: any) => {
        value.image.base64 = ''
      },
    ],
    [
      'metadata-only',
      (value: any) => {
        delete value.image
      },
    ],
    [
      'availability claim',
      (value: any) => {
        value.visualAvailableToModel = true
      },
    ],
    [
      'nested availability claim',
      (value: any) => {
        value.metadata.visualAvailableToModel = true
      },
    ],
    [
      'image extra field',
      (value: any) => {
        value.image.url = 'https://private.example'
      },
    ],
    [
      'oversized image',
      (value: any) => {
        value.image.base64 = Buffer.alloc(OFFICE_SCREENSHOT_PREVIEW_BYTES + 1).toString('base64')
      },
    ],
  ])('rejects %s', (_label, mutate) => {
    const value = JSON.parse(wire())
    mutate(value)
    expect(() => decodeOfficeScreenshotResult(JSON.stringify(value))).toThrow(
      'office_screenshot_unavailable',
    )
  })

  it('bounds actual escaped Relay output, not just image bytes or unescaped metadata', () => {
    expect(() =>
      decodeOfficeScreenshotResult('x'.repeat(OFFICE_SCREENSHOT_WIRE_BYTES + 1)),
    ).toThrow('office_screenshot_unavailable')
    expect(() =>
      encodeOfficeScreenshotResult(JSON.stringify({ text: '"'.repeat(100_000) }), [
        { type: 'image', image },
      ]),
    ).toThrow('office_screenshot_unavailable')
    expect(() => encodeOfficeScreenshotResult('{}', undefined)).toThrow(
      'office_screenshot_unavailable',
    )
    expect(() =>
      encodeOfficeScreenshotResult('{}', [
        { type: 'image', image },
        { type: 'image', image },
      ]),
    ).toThrow('office_screenshot_unavailable')
  })
})
