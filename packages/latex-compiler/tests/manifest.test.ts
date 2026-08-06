import { readFile } from 'node:fs/promises'
import { describe, expect, it } from 'vitest'
import { parseTectonicManifest } from '../src/manifest.js'

const hash = 'a'.repeat(64)

function validManifest() {
  return {
    schemaVersion: 1,
    tectonic: {
      version: '0.16.9',
      license: {
        spdx: 'MIT',
        sourceUrl: 'https://github.com/tectonic-typesetting/tectonic/blob/66b6654/LICENSE',
      },
      assets: [
        {
          id: 'tectonic-0.16.9-darwin-arm64',
          platform: 'darwin-arm64',
          url: 'https://github.com/tectonic-typesetting/tectonic/releases/download/tectonic%400.16.9/tectonic.tar.gz',
          bytes: 20_590_132,
          sha256: hash,
          archive: { format: 'tar.gz', executable: 'tectonic' },
        },
      ],
    },
    bundle: {
      id: 'tectonic-default-bundle-v33',
      url: 'https://relay.fullyjustified.net/default_bundle_v33.tar',
      bytes: 2_881_562_112,
      sha256: 'b'.repeat(64),
      license: {
        spdx: 'LicenseRef-Tectonic-Bundle',
        sourceUrl: 'https://tug.org/texlive/copying.html',
      },
    },
  }
}

describe('parseTectonicManifest', () => {
  it('accepts a pinned manifest and returns immutable data', () => {
    const parsed = parseTectonicManifest(validManifest())
    expect(parsed.tectonic.version).toBe('0.16.9')
    expect(parsed.tectonic.assets[0]?.platform).toBe('darwin-arm64')
    expect(Object.isFrozen(parsed.bundle)).toBe(true)
  })

  it.each([
    [
      'http URL',
      (value: ReturnType<typeof validManifest>) =>
        (value.bundle.url = 'http://relay.fullyjustified.net/bundle.tar'),
    ],
    [
      'unapproved host',
      (value: ReturnType<typeof validManifest>) =>
        (value.bundle.url = 'https://example.com/bundle.tar'),
    ],
    [
      'latest URL',
      (value: ReturnType<typeof validManifest>) =>
        (value.tectonic.assets[0]!.url =
          'https://github.com/tectonic-typesetting/tectonic/releases/latest/download/tectonic.tar.gz'),
    ],
    ['bad digest', (value: ReturnType<typeof validManifest>) => (value.bundle.sha256 = 'abc')],
    ['zero bytes', (value: ReturnType<typeof validManifest>) => (value.bundle.bytes = 0)],
    [
      'missing license',
      (value: ReturnType<typeof validManifest>) =>
        delete (value.bundle as Partial<typeof value.bundle>).license,
    ],
  ])('rejects %s', (_label, mutate) => {
    const value = validManifest()
    mutate(value)
    expect(() => parseTectonicManifest(value)).toThrow()
  })

  it('rejects duplicate platform assets and duplicate IDs', () => {
    const value = validManifest()
    value.tectonic.assets.push({ ...value.tectonic.assets[0]!, id: value.bundle.id })
    expect(() => parseTectonicManifest(value)).toThrow(/duplicate/i)
  })

  it('rejects unknown fields', () => {
    const value = validManifest() as ReturnType<typeof validManifest> & { surprise?: boolean }
    value.surprise = true
    expect(() => parseTectonicManifest(value)).toThrow(/unknown fields/i)
  })

  it('accepts the checked-in manifest with exact reviewed asset identities', async () => {
    const manifestUrl = new URL('../../../tools/tectonic/manifest.json', import.meta.url)
    const parsed = parseTectonicManifest(JSON.parse(await readFile(manifestUrl, 'utf8')))
    expect(parsed.tectonic.version).toBe('0.16.9')
    expect(parsed.tectonic.assets.map((asset) => asset.platform)).toEqual([
      'darwin-arm64',
      'darwin-x64',
    ])
    expect(parsed.bundle).toMatchObject({
      bytes: 2_881_562_112,
      sha256: '425685e124746c15ba9bb8e0596bdaad98fce886afa347fbcf9ec0e9acd7fe79',
    })
  })
})
