import { EventEmitter } from 'node:events'
import { readFileSync } from 'node:fs'
import { request as httpsRequest } from 'node:https'
import type { IncomingMessage } from 'node:http'
import { Readable } from 'node:stream'
import ts from 'typescript'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { createOfficeImageHandoff } from '../src/main/office-image-handoff'
import { createOfficeLocalSearchProxy } from '../src/main/office-retrieval-proxy'

vi.mock('node:https', () => ({ request: vi.fn() }))

type NativeImage = Parameters<typeof createOfficeImageHandoff>[0]
type DecodedImage = ReturnType<NativeImage['createFromBuffer']>
const HANDOFF_LIMIT = 180 * 1024

// Exercise the production wiring without starting Electron/auth/IPC from index.ts.
function productionNormalizer(nativeImage: NativeImage) {
  const source = readFileSync(new URL('../src/main/index.ts', import.meta.url), 'utf8')
  const ast = ts.createSourceFile('index.ts', source, ts.ScriptTarget.Latest, true)
  let expression = ''
  const visit = (node: ts.Node) => {
    if (ts.isPropertyAssignment(node) && node.name.getText(ast) === 'normalizeImage')
      expression = node.initializer.getText(ast)
    ts.forEachChild(node, visit)
  }
  visit(ast)
  expect(expression).not.toBe('')
  const code = ts.transpile(`const normalize = ${expression};`, { target: ts.ScriptTarget.ES2022 })
  return new Function('nativeImage', 'createOfficeImageHandoff', `${code}; return normalize;`)(
    nativeImage,
    createOfficeImageHandoff,
  ) as ReturnType<typeof createOfficeImageHandoff>
}

function decoded(overrides: Partial<DecodedImage> = {}): DecodedImage {
  return {
    isEmpty: () => false,
    getSize: () => ({ width: 800, height: 600 }),
    resize: vi.fn(() => decoded()),
    toPNG: vi.fn(() => new Uint8Array(HANDOFF_LIMIT + 1)),
    toJPEG: vi.fn(() => new Uint8Array(HANDOFF_LIMIT + 1)),
    ...overrides,
  }
}

function serveImage(mime: string, bytes: Uint8Array, declared: number | null = bytes.byteLength) {
  vi.mocked(httpsRequest).mockImplementation(((
    _url: unknown,
    _options: unknown,
    respond: (response: IncomingMessage) => void,
  ) => {
    const response = Object.assign(Readable.from([bytes]), {
      statusCode: 200,
      headers: {
        'content-type': mime,
        ...(declared === null ? {} : { 'content-length': String(declared) }),
      },
    })
    const request = Object.assign(new EventEmitter(), {
      end: () => queueMicrotask(() => respond(response as unknown as IncomingMessage)),
      destroy: (error: Error) => {
        request.emit('error', error)
        request.emit('close')
      },
    })
    response.once('close', () => request.emit('close'))
    return request
  }) as typeof httpsRequest)
}

function retrieval(normalizeImage: ReturnType<typeof createOfficeImageHandoff>) {
  return createOfficeLocalSearchProxy({
    fetchWithAuth: vi.fn(),
    lookupAddresses: async () => [{ address: '93.184.216.34', family: 4 }],
    searchImages: async () => ({
      images: [
        {
          title: 'Cover',
          imageUrl: 'https://images.example/cover.jpg',
          sourceUrl: 'https://example.com/cover',
          source: 'example.com',
        },
      ],
      method: 'serpapi',
    }),
    normalizeImage,
  })
}

describe('Office downloaded image → production normalization → handoff', () => {
  beforeEach(() => vi.clearAllMocks())

  it('constructs a fresh retrieval proxy for each production relay client', () => {
    const source = readFileSync(new URL('../src/main/index.ts', import.meta.url), 'utf8')
    const ast = ts.createSourceFile('index.ts', source, ts.ScriptTarget.Latest, true)
    const bindings: string[] = []
    const visit = (node: ts.Node) => {
      if (ts.isCallExpression(node) && node.expression.getText(ast) === 'createOfficeRelayClient') {
        const options = node.arguments[0]
        if (options && ts.isObjectLiteralExpression(options)) {
          const binding = options.properties.find(
            (property) => property.name?.getText(ast) === 'retrievalProxy',
          )
          if (binding)
            bindings.push(
              ts.isPropertyAssignment(binding)
                ? binding.initializer.getText(ast)
                : binding.getText(ast),
            )
        }
      }
      ts.forEachChild(node, visit)
    }
    visit(ast)
    expect(bindings).toEqual(['createRetrievalProxy()'])
  })

  it.each([1_912_613, 4 * 1024 * 1024])(
    'shrinks a bounded %i-byte source before transport',
    async (sourceBytes) => {
      // NativeImage mock matching the measured 2800×2100, 1.9 MB JPEG expansion.
      const source = new Uint8Array(sourceBytes)
      const bounded = new Uint8Array(90_640)
      const original = decoded({
        getSize: () => ({ width: 2800, height: 2100 }),
        toPNG: vi.fn(() => new Uint8Array(19_860_520)),
        toJPEG: vi.fn(() => new Uint8Array(2_977_581)),
        resize: vi.fn(({ width, height }) =>
          decoded({
            getSize: () => ({ width, height }),
            toJPEG: () => (width <= 1024 ? bounded : new Uint8Array(HANDOFF_LIMIT + 1)),
          }),
        ),
      })
      const nativeImage: NativeImage = {
        createFromBuffer: (bytes) =>
          bytes.byteLength === bounded.byteLength
            ? decoded({ getSize: () => ({ width: 1024, height: 768 }) })
            : original,
      }
      serveImage('image/jpeg', source)
      const proxy = retrieval(productionNormalizer(nativeImage))
      await proxy('image-search.v1', { query: 'cover', max_results: 1 })
      const response = await proxy('image-fetch.v1', { url: 'https://images.example/cover.jpg' })
      const payload = JSON.parse(new TextDecoder().decode(response))
      const result = await createOfficeImageHandoff(nativeImage)({
        mime: payload.mime,
        bytes: Buffer.from(payload.data_base64, 'base64'),
      })
      expect(result.mime).toBe('image/jpeg')
      expect(new Uint8Array(result.bytes)).toEqual(bounded)
      expect(original.resize).toHaveBeenCalled()
      expect(Buffer.byteLength(payload.data_base64)).toBeLessThan(256 * 1024)
    },
  )

  it('normalizes even a fitting source instead of forwarding its original encoding', async () => {
    const canonical = new Uint8Array(80)
    const image = decoded({ toPNG: vi.fn(() => canonical) })
    const result = await productionNormalizer({ createFromBuffer: () => image })({
      mime: 'image/jpeg',
      bytes: new Uint8Array(100),
    })
    expect(result).toEqual({ mime: 'image/png', bytes: canonical })
    expect(image.toPNG).toHaveBeenCalled()
  })

  it.each(['image/webp', 'image/avif', 'image/gif'])(
    'advertises only PNG/JPEG and rejects unexpected %s before normalization',
    async (mime) => {
      serveImage(mime, new Uint8Array([1, 2, 3]))
      const normalize = vi.fn(async () => ({
        mime: 'image/png' as const,
        bytes: new Uint8Array(80),
      }))
      const proxy = retrieval(normalize)
      await proxy('image-search.v1', { query: 'cover', max_results: 1 })
      await expect(
        proxy('image-fetch.v1', { url: 'https://images.example/cover.jpg' }),
      ).rejects.toThrow('image_mime_unsupported')
      expect(normalize).not.toHaveBeenCalled()
      expect(httpsRequest).toHaveBeenCalledWith(
        expect.any(URL),
        expect.objectContaining({
          headers: expect.objectContaining({ Accept: 'image/png,image/jpeg' }),
        }),
        expect.any(Function),
      )
    },
  )

  it.each(['declared', 'chunked'])(
    'keeps the 10 MiB %s source limit before decoding or shrinking',
    async (mode) => {
      const bytes = new Uint8Array(10 * 1024 * 1024 + 1)
      serveImage('image/jpeg', bytes, mode === 'chunked' ? null : bytes.byteLength)
      const createFromBuffer = vi.fn(() => decoded())
      const proxy = retrieval(productionNormalizer({ createFromBuffer }))
      await proxy('image-search.v1', { query: 'cover', max_results: 1 })
      await expect(
        proxy('image-fetch.v1', { url: 'https://images.example/cover.jpg' }),
      ).rejects.toThrow('image_limit')
      expect(createFromBuffer).not.toHaveBeenCalled()
    },
  )
})
