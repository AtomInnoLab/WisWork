import { describe, expect, it, vi } from 'vitest'
import {
  createOfficeDocumentClient,
  normalizeOfficeHost,
  type OfficeRuntime,
} from '../src/office-document.js'

function runtimeWith(
  result: { status: 'succeeded'; value?: string } | { status: 'failed'; message: string },
): OfficeRuntime {
  return {
    ready: vi.fn().mockResolvedValue({ host: 'Word' }),
    context: {
      document: {
        getSelectedDataAsync: vi.fn((_coercionType, callback) => callback(result)),
        setSelectedDataAsync: vi.fn((_value, _options, callback) => callback(result)),
      },
    },
  }
}

describe('normalizeOfficeHost', () => {
  it.each([
    ['Word', 'word'],
    ['Excel', 'excel'],
    ['PowerPoint', 'powerpoint'],
    ['Outlook', 'unknown'],
    [undefined, 'unknown'],
  ] as const)('maps %s to %s', (input, expected) => {
    expect(normalizeOfficeHost(input)).toBe(expected)
  })
})

describe('OfficeDocumentClient', () => {
  it('waits for Office and reports the active host', async () => {
    const runtime = runtimeWith({ status: 'succeeded' })
    const client = createOfficeDocumentClient(runtime)

    await expect(client.initialize()).resolves.toBe('word')
    expect(runtime.ready).toHaveBeenCalledOnce()
  })

  it('fails closed for an unsupported Office host', async () => {
    const runtime = runtimeWith({ status: 'succeeded', value: 'secret' })
    runtime.ready = vi.fn().mockResolvedValue({ host: 'Outlook' })
    const client = createOfficeDocumentClient(runtime)

    await expect(client.initialize()).resolves.toBe('unknown')
    await expect(client.readSelection()).rejects.toThrow('office_host_unsupported')
    await expect(client.replaceSelection('write')).rejects.toThrow('office_host_unsupported')
    expect(runtime.context.document.getSelectedDataAsync).not.toHaveBeenCalled()
    expect(runtime.context.document.setSelectedDataAsync).not.toHaveBeenCalled()
  })

  it('reads the selected text through the shared Office document API', async () => {
    const runtime = runtimeWith({ status: 'succeeded', value: 'Selected text' })

    await expect(createOfficeDocumentClient(runtime).readSelection()).resolves.toBe('Selected text')
    expect(runtime.context.document.getSelectedDataAsync).toHaveBeenCalledWith(
      'text',
      expect.any(Function),
    )
  })

  it('replaces the selection and rejects Office errors', async () => {
    const successRuntime = runtimeWith({ status: 'succeeded' })
    await expect(
      createOfficeDocumentClient(successRuntime).replaceSelection('WisWork'),
    ).resolves.toBeUndefined()
    expect(successRuntime.context.document.setSelectedDataAsync).toHaveBeenCalledWith(
      'WisWork',
      { coercionType: 'text' },
      expect.any(Function),
    )

    const failedRuntime = runtimeWith({ status: 'failed', message: 'Selection is locked' })
    await expect(createOfficeDocumentClient(failedRuntime).readSelection()).rejects.toThrow(
      'Selection is locked',
    )
  })

  it('appends text by preserving the freshly confirmed selection', async () => {
    const runtime = runtimeWith({ status: 'succeeded' })
    await createOfficeDocumentClient(runtime).appendText('before', ' after')
    expect(runtime.context.document.setSelectedDataAsync).toHaveBeenCalledWith(
      'before after',
      { coercionType: 'text' },
      expect.any(Function),
    )
  })
})
