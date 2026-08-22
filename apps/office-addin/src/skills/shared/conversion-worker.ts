/// <reference lib="webworker" />
import { convertDocument, type ConversionInput } from './conversion-engine.js'

declare const self: DedicatedWorkerGlobalScope

self.onmessage = async (event: MessageEvent<ConversionInput & { id: string }>) => {
  const { id, kind, inputName, bytes } = event.data ?? {}
  try {
    if (typeof id !== 'string' || typeof inputName !== 'string' || !(bytes instanceof Uint8Array))
      throw new Error('conversion_invalid_document')
    const outputs = await convertDocument({ kind, inputName, bytes })
    self.postMessage(
      { id, ok: true, outputs },
      outputs.map((output) => output.bytes.buffer),
    )
  } catch (error) {
    const value = error instanceof Error ? error.message : ''
    const safe =
      /^conversion_(?:archive_unsafe|invalid_document|invalid_output|limit|unsupported)$/.test(
        value,
      )
        ? value
        : 'conversion_failed'
    self.postMessage({ id, ok: false, error: safe })
  }
}
