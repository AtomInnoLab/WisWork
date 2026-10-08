import type { Readable, Writable } from 'node:stream'

const DEFAULT_REQUEST_TIMEOUT_MS = 30_000
const DEFAULT_MAX_LINE_BYTES = 1_000_000
const DEFAULT_MAX_BUFFER_BYTES = 2_000_000
const DEFAULT_MAX_PENDING_REQUESTS = 64
const DEFAULT_MAX_QUEUED_WRITE_BYTES = 2_000_000

export interface JsonRpcNotification<TMethod extends string = string, TParams = unknown> {
  readonly method: TMethod
  readonly params: TParams
  readonly emittedAtMs?: number
}

export interface JsonRpcDiagnostic {
  readonly code: string
}

export interface JsonRpcClientOptions {
  readonly input: Readable
  readonly output: Writable
  readonly requestTimeoutMs?: number
  readonly maxLineBytes?: number
  readonly maxBufferBytes?: number
  readonly maxPendingRequests?: number
  readonly maxQueuedWriteBytes?: number
  readonly diagnostics?: (diagnostic: JsonRpcDiagnostic) => void
}

export interface JsonRpcRequestOptions {
  readonly timeoutMs?: number
  readonly signal?: AbortSignal
}

export class JsonRpcError extends Error {
  constructor(
    readonly code: string,
    readonly remoteCode?: number,
  ) {
    super(code)
    this.name = 'JsonRpcError'
  }
}

interface PendingRequest {
  readonly resolve: (result: unknown) => void
  readonly reject: (error: JsonRpcError) => void
  readonly timer: ReturnType<typeof setTimeout>
  readonly signal?: AbortSignal
  readonly onAbort?: () => void
}

function positiveInteger(value: number | undefined, fallback: number, code: string): number {
  const resolved = value ?? fallback
  if (!Number.isSafeInteger(resolved) || resolved <= 0) throw new TypeError(code)
  return resolved
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function hasExactKeys(value: Record<string, unknown>, keys: readonly string[]): boolean {
  const actual = Object.keys(value).sort()
  const expected = [...keys].sort()
  return actual.length === expected.length && actual.every((key, index) => key === expected[index])
}

function validateDataGraph(value: unknown): void {
  const pending = [value]
  const seen = new WeakSet<object>()
  let nodes = 0
  while (pending.length > 0) {
    const current = pending.pop()
    nodes += 1
    if (nodes > 100_000) throw new JsonRpcError('rpc_invalid_request')
    if (current === null || ['string', 'boolean'].includes(typeof current)) continue
    if (typeof current === 'number' && Number.isFinite(current)) continue
    if (typeof current !== 'object') throw new JsonRpcError('rpc_invalid_request')
    if (seen.has(current)) throw new JsonRpcError('rpc_invalid_request')
    seen.add(current)
    try {
      const array = Array.isArray(current)
      const prototype = Object.getPrototypeOf(current)
      if (
        prototype !== (array ? Array.prototype : Object.prototype) &&
        !(prototype === null && !array)
      )
        throw new Error()
      if (Object.getOwnPropertySymbols(current).length !== 0) throw new Error()
      const descriptors = Object.getOwnPropertyDescriptors(current)
      if (array) {
        const length = descriptors.length
        if (
          !length ||
          !('value' in length) ||
          !Number.isSafeInteger(length.value) ||
          (length.value as number) < 0 ||
          (length.value as number) > 100_000
        )
          throw new Error()
        const allowed = new Set([
          'length',
          ...Array.from({ length: length.value as number }, (_, index) => String(index)),
        ])
        if (
          Object.keys(descriptors).some((key) => !allowed.has(key)) ||
          [...allowed].some((key) => key !== 'length' && !(key in descriptors))
        )
          throw new Error()
      }
      for (const [key, descriptor] of Object.entries(descriptors)) {
        if (array && key === 'length') continue
        if (!('value' in descriptor)) throw new Error()
        pending.push(descriptor.value)
      }
    } catch (error) {
      if (error instanceof JsonRpcError) throw error
      throw new JsonRpcError('rpc_invalid_request')
    }
  }
}

export class JsonRpcClient<TNotification extends JsonRpcNotification = JsonRpcNotification> {
  readonly #input: Readable
  readonly #output: Writable
  readonly #requestTimeoutMs: number
  readonly #maxLineBytes: number
  readonly #maxBufferBytes: number
  readonly #maxPendingRequests: number
  readonly #maxQueuedWriteBytes: number
  readonly #diagnostics?: (diagnostic: JsonRpcDiagnostic) => void
  readonly #pending = new Map<number, PendingRequest>()
  readonly #subscribers = new Set<(notification: TNotification) => void>()
  #buffer = Buffer.alloc(0)
  #nextId = 1
  #queuedWriteBytes = 0
  #closedError: JsonRpcError | undefined
  #closePromise: Promise<void> | undefined

  constructor(options: JsonRpcClientOptions) {
    if (!options?.input || !options.output) throw new TypeError('invalid_rpc_options')
    this.#input = options.input
    this.#output = options.output
    this.#requestTimeoutMs = positiveInteger(
      options.requestTimeoutMs,
      DEFAULT_REQUEST_TIMEOUT_MS,
      'invalid_rpc_request_timeout',
    )
    this.#maxLineBytes = positiveInteger(
      options.maxLineBytes,
      DEFAULT_MAX_LINE_BYTES,
      'invalid_rpc_line_limit',
    )
    this.#maxBufferBytes = positiveInteger(
      options.maxBufferBytes,
      DEFAULT_MAX_BUFFER_BYTES,
      'invalid_rpc_buffer_limit',
    )
    this.#maxPendingRequests = positiveInteger(
      options.maxPendingRequests,
      DEFAULT_MAX_PENDING_REQUESTS,
      'invalid_rpc_pending_limit',
    )
    this.#maxQueuedWriteBytes = positiveInteger(
      options.maxQueuedWriteBytes,
      DEFAULT_MAX_QUEUED_WRITE_BYTES,
      'invalid_rpc_write_queue_limit',
    )
    this.#diagnostics = options.diagnostics
    this.#input.on('data', this.#onData)
    this.#input.once('end', this.#onInputEnd)
    this.#input.once('close', this.#onInputClose)
    this.#input.once('error', this.#onInputError)
    this.#output.once('error', this.#onOutputError)
  }

  request<TResult>(
    method: string,
    params: unknown,
    options: JsonRpcRequestOptions = {},
  ): Promise<TResult> {
    if (this.#closedError) return Promise.reject(this.#closedError)
    if (typeof method !== 'string' || method === '') {
      return Promise.reject(new JsonRpcError('rpc_invalid_request'))
    }
    if (options.signal?.aborted) {
      return Promise.reject(new JsonRpcError('rpc_request_aborted'))
    }
    if (this.#pending.size >= this.#maxPendingRequests) {
      return Promise.reject(new JsonRpcError('rpc_pending_limit_exceeded'))
    }
    const timeoutMs = positiveInteger(
      options.timeoutMs,
      this.#requestTimeoutMs,
      'invalid_rpc_request_timeout',
    )
    const id = this.#nextId
    let encoded: string
    try {
      encoded = this.#encode({ jsonrpc: '2.0', id, method, params })
    } catch (error) {
      return Promise.reject(error)
    }
    if (this.#queuedWriteBytes + Buffer.byteLength(encoded, 'utf8') > this.#maxQueuedWriteBytes) {
      return Promise.reject(new JsonRpcError('rpc_write_queue_limit_exceeded'))
    }
    this.#nextId += 1
    return new Promise<TResult>((resolve, reject) => {
      const onAbort = options.signal
        ? (): void => this.#settleRejected(id, new JsonRpcError('rpc_request_aborted'))
        : undefined
      const timer = setTimeout(
        () => this.#settleRejected(id, new JsonRpcError('rpc_request_timeout')),
        timeoutMs,
      )
      timer.unref()
      this.#pending.set(id, {
        resolve: (result) => resolve(result as TResult),
        reject,
        timer,
        signal: options.signal,
        onAbort,
      })
      options.signal?.addEventListener('abort', onAbort!, { once: true })
      if (!this.#writeEncoded(encoded)) {
        this.#settleRejected(id, this.#closedError ?? new JsonRpcError('rpc_transport_closed'))
      }
    })
  }

  notify(method: string, params: unknown): void {
    if (this.#closedError) throw this.#closedError
    if (typeof method !== 'string' || method === '') throw new JsonRpcError('rpc_invalid_request')
    const encoded = this.#encode({ jsonrpc: '2.0', method, params })
    if (this.#queuedWriteBytes + Buffer.byteLength(encoded, 'utf8') > this.#maxQueuedWriteBytes) {
      throw new JsonRpcError('rpc_write_queue_limit_exceeded')
    }
    if (!this.#writeEncoded(encoded)) {
      throw this.#closedError ?? new JsonRpcError('rpc_transport_closed')
    }
  }

  subscribe(listener: (notification: TNotification) => void): () => void {
    if (typeof listener !== 'function') throw new TypeError('invalid_rpc_subscriber')
    if (this.#closedError) return () => undefined
    this.#subscribers.add(listener)
    let subscribed = true
    return () => {
      if (!subscribed) return
      subscribed = false
      this.#subscribers.delete(listener)
    }
  }

  close(code = 'rpc_transport_closed'): Promise<void> {
    if (this.#closePromise) return this.#closePromise
    this.#terminate(new JsonRpcError(code), false)
    this.#closePromise = new Promise((resolve) => {
      if (this.#output.destroyed || this.#output.writableEnded) {
        resolve()
        return
      }
      this.#output.end(resolve)
    })
    return this.#closePromise
  }

  #encode(message: Record<string, unknown>): string {
    let encoded: string
    try {
      validateDataGraph(message)
      encoded = JSON.stringify(message)
    } catch {
      throw new JsonRpcError('rpc_invalid_request')
    }
    if (Buffer.byteLength(encoded, 'utf8') > this.#maxLineBytes) {
      throw new JsonRpcError('rpc_request_limit_exceeded')
    }
    return `${encoded}\n`
  }

  #writeEncoded(encoded: string): boolean {
    if (this.#closedError) return false
    try {
      const bytes = Buffer.byteLength(encoded, 'utf8')
      this.#queuedWriteBytes += bytes
      this.#output.write(encoded, () => {
        this.#queuedWriteBytes = Math.max(0, this.#queuedWriteBytes - bytes)
      })
      return true
    } catch {
      this.#terminate(new JsonRpcError('rpc_transport_closed'), true)
      return false
    }
  }

  readonly #onData = (chunk: Buffer | string): void => {
    if (this.#closedError) return
    const bytes = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk, 'utf8')
    if (this.#buffer.byteLength + bytes.byteLength > this.#maxBufferBytes) {
      this.#protocolFailure(
        this.#buffer.indexOf(0x0a) < 0 &&
          this.#buffer.byteLength + bytes.byteLength > this.#maxLineBytes
          ? 'rpc_line_limit_exceeded'
          : 'rpc_buffer_limit_exceeded',
      )
      return
    }
    this.#buffer = Buffer.concat([this.#buffer, bytes])
    const newline = this.#buffer.indexOf(0x0a)
    if (newline < 0 && this.#buffer.byteLength > this.#maxLineBytes) {
      this.#protocolFailure('rpc_line_limit_exceeded')
      return
    }
    if (this.#buffer.byteLength > this.#maxBufferBytes) {
      this.#protocolFailure('rpc_buffer_limit_exceeded')
      return
    }
    while (!this.#closedError) {
      const boundary = this.#buffer.indexOf(0x0a)
      if (boundary < 0) return
      const line = this.#buffer.subarray(0, boundary)
      this.#buffer = this.#buffer.subarray(boundary + 1)
      const content = line.at(-1) === 0x0d ? line.subarray(0, -1) : line
      if (content.byteLength > this.#maxLineBytes) {
        this.#protocolFailure('rpc_line_limit_exceeded')
        return
      }
      this.#processLine(content)
    }
  }

  #processLine(line: Buffer): void {
    let message: unknown
    try {
      message = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(line)) as unknown
    } catch {
      this.#protocolFailure('rpc_protocol_error')
      return
    }
    if (!isRecord(message) || ('jsonrpc' in message && message.jsonrpc !== '2.0')) {
      this.#protocolFailure('rpc_protocol_error')
      return
    }
    if ('id' in message) {
      this.#processResponse(message)
      return
    }
    if (
      !this.#isNotificationEnvelope(message) ||
      typeof message.method !== 'string' ||
      message.method === ''
    ) {
      this.#protocolFailure('rpc_protocol_error')
      return
    }
    const notification = {
      method: message.method,
      params: message.params,
      ...('emittedAtMs' in message ? { emittedAtMs: message.emittedAtMs as number } : {}),
    } as TNotification
    for (const subscriber of [...this.#subscribers]) {
      try {
        subscriber(notification)
      } catch {
        this.#emitDiagnostic('rpc_notification_handler_error')
      }
    }
  }

  #processResponse(message: Record<string, unknown>): void {
    if (!Number.isSafeInteger(message.id) || (message.id as number) <= 0) {
      this.#protocolFailure('rpc_protocol_error')
      return
    }
    const id = message.id as number
    const hasResult = 'result' in message
    const hasError = 'error' in message
    const expectedKeys = hasResult ? ['id', 'result'] : hasError ? ['id', 'error'] : []
    const keys = 'jsonrpc' in message ? ['jsonrpc', ...expectedKeys] : expectedKeys
    if (hasResult === hasError || !hasExactKeys(message, keys)) {
      this.#protocolFailure('rpc_protocol_error')
      return
    }
    const pending = this.#takePending(id)
    if (!pending) {
      this.#protocolFailure('rpc_protocol_error')
      return
    }
    if (hasResult) {
      pending.resolve(message.result)
      return
    }
    if (
      !isRecord(message.error) ||
      !Object.keys(message.error).every((key) => ['code', 'message', 'data'].includes(key)) ||
      !Number.isSafeInteger(message.error.code) ||
      typeof message.error.message !== 'string'
    ) {
      pending.reject(new JsonRpcError('rpc_protocol_error'))
      this.#protocolFailure('rpc_protocol_error')
      return
    }
    pending.reject(new JsonRpcError('rpc_remote_error', message.error.code as number))
  }

  #isNotificationEnvelope(message: Record<string, unknown>): boolean {
    const baseKeys = ['method', 'params']
    if ('jsonrpc' in message) baseKeys.push('jsonrpc')
    if ('emittedAtMs' in message) {
      if (!Number.isSafeInteger(message.emittedAtMs) || (message.emittedAtMs as number) < 0) {
        return false
      }
      baseKeys.push('emittedAtMs')
    }
    return hasExactKeys(message, baseKeys)
  }

  #takePending(id: number): PendingRequest | undefined {
    const pending = this.#pending.get(id)
    if (!pending) return undefined
    this.#pending.delete(id)
    clearTimeout(pending.timer)
    if (pending.onAbort) pending.signal?.removeEventListener('abort', pending.onAbort)
    return pending
  }

  #settleRejected(id: number, error: JsonRpcError): void {
    this.#takePending(id)?.reject(error)
  }

  #protocolFailure(code: string): void {
    this.#emitDiagnostic(code)
    this.#terminate(new JsonRpcError(code), true)
  }

  #emitDiagnostic(code: string): void {
    try {
      this.#diagnostics?.({ code })
    } catch {
      // Diagnostics are best-effort and contain codes only.
    }
  }

  readonly #onInputEnd = (): void => this.#terminate(new JsonRpcError('rpc_transport_closed'), true)
  readonly #onInputClose = (): void =>
    this.#terminate(new JsonRpcError('rpc_transport_closed'), true)
  readonly #onInputError = (): void =>
    this.#terminate(new JsonRpcError('rpc_transport_closed'), true)
  readonly #onOutputError = (): void =>
    this.#terminate(new JsonRpcError('rpc_transport_closed'), true)
  readonly #swallowStreamError = (): void => undefined

  #terminate(error: JsonRpcError, destroyOutput: boolean): void {
    if (this.#closedError) return
    this.#closedError = error
    this.#input.off('data', this.#onData)
    this.#input.off('end', this.#onInputEnd)
    this.#input.off('close', this.#onInputClose)
    this.#input.off('error', this.#onInputError)
    this.#output.off('error', this.#onOutputError)
    this.#input.on('error', this.#swallowStreamError)
    this.#output.on('error', this.#swallowStreamError)
    this.#subscribers.clear()
    for (const id of [...this.#pending.keys()]) this.#settleRejected(id, error)
    if (destroyOutput && !this.#output.destroyed) this.#output.destroy()
  }
}
