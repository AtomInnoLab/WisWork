import { InMemoryVfs } from './vfs.js'
import type { ConversionKind } from './conversion-engine.js'
import { ConversionWorkerRuntime } from './conversion-runtime.js'

export interface CommandResult {
  output: string
  error?: string
}
type Command = (args: string[], signal?: AbortSignal) => CommandResult | Promise<CommandResult>

const CONVERSIONS = Object.freeze([
  'pdf-to-text',
  'pdf-to-images',
  'docx-to-text',
  'xlsx-to-csv',
] satisfies ConversionKind[])

export function createConversionCommands(
  runtime: Pick<ConversionWorkerRuntime, 'run'>,
): Record<ConversionKind, Command> {
  return Object.fromEntries(
    CONVERSIONS.map((kind) => [
      kind,
      async (args: string[], signal?: AbortSignal): Promise<CommandResult> => {
        if (args.length !== 1) return { output: '', error: 'sandbox_denied' }
        try {
          const outputs = await runtime.run(kind, args[0], signal)
          return { output: outputs.join('\n') }
        } catch (error) {
          const code = error instanceof Error ? error.message : ''
          return {
            output: '',
            error: /^(?:cancelled|command_timeout|conversion_[a-z_]+|vfs_[a-z_]+)$/.test(code)
              ? code
              : 'command_failed',
          }
        }
      },
    ]),
  ) as Record<ConversionKind, Command>
}

export function createSandboxCommands(
  vfs: InMemoryVfs,
  options: {
    timeoutMs?: number
    maxOutputBytes?: number
    maxConcurrent?: number
    extraCommands?: Record<string, Command>
  } = {},
) {
  const timeoutMs = options.timeoutMs ?? 2_000
  const maxOutputBytes = options.maxOutputBytes ?? 64 * 1024
  const maxConcurrent = options.maxConcurrent ?? 2
  let active = 0
  let unavailable = false
  const denied = (): CommandResult => ({ output: '', error: 'sandbox_denied' })
  const builtins: Record<string, Command> = {
    cat: (args) =>
      args.length !== 1
        ? denied()
        : { output: vfs.readText(args[0], { maxBytes: maxOutputBytes }) },
  }
  const commands = { ...builtins, ...options.extraCommands }

  async function run(source: string, signal?: AbortSignal): Promise<CommandResult> {
    if (signal?.aborted) return { output: '', error: 'cancelled' }
    if (unavailable) return { output: '', error: 'command_unavailable' }
    if (active >= maxConcurrent) return { output: '', error: 'command_limit' }
    if (!source || /[;&|`$<>\n\r]/.test(source)) return denied()
    const parts = source.trim().split(/\s+/)
    if (!Object.hasOwn(commands, parts[0])) return denied()
    const command = commands[parts[0]]
    if (
      parts.some(
        (part) =>
          part.startsWith('/') &&
          !part.startsWith('/home/user/') &&
          !part.startsWith('/home/skills/'),
      )
    )
      return denied()
    active += 1
    const controller = new AbortController()
    let timer: ReturnType<typeof setTimeout> | undefined
    let abort: (() => void) | undefined
    const cancellation = new Promise<CommandResult>((resolve) => {
      abort = () => {
        unavailable = true
        controller.abort()
        resolve({ output: '', error: 'cancelled' })
      }
      signal?.addEventListener('abort', abort, { once: true })
    })
    const timeout = new Promise<CommandResult>((resolve) => {
      timer = setTimeout(() => {
        unavailable = true
        controller.abort()
        resolve({ output: '', error: 'command_timeout' })
      }, timeoutMs)
    })
    const execution = Promise.resolve()
      .then(() => command(parts.slice(1), controller.signal))
      .finally(() => {
        active -= 1
      })
    try {
      const result = await Promise.race([execution, timeout, cancellation])
      const output = new TextEncoder().encode(result.output).slice(0, maxOutputBytes)
      return { ...result, output: new TextDecoder().decode(output) }
    } catch (error) {
      const code = error instanceof Error ? error.message : ''
      return {
        output: '',
        error: ['vfs_path_denied', 'vfs_limit', 'vfs_not_found'].includes(code)
          ? code
          : 'command_failed',
      }
    } finally {
      if (timer) clearTimeout(timer)
      if (abort) signal?.removeEventListener('abort', abort)
    }
  }
  return { names: Object.freeze(Object.keys(commands)), run }
}
