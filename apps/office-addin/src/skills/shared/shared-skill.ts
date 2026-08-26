import type { AgentSkill, ToolExecution } from '@wiswork/agent-core'
import { exactObject, stringField } from '../../agent/tool-schema.js'
import { createConversionCommands, createSandboxCommands } from './commands.js'
import { ConversionWorkerRuntime } from './conversion-runtime.js'
import { InMemoryVfs, MAX_VFS_FILE_BYTES } from './vfs.js'
import type { SkillRegistry } from './skill-registry.js'

const MAX_PATH = 512
const MAX_COMMAND = 2_048
const readInput = exactObject({ path: stringField({ minLength: 1, maxLength: MAX_PATH }) })
const bashInput = exactObject({ command: stringField({ minLength: 1, maxLength: MAX_COMMAND }) })
const invalid = (summary: string): ToolExecution => ({
  output: 'invalid_tool_input',
  isError: true,
  mutated: false,
  summary,
})

export function createSharedBrowserSkill(options: {
  vfs: InMemoryVfs
  skills?: SkillRegistry
  maxReadBytes?: number
  conversionRuntime?: Pick<ConversionWorkerRuntime, 'run'>
  enableConversions?: boolean
}): AgentSkill {
  const maxReadBytes = options.maxReadBytes ?? 64 * 1024
  const conversionRuntime = options.conversionRuntime ?? new ConversionWorkerRuntime(options.vfs)
  const commands = createSandboxCommands(options.vfs, {
    timeoutMs: 20_000,
    extraCommands:
      options.enableConversions === false ? undefined : createConversionCommands(conversionRuntime),
  })
  return {
    id: 'office-shared-browser',
    systemPrompt:
      'read and bash operate only on the bounded browser VFS. bash is not a native shell. Never run pwd or ls and never probe the environment for orientation. Use VFS tools only when the user request needs a listed attachment, an installed skill names a path, or a conversion is required.',
    buildContext: () => {
      const files = options.vfs.list('/home/user')
      const attachments = files.length
        ? `Attachments:\n${files.map((path) => `- ${path}`).join('\n')}`
        : ''
      return [attachments, options.skills?.prompt() ?? ''].filter(Boolean).join('\n\n')
    },
    tools: [
      {
        name: 'read',
        description: 'Read bounded UTF-8 text from the browser VFS.',
        inputSchema: {
          type: 'object',
          properties: { path: { type: 'string', maxLength: MAX_PATH } },
          required: ['path'],
          additionalProperties: false,
        },
      },
      {
        name: 'bash',
        description: `Run a sandboxed VFS command. Implemented commands: ${commands.names.join(', ')}.`,
        inputSchema: {
          type: 'object',
          properties: { command: { type: 'string', maxLength: MAX_COMMAND } },
          required: ['command'],
          additionalProperties: false,
        },
      },
    ],
    async executeTool(call, signal) {
      if (call.inputError || call.truncated) return invalid(call.name)
      try {
        if (call.name === 'read') {
          const { path } = readInput(call.input)
          if (path.toLowerCase().endsWith('.gif')) throw new Error('image_mime_unsupported')
          const mime = imageMime(path)
          if (mime) {
            const bytes = options.vfs.readBytes(path, { maxBytes: MAX_VFS_FILE_BYTES + 1 })
            if (bytes.byteLength > MAX_VFS_FILE_BYTES) throw new Error('vfs_limit')
            return {
              output: JSON.stringify({ path, mime, bytes: bytes.byteLength }),
              modelContent: [{ type: 'image', image: { mime, base64: base64(bytes) } }],
              display: { kind: 'images', items: [{ url: `data:${mime};base64,${base64(bytes)}` }] },
              mutated: false,
              summary: `Read ${path}`,
            }
          }
          return {
            output: options.vfs.readText(path, { maxBytes: maxReadBytes }),
            mutated: false,
            summary: `Read ${path}`,
          }
        }
        if (call.name === 'bash') {
          const { command } = bashInput(call.input)
          const result = await commands.run(command, signal)
          const commandName = command.trim().split(/\s+/, 1)[0]
          return {
            output: result.error ?? result.output,
            isError: Boolean(result.error),
            mutated: false,
            summary: commandName === 'cat' ? 'Read attachment' : 'Converted attachment',
          }
        }
        return invalid(call.name)
      } catch (error) {
        const code = error instanceof Error ? error.message : 'command_failed'
        return { output: code, isError: true, mutated: false, summary: call.name }
      }
    },
  }
}

function imageMime(path: string): string | undefined {
  const extension = path.split('.').pop()?.toLowerCase()
  return {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    webp: 'image/webp',
  }[extension ?? '']
}

function base64(bytes: Uint8Array): string {
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 32 * 1024)
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 32 * 1024))
  return btoa(binary)
}
