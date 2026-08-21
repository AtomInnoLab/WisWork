import type { AgentSkill, ToolExecution } from '@wiswork/agent-core'
import { exactObject, stringField } from '../../agent/tool-schema.js'
import { createSandboxCommands } from './commands.js'
import { InMemoryVfs } from './vfs.js'

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
  maxReadBytes?: number
}): AgentSkill {
  const maxReadBytes = options.maxReadBytes ?? 64 * 1024
  const commands = createSandboxCommands(options.vfs)
  return {
    id: 'office-shared-browser',
    systemPrompt:
      'read and bash operate only on the bounded browser VFS. bash is not a native shell.',
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
          return {
            output: options.vfs.readText(path, { maxBytes: maxReadBytes }),
            mutated: false,
            summary: `Read ${path}`,
          }
        }
        if (call.name === 'bash') {
          const { command } = bashInput(call.input)
          const result = await commands.run(command, signal)
          return {
            output: result.error ?? result.output,
            isError: Boolean(result.error),
            mutated: false,
            summary: command,
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
