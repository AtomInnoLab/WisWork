import type { ToolExecution, ToolExecutionSuspension } from './types'
import { mintToolExecutionSuspension } from './types'

/** Internal package boundary used by reviewed transports; not exported from the public root API. */
export function createInternalToolSuspensionIssuer(): Readonly<{
  suspend(result: Promise<ToolExecution>): ToolExecutionSuspension
}> {
  return Object.freeze({ suspend: mintToolExecutionSuspension })
}
