import { readFile, stat, mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { build } from 'esbuild'

/** Offline: compile the same production parser, with no network or tool execution. */
export async function replayFile(path, index = 0) {
  if (!Number.isSafeInteger(index) || index < 0 || index > 3)
    throw new Error('invalid_recording_index')
  if ((await stat(path)).size > 512 * 1024) throw new Error('recording_too_large')
  const input = JSON.parse(await readFile(path, 'utf8'))
  const recording =
    input?.schema === 'wiswork-enhanced-diagnostics/v1' ? input.protocolRecordings?.[index] : input
  const info =
    input?.schema === 'wiswork-enhanced-diagnostics/v1'
      ? input.protocolRecordingInfo?.[index]
      : undefined
  const source =
    info &&
    typeof info.recordingId === 'string' &&
    /^recording_[A-Za-z0-9_-]{24}$/.test(info.recordingId) &&
    Number.isSafeInteger(info.recordedAt) &&
    info.recordedAt >= 0 &&
    ['completed', 'incomplete', 'protocol_rejected', 'interrupted', 'not_observed'].includes(
      info.originalOutcome,
    ) &&
    info.association === 'unattributed'
      ? {
          recordingId: info.recordingId,
          recordedAt: info.recordedAt,
          originalOutcome: info.originalOutcome,
          association: 'unattributed',
        }
      : undefined
  const directory = await mkdtemp(join(tmpdir(), 'wiswork-replay-'))
  try {
    const outfile = join(directory, 'replay.mjs')
    await build({
      stdin: {
        contents: 'export { replayProtocolRecording } from "./packages/codex-bridge/src/index.ts"',
        resolveDir: fileURLToPath(new URL('../', import.meta.url)),
      },
      outfile,
      bundle: true,
      platform: 'node',
      format: 'esm',
      logLevel: 'silent',
    })
    const { replayProtocolRecording } = await import(pathToFileURL(outfile).href)
    return { ...(await replayProtocolRecording(recording)), ...(source ? { source } : {}) }
  } finally {
    await rm(directory, { recursive: true, force: true })
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const args = process.argv.slice(2)
  const [path, flag, value] = args
  if (
    !path ||
    (flag !== undefined && flag !== '--index') ||
    (flag && value === undefined) ||
    args.length > 3
  ) {
    console.error(
      'Usage: node tools/diagnostics-replay.mjs <recording-or-export.json> [--index 0..3]',
    )
    process.exitCode = 2
  } else {
    try {
      console.log(
        JSON.stringify(await replayFile(path, value === undefined ? 0 : Number(value)), null, 2),
      )
    } catch {
      console.error('diagnostics_replay_rejected')
      process.exitCode = 1
    }
  }
}
