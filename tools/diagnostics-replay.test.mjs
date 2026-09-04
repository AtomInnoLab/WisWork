import test from 'node:test'
import assert from 'node:assert/strict'
import { fileURLToPath } from 'node:url'
import { replayFile } from './diagnostics-replay.mjs'
import { mkdtemp, readFile, writeFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { spawnSync } from 'node:child_process'

test('offline CLI uses production state machine for sanitized fixtures', async () => {
  const fixture = fileURLToPath(
    new URL(
      '../packages/codex-bridge/tests/fixtures/protocol-redacted-max-tokens.json',
      import.meta.url,
    ),
  )
  const result = await replayFile(fixture)
  assert.equal(result.fidelity, 'structural-only')
  assert.ok(result.events.includes('response.incomplete'))
  assert.equal(result.error, undefined)
  await assert.rejects(replayFile(fixture, 4), /invalid_recording_index/)
})

test('spawned CLI accepts an explicit index in an exported report and rejects private schema fields', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'wiswork-replay-test-'))
  try {
    const path = join(directory, 'report.json')
    const fixture = JSON.parse(
      await readFile(
        new URL(
          '../packages/codex-bridge/tests/fixtures/protocol-redacted-max-tokens.json',
          import.meta.url,
        ),
        'utf8',
      ),
    )
    await writeFile(
      path,
      JSON.stringify({
        schema: 'wiswork-enhanced-diagnostics/v1',
        protocolRecordings: [fixture, fixture],
        protocolRecordingInfo: [
          null,
          {
            recordingId: `recording_${'a'.repeat(24)}`,
            recordedAt: 100,
            originalOutcome: 'incomplete',
            association: 'unattributed',
            secret: 'SECRET JWT',
          },
        ],
      }),
    )
    const cli = fileURLToPath(new URL('./diagnostics-replay.mjs', import.meta.url))
    const result = spawnSync(process.execPath, [cli, path, '--index', '1'], { encoding: 'utf8' })
    assert.equal(result.status, 0, result.stderr)
    assert.ok(JSON.parse(result.stdout).events.includes('response.incomplete'))
    assert.equal(JSON.parse(result.stdout).source.originalOutcome, 'incomplete')
    assert.equal(JSON.parse(result.stdout).source.recordingId, `recording_${'a'.repeat(24)}`)
    assert.equal(result.stdout.includes('SECRET'), false)
    await writeFile(path, JSON.stringify({ ...fixture, private: 'SECRET JWT' }))
    const rejected = spawnSync(process.execPath, [cli, path], { encoding: 'utf8' })
    assert.equal(rejected.status, 1)
    assert.equal(rejected.stderr.trim(), 'diagnostics_replay_rejected')
    assert.equal(rejected.stdout, '')
  } finally {
    await rm(directory, { recursive: true, force: true })
  }
})
