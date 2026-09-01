import assert from 'node:assert/strict'
import test from 'node:test'
import { orchestrateSlidesAcceptanceE2E } from './run-slides-acceptance-e2e.mjs'

async function scenario(codes) {
  const calls = []
  const reports = []
  const code = await orchestrateSlidesAcceptanceE2E(
    async ({ name }) => {
      calls.push(name)
      return codes[name] ?? 0
    },
    (message) => reports.push(message),
  )
  return { code, calls, reports }
}

test('runs primary and always restores the default artifact', async () => {
  const result = await scenario({})
  assert.equal(result.code, 0)
  assert.deepEqual(result.calls, [
    'e2e-build',
    'e2e-artifact',
    'playwright',
    'default-build',
    'default-artifact',
  ])
})

test('preserves primary failure while reporting cleanup failure', async () => {
  const result = await scenario({ playwright: 7, 'default-build': 9 })
  assert.equal(result.code, 7)
  assert.deepEqual(result.calls.slice(-2), ['default-build', 'default-artifact'])
  assert.match(result.reports[0], /cleanup failed \(9\)/)
})

test('returns cleanup failure when the primary flow succeeds', async () => {
  const result = await scenario({ 'default-artifact': 4 })
  assert.equal(result.code, 4)
})

test('stops primary steps after failure but still runs all cleanup steps', async () => {
  const result = await scenario({ 'e2e-build': 3 })
  assert.equal(result.code, 3)
  assert.deepEqual(result.calls, ['e2e-build', 'default-build', 'default-artifact'])
})
