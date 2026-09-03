import assert from 'node:assert/strict'
import test from 'node:test'
import { runMacosPackaging } from './run-macos-packaging.mjs'

const busy = `dmgbuild.core.DMGError: Unable to detach device cleanly: hdiutil: couldn't eject "disk4" - Resource busy`

test('force-detaches the exact transient dmg device and retries once', async () => {
  const calls = []
  const results = [
    { code: 1, output: busy },
    { code: 0, output: '' },
    { code: 0, output: 'packaged' },
  ]

  const code = await runMacosPackaging(['--mac', 'dmg', 'zip', '--x64'], {
    run: async (command, args) => {
      calls.push([command, args])
      return results.shift()
    },
  })

  assert.equal(code, 0)
  assert.deepEqual(calls, [
    ['npx', ['electron-builder', '--config', 'electron-builder.cjs', '--mac', 'dmg', 'zip', '--x64']],
    ['hdiutil', ['detach', '-force', '/dev/disk4']],
    ['npx', ['electron-builder', '--config', 'electron-builder.cjs', '--mac', 'dmg', 'zip', '--x64']],
  ])
})

test('does not retry unrelated packaging failures', async () => {
  let calls = 0
  const code = await runMacosPackaging(['--mac', 'dmg', '--arm64'], {
    run: async () => {
      calls += 1
      return { code: 1, output: 'codesign failed' }
    },
  })

  assert.equal(code, 1)
  assert.equal(calls, 1)
})

test('rejects malformed device identifiers without invoking hdiutil', async () => {
  const calls = []
  const code = await runMacosPackaging(['--mac', 'dmg'], {
    run: async (command) => {
      calls.push(command)
      return {
        code: 1,
        output:
          'Unable to detach device cleanly: hdiutil: couldn\'t eject "../../disk4" - Resource busy',
      }
    },
  })

  assert.equal(code, 1)
  assert.deepEqual(calls, ['npx'])
})

test('stops when forced detach fails', async () => {
  const calls = []
  const code = await runMacosPackaging(['--mac', 'dmg'], {
    run: async (command) => {
      calls.push(command)
      return command === 'npx' ? { code: 1, output: busy } : { code: 1, output: 'busy' }
    },
  })

  assert.equal(code, 1)
  assert.deepEqual(calls, ['npx', 'hdiutil'])
})
