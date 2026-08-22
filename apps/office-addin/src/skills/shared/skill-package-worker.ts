/// <reference lib="webworker" />
import { parseSkillArchive } from './skill-package.js'

declare const self: DedicatedWorkerGlobalScope

self.onmessage = async (event: MessageEvent<{ id: string; bytes: Uint8Array }>) => {
  const { id, bytes } = event.data ?? {}
  try {
    if (typeof id !== 'string' || !(bytes instanceof Uint8Array))
      throw new Error('invalid_skill_package')
    const pkg = await parseSkillArchive(bytes)
    self.postMessage(
      { id, ok: true, pkg },
      pkg.files.map((file) => file.bytes.buffer),
    )
  } catch (error) {
    const code = error instanceof Error ? error.message : ''
    const safe = ['invalid_skill_package', 'skill_package_limit'].includes(code)
      ? code
      : 'invalid_skill_package'
    self.postMessage({ id, ok: false, error: safe })
  }
}
