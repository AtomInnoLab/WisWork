import { describe, expect, it, vi } from 'vitest'
import { composeOfficeSkills } from '../src/agent/skill-registry.js'
import { exactObject, stringField } from '../src/agent/tool-schema.js'
import { createSandboxCommands } from '../src/skills/shared/commands.js'
import { createSharedBrowserSkill } from '../src/skills/shared/shared-skill.js'
import { parseSkillPackage, SkillRegistry } from '../src/skills/shared/skill-registry.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
import JSZip from 'jszip'

describe('exact schema helpers', () => {
  it('rejects unknown, missing, and oversized fields', () => {
    const parse = exactObject({ text: stringField({ minLength: 1, maxLength: 3 }) })
    expect(parse({ text: 'ok' })).toEqual({ text: 'ok' })
    expect(() => parse({ text: 'toolong' })).toThrow('invalid_tool_input')
    expect(() => parse({ text: 'ok', extra: true })).toThrow('invalid_tool_input')
    expect(() => parse({})).toThrow('invalid_tool_input')
  })
})

describe('in-memory VFS', () => {
  it('normalizes allowed roots and rejects traversal, device, NUL, and absolute escape paths', () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/a/../note.txt', 'safe')
    expect(vfs.readText('/home/user/note.txt')).toBe('safe')
    for (const path of ['/etc/passwd', '/home/user/../../etc/passwd', '/home/user/a\0b', 'C:\\x']) {
      expect(() => vfs.writeFile(path, 'no')).toThrow('vfs_path_denied')
    }
  })

  it('enforces file, total, count, name, and bounded-read limits', () => {
    const vfs = new InMemoryVfs({ maxFileBytes: 4, maxTotalBytes: 6, maxFiles: 2, maxNameBytes: 8 })
    vfs.writeFile('/home/user/a', '1234')
    expect(() => vfs.writeFile('/home/user/b', '12345')).toThrow('vfs_limit')
    vfs.writeFile('/home/user/b', '12')
    expect(() => vfs.writeFile('/home/user/c', '')).toThrow('vfs_limit')
    expect(() => vfs.writeFile('/home/user/toolongname', '')).toThrow('vfs_limit')
    expect(vfs.readText('/home/user/a', { offset: 1, maxBytes: 2 })).toBe('23')
  })

  it('returns valid UTF-8 when a byte bound lands inside a character', () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/utf8.txt', '你好')
    expect(vfs.readText('/home/user/utf8.txt', { maxBytes: 4 })).toBe('你')
    expect(() => vfs.readText('/home/user/utf8.txt', { offset: 1, maxBytes: 5 })).toThrow(
      'vfs_invalid_utf8_boundary',
    )
  })

  it('mounts skill files read-only', () => {
    const vfs = new InMemoryVfs()
    vfs.mountReadOnly('/home/skills/writer/SKILL.md', '# Writer')
    expect(vfs.readText('/home/skills/writer/SKILL.md')).toBe('# Writer')
    expect(() => vfs.writeFile('/home/skills/writer/SKILL.md', 'changed')).toThrow(
      'vfs_path_denied',
    )
    expect(() => vfs.mountReadOnly('/home/skills/writer/SKILL.md', 'changed')).toThrow(
      'vfs_path_denied',
    )
  })

  it('mounts read-only packages atomically when quota validation fails', () => {
    const vfs = new InMemoryVfs({ maxFiles: 1 })
    expect(() =>
      vfs.mountReadOnlyBatch([
        ['/home/skills/a/SKILL.md', 'one'],
        ['/home/skills/a/ref.md', 'two'],
      ]),
    ).toThrow('vfs_limit')
    expect(() => vfs.readText('/home/skills/a/SKILL.md')).toThrow('vfs_not_found')
  })
})

describe('skill packages', () => {
  it('parses strict metadata, mounts files, and supplies bounded prompt metadata', () => {
    const source = `---\nname: writer\ndescription: Helps edit prose\nversion: 1.2.0\n---\n# Rules\nBe concise.`
    const parsed = parseSkillPackage(source)
    expect(parsed.metadata).toEqual({
      name: 'writer',
      description: 'Helps edit prose',
      version: '1.2.0',
    })

    const vfs = new InMemoryVfs()
    const registry = new SkillRegistry(vfs)
    registry.install(source, { 'references/style.md': 'Use plain language.' })
    expect(vfs.readText('/home/skills/writer/references/style.md')).toBe('Use plain language.')
    expect(registry.prompt()).toContain('writer: Helps edit prose (/home/skills/writer/SKILL.md)')
  })

  it.each([
    'name: nope',
    '---\nname: ../escape\ndescription: bad\n---\nbody',
    '---\nname: ok\ndescription: fine\nunknown: no\n---\nbody',
  ])('rejects malformed metadata', (source) => {
    expect(() => parseSkillPackage(source)).toThrow('invalid_skill_package')
  })
})

describe('shared browser skill', () => {
  it('exports only implemented read and bash tools and bounded read output', async () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/a.txt', 'abcdef')
    const skill = createSharedBrowserSkill({ vfs, maxReadBytes: 3 })
    expect(skill.tools.map((tool) => tool.name)).toEqual(['read', 'bash'])
    await expect(
      skill.executeTool({ id: '1', name: 'read', input: { path: '/home/user/a.txt' } }),
    ).resolves.toMatchObject({ output: 'abc', mutated: false })
    await expect(
      skill.executeTool({
        id: '2',
        name: 'read',
        input: { path: '/home/user/a.txt', extra: true },
      }),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true })
  })

  it('returns uploaded images through model-visible display data', async () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/page.png', Uint8Array.from([137, 80, 78, 71]))
    const skill = createSharedBrowserSkill({ vfs })
    await expect(
      skill.executeTool({ id: 'image', name: 'read', input: { path: '/home/user/page.png' } }),
    ).resolves.toMatchObject({
      output: '{"path":"/home/user/page.png","mime":"image/png","bytes":4}',
      display: { kind: 'images', items: [{ url: 'data:image/png;base64,iVBORw==' }] },
    })
  })

  it('runs a real bounded DOCX-to-text conversion entirely inside the VFS', async () => {
    const zip = new JSZip()
    zip.file(
      'word/document.xml',
      '<w:document><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>',
    )
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/input.docx', await zip.generateAsync({ type: 'uint8array' }))
    const skill = createSharedBrowserSkill({ vfs })
    await expect(
      skill.executeTool({
        id: 'convert',
        name: 'bash',
        input: { command: 'docx-to-text /home/user/input.docx /home/user/output.txt' },
      }),
    ).resolves.toMatchObject({ output: '/home/user/output.txt', mutated: false })
    expect(vfs.readText('/home/user/output.txt')).toBe('Hello')
  })

  it('denies native/global/network syntax and supports cancellation and timeout', async () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/a.txt', 'hello')
    const commands = createSandboxCommands(vfs, {
      timeoutMs: 10,
      extraCommands: { hang: () => new Promise(() => undefined) },
    })
    await expect(commands.run('cat /home/user/a.txt')).resolves.toMatchObject({ output: 'hello' })
    for (const command of [
      'curl https://example.com',
      'node -e fetch(1)',
      'cat /etc/passwd',
      'echo $HOME',
      'ls; pwd',
    ]) {
      await expect(commands.run(command)).resolves.toMatchObject({ error: 'sandbox_denied' })
    }
    await expect(commands.run('hang')).resolves.toMatchObject({ error: 'command_timeout' })
    const controller = new AbortController()
    controller.abort()
    await expect(commands.run('pwd', controller.signal)).resolves.toMatchObject({
      error: 'cancelled',
    })
  })

  it('limits concurrent sandbox commands', async () => {
    let finish!: () => void
    const commands = createSandboxCommands(new InMemoryVfs(), {
      maxConcurrent: 1,
      extraCommands: {
        wait: () =>
          new Promise((resolve) => {
            finish = () => resolve({ output: 'done' })
          }),
      },
    })
    const first = commands.run('wait')
    await expect(commands.run('pwd')).resolves.toMatchObject({ error: 'command_limit' })
    finish()
    await expect(first).resolves.toMatchObject({ output: 'done' })
  })

  it('rejects inherited names and fail-closes after an uncooperative command timeout', async () => {
    const commands = createSandboxCommands(new InMemoryVfs(), {
      maxConcurrent: 1,
      timeoutMs: 5,
      extraCommands: { hang: () => new Promise(() => undefined) },
    })
    await expect(commands.run('toString')).resolves.toMatchObject({ error: 'sandbox_denied' })
    await expect(commands.run('constructor')).resolves.toMatchObject({ error: 'sandbox_denied' })
    await expect(commands.run('hang')).resolves.toMatchObject({ error: 'command_timeout' })
    await expect(commands.run('pwd')).resolves.toMatchObject({ error: 'command_unavailable' })
  })

  it('keeps the runtime fused when another concurrent command settles after an orphan timeout', async () => {
    let finish!: () => void
    const commands = createSandboxCommands(new InMemoryVfs(), {
      maxConcurrent: 2,
      timeoutMs: 30,
      extraCommands: {
        hang: () => new Promise(() => undefined),
        wait: () =>
          new Promise((resolve) => {
            finish = () => resolve({ output: 'done' })
          }),
      },
    })
    const orphan = commands.run('hang')
    await new Promise((resolve) => setTimeout(resolve, 5))
    const normal = commands.run('wait')
    await expect(orphan).resolves.toMatchObject({ error: 'command_timeout' })
    finish()
    await expect(normal).resolves.toMatchObject({ output: 'done' })
    await expect(commands.run('pwd')).resolves.toMatchObject({ error: 'command_unavailable' })
  })

  it('composes through agent-core and rejects duplicate tool names', () => {
    const shared = createSharedBrowserSkill({ vfs: new InMemoryVfs() })
    const host = { id: 'host', systemPrompt: '', tools: [], executeTool: vi.fn() }
    expect(composeOfficeSkills(host, shared).tools.map((tool) => tool.name)).toEqual([
      'read',
      'bash',
    ])
    expect(() => composeOfficeSkills(shared, shared)).toThrow('duplicate tool name: read')
  })
})
