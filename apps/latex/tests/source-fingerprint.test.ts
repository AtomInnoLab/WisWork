import { mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { fingerprintProjectDirectory } from '../src/main/source-fingerprint.js'

describe('project source fingerprint', () => {
  const roots: string[] = []
  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  it('changes for unopened dependencies and binary assets', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-fingerprint-'))
    roots.push(root)
    await mkdir(join(root, 'chapters'))
    await writeFile(join(root, 'main.tex'), '\\input{chapters/a}')
    await writeFile(join(root, 'chapters/a.tex'), 'before')
    await writeFile(join(root, 'figure.png'), Buffer.from([1, 2, 3]))
    const before = await fingerprintProjectDirectory(root)
    await writeFile(join(root, 'chapters/a.tex'), 'after')
    expect(await fingerprintProjectDirectory(root)).not.toBe(before)
    const afterText = await fingerprintProjectDirectory(root)
    await writeFile(join(root, 'figure.png'), Buffer.from([1, 2, 4]))
    expect(await fingerprintProjectDirectory(root)).not.toBe(afterText)
  })
})
