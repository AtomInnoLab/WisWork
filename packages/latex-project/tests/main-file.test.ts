import { mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { discoverMainFile, ProjectPathPolicy } from '../src/index.js'

describe('discoverMainFile', () => {
  let root: string

  beforeEach(async () => {
    root = await mkdtemp(join(tmpdir(), 'latex-main-'))
  })

  afterEach(async () => {
    await rm(root, { recursive: true, force: true })
  })

  async function discover(savedMainFile?: string) {
    return discoverMainFile(await ProjectPathPolicy.open(root), { savedMainFile })
  }

  it('prefers a valid saved main file over every discovered candidate', async () => {
    await writeFile(join(root, 'saved.tex'), '\\documentclass{article}')
    await writeFile(join(root, 'main.tex'), '\\documentclass{article}')
    await writeFile(join(root, 'Tectonic.toml'), 'inputs = ["tectonic.tex"]')
    await writeFile(join(root, 'tectonic.tex'), '\\documentclass{book}')
    await expect(discover('saved.tex')).resolves.toEqual({
      kind: 'found',
      path: 'saved.tex',
      source: 'saved',
    })
  })

  it('uses the single Tectonic.toml input before root main.tex', async () => {
    await writeFile(join(root, 'Tectonic.toml'), 'inputs = "src/paper.tex"')
    await mkdir(join(root, 'src'))
    await writeFile(join(root, 'src', 'paper.tex'), '\\documentclass{article}')
    await writeFile(join(root, 'main.tex'), '\\documentclass{article}')
    await expect(discover()).resolves.toEqual({
      kind: 'found',
      path: 'src/paper.tex',
      source: 'tectonic',
    })
  })

  it('parses arrays and comments without treating commented inputs as configuration', async () => {
    await writeFile(
      join(root, 'Tectonic.toml'),
      '# inputs = "ignored.tex"\ninputs = ["a.tex", "b.tex"] # real inputs',
    )
    await writeFile(join(root, 'ignored.tex'), '\\documentclass{article}')
    await writeFile(join(root, 'a.tex'), '\\documentclass{article}')
    await writeFile(join(root, 'b.tex'), '\\documentclass{book}')
    await expect(discover()).resolves.toEqual({
      kind: 'selection-required',
      candidates: ['a.tex', 'b.tex'],
    })
  })

  it('ignores a saved non-tex file', async () => {
    await writeFile(join(root, 'notes.txt'), 'saved')
    await writeFile(join(root, 'main.tex'), '\\documentclass{article}')
    await expect(discover('notes.txt')).resolves.toEqual({
      kind: 'found',
      path: 'main.tex',
      source: 'main',
    })
  })

  it('uses official output scalar inputs relative to src before legacy top-level inputs', async () => {
    await writeFile(
      join(root, 'Tectonic.toml'),
      [
        'inputs = "legacy.tex"',
        '[doc]',
        'name = "paper"',
        '',
        '[[output]]',
        'name = "default"',
        'type = "pdf"',
        'inputs = "paper.tex" # relative to src',
      ].join('\n'),
    )
    await mkdir(join(root, 'src'))
    await writeFile(join(root, 'src', 'paper.tex'), '\\\\documentclass{article}')
    await writeFile(join(root, 'legacy.tex'), '\\\\documentclass{book}')
    await expect(discover()).resolves.toEqual({
      kind: 'found',
      path: 'src/paper.tex',
      source: 'tectonic',
    })
  })

  it('supports official output arrays while ignoring inline input objects', async () => {
    await writeFile(
      join(root, 'Tectonic.toml'),
      [
        '[doc]',
        'name = "paper"',
        '[[output]]',
        'inputs = [{ inline = "ignored" }, "a.tex", "b.tex"]',
      ].join('\n'),
    )
    await mkdir(join(root, 'src'))
    await writeFile(join(root, 'src', 'a.tex'), '\\\\documentclass{article}')
    await writeFile(join(root, 'src', 'b.tex'), '\\\\documentclass{book}')
    await expect(discover()).resolves.toEqual({
      kind: 'selection-required',
      candidates: ['src/a.tex', 'src/b.tex'],
    })
  })

  it('requires selection across multiple official outputs', async () => {
    await writeFile(
      join(root, 'Tectonic.toml'),
      [
        '[doc]',
        'name = "paper"',
        '[[output]]',
        'name = "screen"',
        'inputs = "screen.tex"',
        '[[output]]',
        'name = "print"',
        'inputs = "print.tex"',
      ].join('\n'),
    )
    await mkdir(join(root, 'src'))
    await writeFile(join(root, 'src', 'screen.tex'), '\\\\documentclass{article}')
    await writeFile(join(root, 'src', 'print.tex'), '\\\\documentclass{article}')
    await expect(discover()).resolves.toEqual({
      kind: 'selection-required',
      candidates: ['src/print.tex', 'src/screen.tex'],
    })
  })

  it('falls back to root main.tex and then a unique documentclass candidate', async () => {
    await writeFile(join(root, 'main.tex'), 'plain')
    await expect(discover()).resolves.toEqual({ kind: 'found', path: 'main.tex', source: 'main' })
    await rm(join(root, 'main.tex'))
    await mkdir(join(root, 'src'))
    await writeFile(join(root, 'src', 'paper.tex'), '% comment\n\\documentclass{article}')
    await expect(discover()).resolves.toEqual({
      kind: 'found',
      path: 'src/paper.tex',
      source: 'documentclass',
    })
  })

  it('requires user selection when multiple documentclass files are candidates', async () => {
    await writeFile(join(root, 'a.tex'), '\\documentclass{article}')
    await writeFile(join(root, 'b.tex'), '\\documentclass{book}')
    await expect(discover()).resolves.toEqual({
      kind: 'selection-required',
      candidates: ['a.tex', 'b.tex'],
    })
  })
})
