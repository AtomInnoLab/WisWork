import { spawnSync } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkPlannedDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import {
  convertSinglePagePackageToPng,
  libreOfficeCommands,
  readBoundedRenderedPng,
} from '../src/main/presentation-page-render'

const sofficeAvailable = libreOfficeCommands().some(
  (command) => spawnSync(command, ['--version'], { timeout: 5_000 }).status === 0,
)

it('requires a working LibreOffice executable in cross-platform CI', () => {
  if (process.env.WISWORK_REQUIRE_LIBREOFFICE === '1') expect(sofficeAvailable).toBe(true)
})

it('checks standard Mac and Windows LibreOffice installations before PATH', () => {
  expect(libreOfficeCommands('darwin', {}, '/Users/test')).toEqual([
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    '/Users/test/Applications/LibreOffice.app/Contents/MacOS/soffice',
    'soffice',
  ])
  expect(
    libreOfficeCommands('win32', {
      ProgramFiles: 'C:\\Program Files',
      'ProgramFiles(x86)': 'C:\\Program Files (x86)',
    }),
  ).toEqual([
    'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    'soffice.exe',
  ])
  expect(libreOfficeCommands('linux', {})).toEqual(['soffice'])
})

it.skipIf(!sofficeAvailable)(
  'converts a real one-page PPTX to PNG for the fallback renderer',
  async () => {
    const deck = benchmarkPlannedDeck()
    const { bytes } = await compilePresentationDeck({ ...deck, slides: [deck.slides[0]!] })
    const png = Buffer.from(
      await convertSinglePagePackageToPng(bytes, new AbortController().signal),
    )
    expect(png.subarray(0, 8)).toEqual(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))
    expect(png.readUInt32BE(16)).toBeGreaterThan(0)
    expect(png.readUInt32BE(20)).toBeGreaterThan(0)
  },
)

it('rejects cancellation before starting a converter', async () => {
  const controller = new AbortController()
  controller.abort()
  await expect(
    convertSinglePagePackageToPng(new Uint8Array([1]), controller.signal),
  ).rejects.toThrow('aborted')
})

it.skipIf(!sofficeAvailable)('stops a running converter after cancellation', async () => {
  const deck = benchmarkPlannedDeck()
  const { bytes } = await compilePresentationDeck({ ...deck, slides: [deck.slides[0]!] })
  const controller = new AbortController()
  const rendering = convertSinglePagePackageToPng(bytes, controller.signal)
  setTimeout(() => controller.abort(), 50)
  await expect(rendering).rejects.toThrow('aborted')
})

it('rejects renderer output larger than the source cap before loading it', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'wiswork-render-test-'))
  try {
    const path = join(dir, 'page.png')
    await writeFile(path, Buffer.alloc(4 * 1024 * 1024 + 1))
    await expect(readBoundedRenderedPng(path)).rejects.toThrow('renderer_unavailable')
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
})

it('rejects compressed PNG dimensions beyond the decode cap', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'wiswork-render-test-'))
  try {
    const path = join(dir, 'page.png')
    const header = Buffer.alloc(33)
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]).copy(header)
    header.writeUInt32BE(13, 8)
    header.write('IHDR', 12)
    header.writeUInt32BE(100_000, 16)
    header.writeUInt32BE(1, 20)
    await writeFile(path, header)
    await expect(readBoundedRenderedPng(path)).rejects.toThrow('renderer_unavailable')
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
})
