import { readdir, readFile } from 'node:fs/promises'
import { resolve } from 'node:path'

const expected = process.argv[2]
if (expected !== 'present' && expected !== 'absent')
  throw new Error('usage: check-slides-e2e-artifact.mjs present|absent')

const root = resolve('apps/slides/out/renderer')
const files = await readdir(resolve(root, 'assets'))
const bundles = files.filter((file) => file.endsWith('.js'))
let occurrences = 0
for (const file of bundles) {
  const source = await readFile(resolve(root, 'assets', file), 'utf8')
  occurrences += source.split('__wisworkSlidesRunAcceptanceAgent').length - 1
  for (const forbidden of [
    '__wisworkSlidesRenderAcceptance',
    '__wisworkSlidesExecuteAcceptanceGolden',
  ]) {
    if (source.includes(forbidden)) throw new Error(`dead E2E symbol present: ${forbidden}`)
  }
}
if (expected === 'absent' && occurrences !== 0)
  throw new Error(`default artifact contains E2E entry (${occurrences})`)
if (expected === 'present' && occurrences !== 2)
  throw new Error(`E2E artifact must contain one assign/delete symbol pair (${occurrences})`)
