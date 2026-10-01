import { createHash } from 'node:crypto'
import { readFile, readdir, stat } from 'node:fs/promises'
import { basename, extname, join } from 'node:path'

const supported = new Set(['.pdf', '.csv', '.json', '.docx', '.txt', '.md', '.xlsx', '.pptx'])
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')

export async function loadBenchmarkBundle(repoRoot, caseId, variant) {
  if (!/^PPT-P0-(?:0[1-9]|1\d|20)$/.test(caseId)) throw new Error('invalid benchmark case')
  if (
    variant !== undefined &&
    (caseId !== 'PPT-P0-17' || !['science', 'legal', 'finance'].includes(variant))
  )
    throw new Error('invalid benchmark variant')
  const directory = join(repoRoot, 'docs/product/ppt-benchmark-materials', caseId)
  const prefix = variant ? `reference-${variant}` : 'reference'
  const plan = JSON.parse(await readFile(join(directory, `${prefix}-plan.json`), 'utf8'))
  const deck = JSON.parse(await readFile(join(directory, `${prefix}-deck.json`), 'utf8'))
  if (
    plan.projectId !== deck.id ||
    !Array.isArray(plan.sources) ||
    !Array.isArray(deck.slides) ||
    !deck.slides.length
  )
    throw new Error('invalid benchmark plan or deck')
  const expected = new Set(plan.sources.map((source) => source.snapshotAttachmentId))
  if (!expected.size || [...expected].some((id) => !/^[a-f0-9]{64}$/.test(id)))
    throw new Error('benchmark source snapshot missing')
  const matched = new Map()
  let visited = 0
  async function scan(dir, depth) {
    if (depth > 2) return
    for (const entry of await readdir(dir, { withFileTypes: true })) {
      if (++visited > 128) throw new Error('benchmark material directory too large')
      const path = join(dir, entry.name)
      if (entry.isDirectory()) await scan(path, depth + 1)
      else if (entry.isFile() && supported.has(extname(entry.name).toLowerCase())) {
        const size = (await stat(path)).size
        if (size < 1 || size > 10 * 1024 * 1024) continue
        const bytes = await readFile(path)
        const hash = sha(bytes)
        if (!expected.has(hash)) continue
        if (matched.has(hash)) throw new Error(`duplicate benchmark source: ${hash}`)
        matched.set(hash, { name: basename(path), bytes, sha256: hash })
      }
    }
  }
  await scan(directory, 0)
  if (matched.size !== expected.size)
    throw new Error(
      `benchmark source missing: ${[...expected].filter((id) => !matched.has(id)).join(',')}`,
    )
  return { plan, deck, sourceAttachments: [...expected].map((id) => matched.get(id)) }
}
