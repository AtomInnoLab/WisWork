import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const folder = dirname(fileURLToPath(import.meta.url))
const scenario = JSON.parse(readFileSync(join(folder, 'scenario.json'), 'utf8'))
const expected = new Map(
  readFileSync(join(folder, 'SHA256SUMS'), 'utf8')
    .trim()
    .split('\n')
    .map((line) => {
      const match = /^([a-f0-9]{64})  ([a-z0-9-]+\.pdf)$/.exec(line)
      if (!match) throw new Error('invalid SHA256SUMS line')
      return [match[2], match[1]]
    }),
)
if (scenario.taskId !== 'PPT-P0-17' || scenario.documents?.length !== 3 || expected.size !== 3)
  throw new Error('invalid three-document scenario')
const markers = new Set()
const files = new Set()
for (const doc of scenario.documents) {
  if (
    typeof doc.source !== 'string' ||
    !expected.has(doc.source) ||
    files.has(doc.source) ||
    typeof doc.marker !== 'string' ||
    !/^(?:SCI|LAW|FIN)-P0-17$/.test(doc.marker) ||
    markers.has(doc.marker) ||
    !Array.isArray(doc.slideTitles) ||
    doc.slideTitles.length !== 8 ||
    doc.slideTitles.some((title) => typeof title !== 'string' || !title.trim())
  )
    throw new Error('invalid document plan')
  files.add(doc.source)
  markers.add(doc.marker)
  const path = join(folder, doc.source)
  const sha = createHash('sha256').update(readFileSync(path)).digest('hex')
  if (sha !== expected.get(doc.source)) throw new Error(`source hash mismatch: ${doc.source}`)
  const info = execFileSync('pdfinfo', [path], { encoding: 'utf8' })
  if (Number(/^Pages:\s+(\d+)$/m.exec(info)?.[1]) !== doc.pages)
    throw new Error(`page count mismatch: ${doc.source}`)
  const content = execFileSync('pdftotext', [path, '-'], {
    encoding: 'utf8',
    maxBuffer: 12 * 1024 * 1024,
  })
  const phrase = {
    science: 'computational reproducibility',
    legal: 'Guidelines 07/2020',
    finance: 'Apple Inc.',
  }[doc.key]
  if (!phrase || !content.toLowerCase().includes(phrase.toLowerCase()))
    throw new Error(`source identity mismatch: ${doc.source}`)
}
if (scenario.interleaving?.length !== 5 || files.size !== 3 || markers.size !== 3)
  throw new Error('invalid interleaving scenario')
process.stdout.write('PPT-P0-17: 3 source PDFs, 24 planned slides and scenario verified\n')
