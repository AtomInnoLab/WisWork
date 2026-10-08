import { createHash } from 'node:crypto'
import { readFile, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { CASE_IDS, verifyPptx } from './ppt-agent-acceptance.mjs'

const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex')
const validHash = (value) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const validId = (value) => typeof value === 'string' && value.length > 0 && value.length <= 4096
const validTime = (value) =>
  typeof value === 'string' &&
  /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(value) &&
  !Number.isNaN(Date.parse(value)) &&
  new Date(value).toISOString() === value

async function stage(directory) {
  const path = resolve(directory)
  const raw = await readFile(join(path, 'acceptance-draft.json'))
  if (raw.length > 64 * 1024) throw new Error('reopen_draft_invalid')
  let draft
  try {
    draft = JSON.parse(raw.toString('utf8'))
  } catch {
    throw new Error('reopen_draft_invalid')
  }
  if (
    draft?.status !== 'needs_human_review' ||
    !CASE_IDS.includes(draft.caseId) ||
    !validHash(draft.sourceBundleSha256) ||
    !validTime(draft.bundleCreatedAt) ||
    !validId(draft.documentId) ||
    !validId(draft.projectId) ||
    !validId(draft.requestId) ||
    !Number.isSafeInteger(draft.planRevision) ||
    draft.planRevision < 1 ||
    !validHash(draft.inputDigest) ||
    !validHash(draft.planDigest) ||
    draft.pptxFile !== 'presentation.pptx' ||
    !validHash(draft.pptxSha256) ||
    draft.structureGate !== 'passed'
  )
    throw new Error('reopen_draft_invalid')
  const pptxPath = join(path, 'presentation.pptx')
  const bytes = await readFile(pptxPath)
  if (bytes.length < 4 || bytes.length > 20 * 1024 * 1024 || sha256(bytes) !== draft.pptxSha256)
    throw new Error('reopen_pptx_invalid')
  await verifyPptx(pptxPath)
  return { draft, bytes }
}

async function entries(bytes) {
  const zip = await JSZip.loadAsync(bytes)
  const paths = Object.keys(zip.files)
    .filter((name) => !zip.files[name].dir)
    .sort()
  if (paths.length > 2000) throw new Error('reopen_compare_unavailable')
  const result = new Map()
  let total = 0
  for (const path of paths) {
    const file = zip.files[path]
    const size = file._data?.uncompressedSize
    if (!Number.isSafeInteger(size) || size < 0 || size > 100 * 1024 * 1024)
      throw new Error('reopen_compare_unavailable')
    total += size
    if (total > 100 * 1024 * 1024) throw new Error('reopen_compare_unavailable')
    result.set(path, sha256(await file.async('nodebuffer')))
  }
  return result
}

/** Compare two verified staged exports; a matching package cannot prove the host was reopened. */
export async function compareStagedReopen(beforeDirectory, afterDirectory) {
  const [before, after] = await Promise.all([stage(beforeDirectory), stage(afterDirectory)])
  const a = before.draft,
    b = after.draft
  if (
    a.caseId !== b.caseId ||
    a.documentId !== b.documentId ||
    a.projectId !== b.projectId ||
    a.requestId !== b.requestId ||
    a.planRevision !== b.planRevision ||
    a.inputDigest !== b.inputDigest ||
    a.planDigest !== b.planDigest ||
    a.sourceBundleSha256 === b.sourceBundleSha256 ||
    b.bundleCreatedAt <= a.bundleCreatedAt
  )
    throw new Error('reopen_identity_invalid')
  const [beforeEntries, afterEntries] = await Promise.all([
    entries(before.bytes),
    entries(after.bytes),
  ])
  const changedParts = [...new Set([...beforeEntries.keys(), ...afterEntries.keys()])]
    .filter((path) => beforeEntries.get(path) !== afterEntries.get(path))
    .sort()
  return {
    version: 1,
    caseId: a.caseId,
    documentId: a.documentId,
    projectId: a.projectId,
    requestId: a.requestId,
    planRevision: a.planRevision,
    inputDigest: a.inputDigest,
    planDigest: a.planDigest,
    before: {
      bundleSha256: a.sourceBundleSha256,
      bundleCreatedAt: a.bundleCreatedAt,
      pptxSha256: a.pptxSha256,
    },
    after: {
      bundleSha256: b.sourceBundleSha256,
      bundleCreatedAt: b.bundleCreatedAt,
      pptxSha256: b.pptxSha256,
    },
    packageComparison: changedParts.length ? 'changed_requires_review' : 'exact_part_bytes',
    changedParts,
    hostReopenVerified: false,
    requiredHumanEvidence: 'PowerPoint 保存、关闭、重开及重新编辑的现场录屏和审阅记录',
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const [beforeDirectory, afterDirectory, outputFile] = process.argv.slice(2)
  if (!beforeDirectory || !afterDirectory || !outputFile) {
    process.stderr.write(
      'Usage: node tools/ppt-agent-compare-reopen.mjs <before-stage-dir> <after-stage-dir> <new-report.json>\n',
    )
    process.exitCode = 2
  } else {
    compareStagedReopen(beforeDirectory, afterDirectory)
      .then(async (report) => {
        await writeFile(resolve(outputFile), `${JSON.stringify(report, null, 2)}\n`, { flag: 'wx' })
        process.stdout.write(`${JSON.stringify(report, null, 2)}\n`)
      })
      .catch((error) => {
        process.stderr.write(
          `${error instanceof Error ? error.message : 'reopen_compare_failed'}\n`,
        )
        process.exitCode = 1
      })
  }
}
