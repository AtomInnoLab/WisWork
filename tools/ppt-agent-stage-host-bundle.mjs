import { createHash } from 'node:crypto'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { basename, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { CASE_IDS, verifyPptx } from './ppt-agent-acceptance.mjs'

const REQUIRED = [
  'presentation.pptx',
  'evidence.json',
  'evidence.md',
  'claims.json',
  'sources.json',
  'quality.json',
  'checkpoints.json',
  'README.md',
]
const SCREENSHOTS = Array.from({ length: 8 }, (_, index) => `page-${index + 1}.png`)
const OPTIONAL = ['presentation.pdf', 'research.json', 'research.md', ...SCREENSHOTS]
const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex')
const validId = (value) => typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const validHash = (value) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)

/** Stage verified host bytes for a human PowerPoint acceptance run; never records a pass. */
export async function stagePresentationHostBundle(bundlePath, outputDirectory, caseId) {
  if (!CASE_IDS.includes(caseId)) throw new Error('stage_case_invalid')
  const bundle = await readFile(bundlePath)
  if (bundle.length < 4 || bundle.length > 20 * 1024 * 1024) throw new Error('stage_bundle_invalid')
  const zip = await JSZip.loadAsync(bundle)
  const entries = Object.values(zip.files)
  if (
    entries.some(
      (entry) =>
        entry.dir ||
        ![...REQUIRED, ...OPTIONAL, 'manifest.json'].includes(entry.name) ||
        !Number.isSafeInteger(entry._data?.uncompressedSize) ||
        entry._data.uncompressedSize < 1 ||
        entry._data.uncompressedSize > 20 * 1024 * 1024,
    ) ||
    entries.reduce((sum, entry) => sum + entry._data.uncompressedSize, 0) > 32 * 1024 * 1024
  )
    throw new Error('stage_bundle_invalid')
  const manifestEntry = zip.file('manifest.json')
  if (!manifestEntry) throw new Error('stage_bundle_invalid')
  let manifest
  try {
    manifest = JSON.parse(await manifestEntry.async('string'))
  } catch {
    throw new Error('stage_manifest_invalid')
  }
  if (
    manifest?.version !== 1 ||
    manifest.scope !== 'current_office_document' ||
    !validId(manifest.projectId) ||
    !validId(manifest.requestId) ||
    typeof manifest.documentId !== 'string' ||
    !manifest.documentId.trim() ||
    !Array.isArray(manifest.files) ||
    manifest.checks?.completion !== 'not_verified' ||
    manifest.checks?.roundTrip !== 'not_run' ||
    manifest.checks?.sourceAuthority !== 'not_verified' ||
    manifest.checks?.timeliness !== 'not_verified'
  )
    throw new Error('stage_manifest_invalid')
  const declared = new Map()
  for (const file of manifest.files) {
    if (
      ![...REQUIRED, ...OPTIONAL].includes(file?.name) ||
      declared.has(file.name) ||
      !Number.isSafeInteger(file.sizeBytes) ||
      file.sizeBytes < 1 ||
      file.sizeBytes > (SCREENSHOTS.includes(file.name) ? 64 * 1024 : 20 * 1024 * 1024) ||
      !validHash(file.sha256)
    )
      throw new Error('stage_manifest_invalid')
    declared.set(file.name, file)
  }
  if (
    REQUIRED.some((name) => !declared.has(name)) ||
    declared.size + 1 !== entries.length ||
    (manifest.checks.pdf === 'included') !== declared.has('presentation.pdf') ||
    declared.has('research.json') !== declared.has('research.md') ||
    (manifest.checks.pageScreenshots === 'captured_unreviewed') !==
      SCREENSHOTS.every((name) => declared.has(name)) ||
    (manifest.checks.pageScreenshots !== 'captured_unreviewed' &&
      SCREENSHOTS.some((name) => declared.has(name)))
  )
    throw new Error('stage_manifest_invalid')
  const files = new Map()
  for (const [name, file] of declared) {
    const part = zip.file(name)
    if (!part || part._data?.uncompressedSize !== file.sizeBytes)
      throw new Error('stage_manifest_invalid')
    const bytes = await part.async('nodebuffer')
    if (bytes.length !== file.sizeBytes || sha256(bytes) !== file.sha256)
      throw new Error('stage_digest_mismatch')
    files.set(name, bytes)
  }
  let evidence, quality, claims
  try {
    evidence = JSON.parse(files.get('evidence.json').toString('utf8'))
    quality = JSON.parse(files.get('quality.json').toString('utf8'))
    claims = JSON.parse(files.get('claims.json').toString('utf8'))
  } catch {
    throw new Error('stage_evidence_invalid')
  }
  if (
    evidence?.documentId !== manifest.documentId ||
    evidence?.projectId !== manifest.projectId ||
    evidence?.requestId !== manifest.requestId ||
    quality?.scope !== 'historical_records_only' ||
    quality?.needsRecapture !== true ||
    !Array.isArray(claims) ||
    (manifest.checks.pageScreenshots === 'captured_unreviewed' &&
      (!Array.isArray(quality.currentHostScreenshots) ||
        quality.currentHostScreenshots.length !== 8 ||
        quality.currentHostScreenshots.some(
          (shot, index) =>
            shot?.pageNo !== index + 1 ||
            typeof shot.hostSlideId !== 'string' ||
            !shot.hostSlideId ||
            shot.sha256 !== declared.get(SCREENSHOTS[index]).sha256,
        )))
  )
    throw new Error('stage_evidence_invalid')
  const output = resolve(outputDirectory)
  try {
    await mkdir(output)
  } catch {
    throw new Error('stage_output_exists')
  }
  for (const [name, bytes] of files) await writeFile(resolve(output, name), bytes)
  const draft = {
    version: 1,
    status: 'needs_human_review',
    caseId,
    sourceBundle: basename(bundlePath),
    sourceBundleSha256: sha256(bundle),
    projectId: manifest.projectId,
    requestId: manifest.requestId,
    documentId: manifest.documentId,
    pptxFile: 'presentation.pptx',
    pptxSha256: declared.get('presentation.pptx').sha256,
    historicalClaimsFile: 'claims.json',
    historicalQualityFile: 'quality.json',
    historicalEvidenceFile: 'evidence.json',
    pageScreenshots:
      manifest.checks.pageScreenshots === 'captured_unreviewed'
        ? quality.currentHostScreenshots.map((shot, index) => ({
            pageNo: index + 1,
            file: SCREENSHOTS[index],
            sha256: shot.sha256,
            hostSlideId: shot.hostSlideId,
            status: 'unreviewed',
          }))
        : [],
    missingEvidence: [
      'PowerPoint 保存、关闭、重开后的文件和可编辑对象核验',
      manifest.checks.pageScreenshots === 'captured_unreviewed'
        ? '已采集宿主截图的逐页人工视觉复核与当前文稿一致性检查'
        : '当前文稿逐页宿主截图与视觉复核',
      '当前版本五层 QA 与来源/专业判断复核',
      '操作员、宿主版本、时间、故障和人工修正记录',
    ],
  }
  let structureGate = 'passed'
  try {
    await verifyPptx(resolve(output, 'presentation.pptx'))
  } catch (error) {
    structureGate = error instanceof Error ? error.message : 'acceptance_pptx_invalid'
  }
  draft.structureGate = structureGate
  await writeFile(resolve(output, 'acceptance-draft.json'), `${JSON.stringify(draft, null, 2)}\n`)
  return draft
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const [bundlePath, outputDirectory, caseId] = process.argv.slice(2)
  if (!bundlePath || !outputDirectory || !caseId) {
    process.stderr.write(
      'Usage: node tools/ppt-agent-stage-host-bundle.mjs <bundle.zip> <new-output-directory> <PPT-P0-NN>\n',
    )
    process.exitCode = 2
  } else {
    stagePresentationHostBundle(bundlePath, outputDirectory, caseId)
      .then((draft) => process.stdout.write(`${JSON.stringify(draft, null, 2)}\n`))
      .catch((error) => {
        process.stderr.write(`${error instanceof Error ? error.message : 'stage_failed'}\n`)
        process.exitCode = 1
      })
  }
}
