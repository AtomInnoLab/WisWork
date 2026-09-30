import { createHash } from 'node:crypto'
import { createReadStream } from 'node:fs'
import { readdir, readFile, realpath, stat } from 'node:fs/promises'
import { isAbsolute, posix, relative, resolve, sep } from 'node:path'
import { pathToFileURL } from 'node:url'
import { XMLParser, XMLValidator } from 'fast-xml-parser'
import JSZip from 'jszip'
import { PNG } from 'pngjs'

export const CASE_IDS = Array.from(
  { length: 20 },
  (_, index) => `PPT-P0-${String(index + 1).padStart(2, '0')}`,
)
const digest = (value) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const nonempty = (value) => typeof value === 'string' && value.trim().length > 0
const nonnegative = (value) => Number.isSafeInteger(value) && value >= 0
const RATIO_KEYS = [
  'first_two_pages_style_revisions',
  'first_round_visual_passes',
  'native_editable_objects',
  'critical_facts_sourced',
  'critical_claims_traced',
  'citations_accurate',
  'unsupported_factual_claims',
  'timely_numeric_claims',
  'reproducible_calculations',
  'successful_recoveries',
  'prepared_images',
  'user_interruptions',
  'taskpane_recoveries',
  'pairing_first_try',
  'manual_changes_preserved',
]
const COUNTER_KEYS = [
  'manual_correction_pages',
  'duplicate_writes',
  'screenshot_failures',
  'image_failures',
  'confidentiality_violations',
  'cross_document_writes',
]
const timestamp = (value) =>
  typeof value === 'string' &&
  /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(value) &&
  !Number.isNaN(Date.parse(value)) &&
  new Date(value).toISOString() === value

function validateMeasurements(value) {
  if (value === undefined) return
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('acceptance_measurements_invalid')
  const keys = ['started_at', 'first_real_page_at', 'finished_at', ...COUNTER_KEYS, 'ratios']
  if (Object.keys(value).some((key) => !keys.includes(key)))
    throw new Error('acceptance_measurements_invalid')
  for (const key of keys.slice(0, 3))
    if (value[key] !== undefined && !timestamp(value[key]))
      throw new Error('acceptance_measurements_invalid')
  for (const key of COUNTER_KEYS)
    if (value[key] !== undefined && (!nonnegative(value[key]) || value[key] > 1_000_000))
      throw new Error('acceptance_measurements_invalid')
  if (value.ratios !== undefined) {
    if (!value.ratios || typeof value.ratios !== 'object' || Array.isArray(value.ratios))
      throw new Error('acceptance_measurements_invalid')
    for (const [key, ratio] of Object.entries(value.ratios))
      if (
        !RATIO_KEYS.includes(key) ||
        !ratio ||
        typeof ratio !== 'object' ||
        Array.isArray(ratio) ||
        Object.keys(ratio).sort().join(',') !== 'denominator,numerator' ||
        !nonnegative(ratio.numerator) ||
        !nonnegative(ratio.denominator) ||
        ratio.numerator > ratio.denominator ||
        ratio.denominator > 1_000_000
      )
        throw new Error('acceptance_measurements_invalid')
  }
  const { started_at: start, first_real_page_at: first, finished_at: finish } = value
  if ((first || finish) && !start) throw new Error('acceptance_measurements_invalid')
  if (first && first < start) throw new Error('acceptance_measurements_invalid')
  if (finish && (finish < start || (first && finish < first)))
    throw new Error('acceptance_measurements_invalid')
}

function measurementSummary(latest) {
  const completed = latest.filter(Boolean)
  const ready = completed.length === CASE_IDS.length
  const counters = {}
  for (const key of COUNTER_KEYS) {
    const values = completed.map((record) => record.measurements?.[key])
    counters[key] = {
      observed: values.filter((value) => value !== undefined).length,
      total:
        ready && values.every((value) => value !== undefined)
          ? values.reduce((sum, value) => sum + value, 0)
          : 'not_measured',
    }
  }
  const latency = (end) => {
    const values = completed
      .map((record) => record.measurements)
      .filter((value) => value?.started_at && value?.[end])
      .map((value) => Date.parse(value[end]) - Date.parse(value.started_at))
      .sort((a, b) => a - b)
    return {
      observed: values.length,
      p95_ms:
        ready && values.length === CASE_IDS.length
          ? values[Math.ceil(0.95 * values.length) - 1]
          : 'not_measured',
    }
  }
  const ratios = {}
  for (const key of RATIO_KEYS) {
    const values = completed.map((record) => record.measurements?.ratios?.[key])
    const observed = values.filter(Boolean).length
    const complete = ready && observed === CASE_IDS.length
    const numerator = complete
      ? values.reduce((sum, value) => sum + value.numerator, 0)
      : 'not_measured'
    const denominator = complete
      ? values.reduce((sum, value) => sum + value.denominator, 0)
      : 'not_measured'
    ratios[key] = {
      observed,
      numerator,
      denominator,
      rate:
        complete && denominator > 0
          ? `${Math.round((numerator / denominator) * 10_000) / 100}%`
          : 'not_measured',
    }
  }
  return {
    durations: latency('finished_at'),
    first_page_latency: latency('first_real_page_at'),
    ratios,
    ...counters,
  }
}

export function summarizePresentationAcceptance(records) {
  if (!Array.isArray(records)) throw new Error('acceptance_records_invalid')
  const byCase = new Map(CASE_IDS.map((id) => [id, []]))
  const attempts = new Set()
  for (const record of records) {
    if (!record || typeof record !== 'object' || Array.isArray(record))
      throw new Error('acceptance_record_invalid')
    const entries = byCase.get(record.case_id)
    if (
      !entries ||
      !nonempty(record.attempt_id) ||
      !Number.isSafeInteger(record.attempt_no) ||
      record.attempt_no < 1
    )
      throw new Error('acceptance_record_invalid')
    const key = `${record.case_id}/${record.attempt_id}`
    if (attempts.has(key) || entries.some((entry) => entry.attempt_no === record.attempt_no))
      throw new Error('acceptance_duplicate_attempt')
    attempts.add(key)
    if (!['passed', 'failed', 'blocked'].includes(record.outcome))
      throw new Error('acceptance_outcome_invalid')
    validateMeasurements(record.measurements)
    if (record.outcome === 'passed') {
      if (
        record.material_status !== 'ready' ||
        !nonempty(record.material_manifest) ||
        !nonempty(record.commit_and_versions) ||
        !nonempty(record.identity) ||
        !nonempty(record.reviewer_and_date) ||
        !digest(record.artifacts?.pptx_sha256) ||
        record.artifacts?.powerpoint_reopened !== true ||
        record.artifacts?.editable_after_reopen !== true ||
        record.restart_required !== false ||
        record.p0_defects !== 0 ||
        record.measurements?.confidentiality_violations !== 0 ||
        record.measurements?.cross_document_writes !== 0
      )
        throw new Error('acceptance_pass_evidence_missing')
    }
    entries.push(record)
  }
  let executed = 0
  let passed = 0
  let firstAttemptPassed = 0
  let blocked = 0
  const latestAttempts = []
  const cases = CASE_IDS.map((id) => {
    const entries = byCase.get(id).sort((a, b) => a.attempt_no - b.attempt_no)
    if (entries.some((entry, index) => entry.attempt_no !== index + 1))
      throw new Error('acceptance_attempt_gap')
    const latest = entries.at(-1)
    if (entries[0]?.outcome === 'passed') firstAttemptPassed += 1
    latestAttempts.push(latest)
    if (latest) executed += 1
    if (latest?.outcome === 'passed') passed += 1
    if (latest?.outcome === 'blocked') blocked += 1
    return { id, attempts: entries.length, status: latest?.outcome ?? 'not_run' }
  })
  return {
    denominator: 20,
    attempted: executed,
    passed,
    firstAttemptPassed,
    failed: executed - passed - blocked,
    blocked,
    notRun: 20 - executed,
    attempts: records.length,
    completionRate: executed === 20 ? `${Math.round((passed / 20) * 100)}%` : 'not_measured',
    firstAttemptDeliveryRate:
      executed === 20 ? `${Math.round((firstAttemptPassed / 20) * 100)}%` : 'not_measured',
    rateThresholdMet: executed === 20 && passed >= 16,
    measurements: measurementSummary(latestAttempts),
    cases,
  }
}

export async function readPresentationAcceptance(directory) {
  const root = await realpath(directory)
  const names = (await readdir(directory)).filter((name) => name.endsWith('.json')).sort()
  const records = []
  for (const name of names) {
    const path = await boundedFile(root, name, 1024 * 1024)
    const parsed = JSON.parse(await readFile(path, 'utf8'))
    if (!Array.isArray(parsed)) throw new Error(`acceptance_file_invalid:${name}`)
    records.push(...parsed)
  }
  const report = summarizePresentationAcceptance(records)
  for (const record of records) {
    if (record.outcome !== 'passed') continue
    const materialPath = await verifyArtifact(
      root,
      record.material_manifest,
      record.material_manifest_sha256,
    )
    if (
      typeof record.artifacts?.pptx_file !== 'string' ||
      !record.artifacts.pptx_file.endsWith('.pptx')
    )
      throw new Error('acceptance_artifact_invalid:pptx_file')
    const pptxPath = await verifyArtifact(
      root,
      record.artifacts.pptx_file,
      record.artifacts.pptx_sha256,
    )
    await verifyPptx(pptxPath)
    const reopenPath = await verifyArtifact(
      root,
      record.artifacts.reopen_evidence_file,
      record.artifacts.reopen_evidence_sha256,
    )
    if (pptxPath === reopenPath) throw new Error('acceptance_artifact_invalid:reopen_evidence_file')
    const bundlePaths = new Set([materialPath, pptxPath, reopenPath])
    if (bundlePaths.size !== 3) throw new Error('acceptance_artifact_invalid:duplicate_file')
    if (
      record.artifacts.roundtrip_report_file !== undefined ||
      record.artifacts.roundtrip_report_sha256 !== undefined
    ) {
      if (
        typeof record.artifacts.roundtrip_report_file !== 'string' ||
        !record.artifacts.roundtrip_report_file.endsWith('.json')
      )
        throw new Error('acceptance_roundtrip_report_invalid')
      const path = await verifyArtifact(
        root,
        record.artifacts.roundtrip_report_file,
        record.artifacts.roundtrip_report_sha256,
        1024 * 1024,
      )
      if (bundlePaths.has(path)) throw new Error('acceptance_roundtrip_report_invalid')
      let comparison
      try {
        comparison = JSON.parse(await readFile(path, 'utf8'))
      } catch {
        throw new Error('acceptance_roundtrip_report_invalid')
      }
      if (
        comparison?.version !== 1 ||
        comparison.caseId !== record.case_id ||
        !nonempty(comparison.documentId) ||
        !nonempty(comparison.projectId) ||
        !nonempty(comparison.requestId) ||
        comparison.hostReopenVerified !== false ||
        !nonempty(comparison.requiredHumanEvidence) ||
        !['exact_part_bytes', 'changed_requires_review'].includes(comparison.packageComparison) ||
        !Array.isArray(comparison.changedParts) ||
        comparison.changedParts.length > 2000 ||
        comparison.changedParts.some((part) => typeof part !== 'string' || !part) ||
        (comparison.packageComparison === 'exact_part_bytes') !==
          (comparison.changedParts.length === 0) ||
        !digest(comparison.before?.bundleSha256) ||
        !digest(comparison.after?.bundleSha256) ||
        comparison.before.bundleSha256 === comparison.after.bundleSha256 ||
        !timestamp(comparison.before?.bundleCreatedAt) ||
        !timestamp(comparison.after?.bundleCreatedAt) ||
        comparison.after.bundleCreatedAt <= comparison.before.bundleCreatedAt ||
        !digest(comparison.before?.pptxSha256) ||
        comparison.after?.pptxSha256 !== record.artifacts.pptx_sha256
      )
        throw new Error('acceptance_roundtrip_report_invalid')
      bundlePaths.add(path)
    }
    for (const kind of ['claim_ledger', 'qa_report']) {
      const path = await verifyArtifact(
        root,
        record.artifacts[`${kind}_file`],
        record.artifacts[`${kind}_sha256`],
      )
      if (bundlePaths.has(path)) throw new Error(`acceptance_artifact_invalid:${kind}_file`)
      bundlePaths.add(path)
    }
    const screenshots = record.artifacts.page_screenshots
    if (
      !Array.isArray(screenshots) ||
      screenshots.length !== 8 ||
      screenshots.some((shot, index) => shot?.page_no !== index + 1)
    )
      throw new Error('acceptance_artifact_invalid:page_screenshots')
    for (const shot of screenshots) {
      if (typeof shot.file !== 'string' || !shot.file.endsWith('.png'))
        throw new Error('acceptance_artifact_invalid:page_screenshots')
      const path = await verifyArtifact(root, shot.file, shot.sha256, 20 * 1024 * 1024)
      if (bundlePaths.has(path)) throw new Error('acceptance_artifact_invalid:page_screenshots')
      bundlePaths.add(path)
      try {
        const data = await readFile(path)
        const pngSignature = Buffer.from('89504e470d0a1a0a', 'hex')
        if (
          data.length < 24 ||
          !data.subarray(0, 8).equals(pngSignature) ||
          !data.subarray(12, 16).equals(Buffer.from('IHDR'))
        )
          throw new Error('image_header_invalid')
        const width = data.readUInt32BE(16)
        const height = data.readUInt32BE(20)
        if (width < 1 || height < 1 || width * height > 16_000_000)
          throw new Error('image_size_invalid')
        PNG.sync.read(data)
      } catch {
        throw new Error('acceptance_artifact_invalid:page_screenshots')
      }
    }
  }
  return report
}

export async function verifyPptx(path) {
  try {
    const zip = await JSZip.loadAsync(await readFile(path))
    const expandedBytes = Object.values(zip.files).reduce(
      (total, entry) => total + (entry._data?.uncompressedSize ?? 0),
      0,
    )
    if (expandedBytes > 500 * 1024 * 1024) throw new Error('pptx_expanded_size_invalid')
    const xml = async (name) => {
      const part = zip.file(name)
      if (!part || part._data?.uncompressedSize > 10 * 1024 * 1024) throw new Error('part_missing')
      const content = await part.async('string')
      if (XMLValidator.validate(content) !== true) throw new Error('xml_invalid')
      return new XMLParser({ ignoreAttributes: false, parseTagValue: false }).parse(content)
    }
    const types = await xml('[Content_Types].xml')
    const overrides = [].concat(types.Types?.Override ?? [])
    if (!overrides.some((entry) => entry['@_PartName'] === '/ppt/presentation.xml'))
      throw new Error('presentation_type_missing')
    const presentation = await xml('ppt/presentation.xml')
    const slideIds = [].concat(presentation['p:presentation']?.['p:sldIdLst']?.['p:sldId'] ?? [])
    if (slideIds.length !== 8) throw new Error('acceptance_pptx_page_count')
    const relationships = await xml('ppt/_rels/presentation.xml.rels')
    const slideRels = new Map(
      []
        .concat(relationships.Relationships?.Relationship ?? [])
        .filter((entry) => entry['@_Type']?.endsWith('/slide'))
        .map((entry) => [entry['@_Id'], entry['@_Target']]),
    )
    const seenTargets = new Set()
    const seenContent = new Set()
    const seenSlideIds = new Set()
    const tagged = (value, key) => {
      if (Array.isArray(value)) return value.flatMap((item) => tagged(item, key))
      if (!value || typeof value !== 'object') return []
      return Object.entries(value).flatMap(([name, child]) => [
        ...(name === key ? [child] : []),
        ...tagged(child, key),
      ])
    }
    for (const slide of slideIds) {
      const target = slideRels.get(slide['@_r:id'])
      if (!target || !/^slides\/slide\d+\.xml$/.test(target))
        throw new Error('slide_relationship_invalid')
      if (seenTargets.has(target) || seenSlideIds.has(slide['@_id']))
        throw new Error('acceptance_pptx_duplicate_slide')
      seenTargets.add(target)
      seenSlideIds.add(slide['@_id'])
      const slideXml = await xml(`ppt/${target}`)
      if (!slideXml['p:sld']) throw new Error('slide_xml_invalid')
      const tree = slideXml['p:sld']['p:cSld']?.['p:spTree']
      if (
        !tree ||
        !['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp'].some((key) => tree[key])
      )
        throw new Error('acceptance_pptx_blank_slide')
      const pictures = tagged(tree, 'p:pic')
      const charts = tagged(tree, 'c:chart')
      if (pictures.length || charts.length) {
        const slideName = posix.basename(target)
        const relations = await xml(`ppt/slides/_rels/${slideName}.rels`)
        const byId = new Map(
          []
            .concat(relations.Relationships?.Relationship ?? [])
            .map((entry) => [entry['@_Id'], entry]),
        )
        for (const [kind, id] of [
          ...pictures.map((picture) => [
            'image',
            picture?.['p:blipFill']?.['a:blip']?.['@_r:embed'],
          ]),
          ...charts.map((chart) => ['chart', chart?.['@_r:id']]),
        ]) {
          const relation = byId.get(id)
          const assetTarget = relation?.['@_Target']
          const assetPath =
            typeof assetTarget === 'string' && assetTarget.startsWith('/ppt/')
              ? assetTarget.slice(1)
              : typeof assetTarget === 'string' && !assetTarget.startsWith('/')
                ? posix.normalize(posix.join('ppt/slides', assetTarget))
                : ''
          if (
            !id ||
            !relation ||
            relation['@_TargetMode'] !== undefined ||
            relation['@_Type'] !==
              `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}` ||
            !assetPath.startsWith(kind === 'image' ? 'ppt/media/' : 'ppt/charts/') ||
            !zip.file(assetPath) ||
            zip.file(assetPath)._data?.uncompressedSize === 0
          )
            throw new Error('acceptance_pptx_missing_asset')
          if (kind === 'chart') {
            const chartPart = await xml(assetPath)
            if (!chartPart['c:chartSpace']?.['c:chart'])
              throw new Error('acceptance_pptx_missing_asset')
          }
        }
      }
      const content = JSON.stringify(tree)
      if (!/"@_r:(?:embed|link|id)"/.test(content)) {
        if (seenContent.has(content)) throw new Error('acceptance_pptx_duplicate_slide')
        seenContent.add(content)
      }
    }
  } catch (error) {
    if (error?.message?.startsWith('acceptance_pptx_')) throw error
    throw new Error('acceptance_pptx_invalid', { cause: error })
  }
}

async function boundedFile(root, name, maxBytes) {
  if (typeof name !== 'string' || !name.trim() || isAbsolute(name))
    throw new Error('acceptance_artifact_invalid:path')
  const path = resolve(root, name)
  let actual, details
  try {
    actual = await realpath(path)
    details = await stat(actual)
  } catch {
    throw new Error('acceptance_artifact_invalid:path')
  }
  const fromRoot = relative(root, actual)
  if (
    !fromRoot ||
    fromRoot === '..' ||
    fromRoot.startsWith(`..${sep}`) ||
    isAbsolute(fromRoot) ||
    !details.isFile() ||
    details.size < 1 ||
    details.size > maxBytes
  )
    throw new Error('acceptance_artifact_invalid:path')
  return actual
}

async function verifyArtifact(root, name, expected, maxBytes = 100 * 1024 * 1024) {
  if (!digest(expected)) throw new Error('acceptance_artifact_invalid:digest')
  const path = await boundedFile(root, name, maxBytes)
  const hash = createHash('sha256')
  let size = 0
  for await (const chunk of createReadStream(path)) {
    size += chunk.length
    if (size > maxBytes) throw new Error('acceptance_artifact_invalid:size')
    hash.update(chunk)
  }
  if (hash.digest('hex') !== expected) throw new Error('acceptance_artifact_digest_mismatch')
  return path
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const directory = process.argv[2]
  if (!directory) {
    process.stderr.write('Usage: node tools/ppt-agent-acceptance.mjs <records-directory>\n')
    process.exitCode = 2
  } else {
    readPresentationAcceptance(resolve(directory))
      .then((report) => process.stdout.write(`${JSON.stringify(report, null, 2)}\n`))
      .catch((error) => {
        process.stderr.write(`${error instanceof Error ? error.message : 'acceptance_failed'}\n`)
        process.exitCode = 1
      })
  }
}
