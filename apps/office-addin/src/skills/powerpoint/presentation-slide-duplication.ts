import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import {
  selectionFingerprint,
  type StructuredProposalController,
} from '../../agent/proposal-controller.js'
import {
  validatePowerPointPageScreenshot,
  type PowerPointAdapter,
} from './browser-powerpoint-adapter.js'
import type { PresentationPageReplacementAdapter } from './browser-presentation-page-replacement-adapter.js'
import {
  validatePresentationPageReplacement,
  type PresentationPageReplacement,
} from './presentation-page-replacement-record.js'
import {
  validatePresentationExistingBatch,
  type PresentationExistingBatch,
  type PresentationSlideDuplicationBatch,
} from './presentation-existing-batch.js'
import {
  describePagePackageBackup,
  readChartPackageBackup,
  saveChartPackageBackup,
  cleanupUncommittedChartPackageBackup,
} from './presentation-chart-backup.js'
import {
  loadBoundedZip,
  presentationPackageDigest,
  MAX_PPTX_XML_BYTES,
} from './powerpoint-package.js'
import { XMLParser, XMLValidator } from 'fast-xml-parser'
interface Options {
  documentId(): Promise<string>
  available(): boolean
  adapter: PowerPointAdapter
  pageAdapter: PresentationPageReplacementAdapter
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  proposals: StructuredProposalController
  readExistingBatch(id: string): PresentationExistingBatch | undefined
  writeExistingBatch(
    next: PresentationExistingBatch,
    expected: PresentationExistingBatch | undefined,
  ): Promise<void>
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const many = (value: any): any[] =>
  value === undefined ? [] : Array.isArray(value) ? value : [value]
const xmlParser = new XMLParser({
  ignoreAttributes: false,
  parseAttributeValue: false,
  parseTagValue: false,
  processEntities: false,
})
async function sourceIdentity(base64: string, signal?: AbortSignal) {
  const zip = await loadBoundedZip(base64, signal)
  const presentation = await zip.file('ppt/presentation.xml')?.async('string'),
    relations = await zip.file('ppt/_rels/presentation.xml.rels')?.async('string')
  for (const xml of [presentation, relations])
    if (
      !xml ||
      new TextEncoder().encode(xml).byteLength > MAX_PPTX_XML_BYTES ||
      /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml) ||
      XMLValidator.validate(xml) !== true
    )
      throw Error('presentation_page_source_invalid')
  const ids = many(xmlParser.parse(presentation!)?.['p:presentation']?.['p:sldIdLst']?.['p:sldId'])
  const paths = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  const rels = many(xmlParser.parse(relations!)?.Relationships?.Relationship)
  if (
    ids.length !== 1 ||
    paths.length !== 1 ||
    typeof ids[0]?.['@_id'] !== 'string' ||
    !/^[1-9][0-9]{0,9}$/.test(ids[0]['@_id']) ||
    Number(ids[0]['@_id']) < 256 ||
    Number(ids[0]['@_id']) > 4294967295
  )
    throw Error('presentation_page_source_invalid')
  const matching = rels.filter(
    (rel) =>
      rel['@_Id'] === ids[0]['@_r:id'] &&
      rel['@_Type'] ===
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide' &&
      rel['@_Target'] === `slides/${paths[0].split('/').at(-1)}` &&
      rel['@_TargetMode'] === undefined,
  )
  if (matching.length !== 1) throw Error('presentation_page_source_invalid')
  return `${ids[0]['@_id']}#`
}
function temporary(
  r: PresentationSlideDuplicationBatch,
  state?: PresentationPageReplacement['state'],
): PresentationPageReplacement {
  const record: PresentationPageReplacement = {
    version: 1,
    changeId: r.changeId,
    documentId: r.documentId,
    projectId: 'native-duplication',
    parentRequestId: 'original-source',
    requestId: 'native-copy',
    pageId: 'native-source-page',
    backupId: r.backups[0].backupId,
    parentArtifactDigest: r.baselineDigest,
    backupDigest: r.backups[0].sha256,
    originalPackageDigest: r.baselineDigest,
    replacementPackageDigest: r.baselineDigest,
    sourceSlideId: r.sourceSlideId,
    oldSlideId: r.hostSlideId,
    beforeSlideIds: [...r.beforeSlideIds],
    state: state ?? (r.insertedSlideId ? 'staged' : 'pending'),
    ...(r.insertedSlideId ? { newSlideId: r.insertedSlideId } : {}),
  }
  if (!validatePresentationPageReplacement(record))
    throw Error('presentation_slide_duplication_record_invalid')
  return record
}
const names = [
  'inspect_slide_duplication',
  'reconcile_slide_duplication',
  'undo_slide_duplication',
  'capture_slide_duplication_page',
  'record_slide_duplication_page_review',
] as const
const tools: AgentToolDef[] = names.map((name) => ({
  name,
  description:
    name === 'inspect_slide_duplication'
      ? 'Inspect actual source and owned copied page packages and exact order without writing or accepting QA.'
      : name === 'reconcile_slide_duplication'
        ? 'Propose confirmed recognition of a uniquely adjacent unchanged copy after a durable insertion intent, or close a proven unchanged original deck when no copy was created. Never replays insertion.'
        : name === 'undo_slide_duplication'
          ? 'Propose removing only the unchanged owned copied page while retaining the source. An uncertain removal is inspected before confirmed receipt closure without deleting again.'
          : name === 'capture_slide_duplication_page'
            ? 'Capture the applied owned copy only after fresh exact package and order checks. Does not accept whole-deck QA.'
            : 'Save a historical assessment of this session copied-page capture only after fresh screenshot and package checks.',
  inputSchema: {
    type: 'object',
    properties: {
      change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
      ...(name === 'record_slide_duplication_page_review'
        ? {
            screenshot_digest: { type: 'string', pattern: '^[a-f0-9]{64}$' },
            status: { type: 'string', enum: ['pass', 'fail'] },
            notes: { type: 'string', maxLength: 2000 },
          }
        : {}),
    },
    required:
      name === 'record_slide_duplication_page_review'
        ? ['change_id', 'screenshot_digest', 'status', 'notes']
        : ['change_id'],
    additionalProperties: false,
  },
}))
export function createPresentationSlideDuplicationSkill(options: Options) {
  let epoch = 0,
    visualEpoch = 0
  const captures = new Map<
    string,
    { digest: string; capturedAt: string; visualEpoch: number; insertedSlideId: string }
  >()
  const guard = async (
    documentId: string,
    signal?: AbortSignal,
    token = epoch,
    writing = false,
  ) => {
    const active = () => {
      if (signal?.aborted || epoch !== token) throw Error('cancelled')
      if (writing && !options.available())
        throw Error('presentation_existing_persistence_unavailable')
    }
    active()
    if ((await options.documentId()) !== documentId) throw Error('presentation_document_changed')
    active()
  }
  const saved = (id: string) => {
    const r = options.readExistingBatch(id)
    if (!r || r.version !== 4 || !validatePresentationExistingBatch(r))
      throw Error('presentation_existing_batch_missing')
    return structuredClone(r)
  }
  const current = async (
    r: PresentationSlideDuplicationBatch,
    signal?: AbortSignal,
    token = epoch,
    writing = false,
  ) => {
    await guard(r.documentId, signal, token, writing)
    if (!same(options.readExistingBatch(r.changeId), r))
      throw Error('presentation_existing_batch_stale')
  }
  const observe = async (
    r: PresentationSlideDuplicationBatch,
    signal?: AbortSignal,
    token = epoch,
    writing = false,
  ) => {
    await current(r, signal, token, writing)
    if (!r.insertedSlideId) {
      const proof = structuredClone(
        await options.pageAdapter.reconcilePending(temporary(r), signal),
      )
      await current(r, signal, token, writing)
      return { status: proof.status, insertedSlideId: proof.newSlideId }
    }
    const proof = structuredClone(await options.pageAdapter.inspect(temporary(r), signal))
    await current(r, signal, token, writing)
    return {
      status:
        proof.status === 'staged'
          ? ('inserted' as const)
          : proof.status === 'baseline'
            ? ('baseline' as const)
            : ('conflict' as const),
      insertedSlideId: r.insertedSlideId,
    }
  }
  const backup = async (
    r: PresentationSlideDuplicationBatch,
    signal?: AbortSignal,
    token = epoch,
    writing = false,
  ) => {
    const base64 = await readChartPackageBackup(
      {
        request: options.request,
        documentId: r.documentId,
        hostSlideId: r.hostSlideId,
        slideIds: r.beforeSlideIds,
        backup: r.backups[0],
        expectedPackageDigest: r.baselineDigest,
      },
      signal,
    )
    await current(r, signal, token, writing)
    const identity = await sourceIdentity(base64, signal)
    await current(r, signal, token, writing)
    if (identity !== r.sourceSlideId) throw Error('presentation_page_source_invalid')
    return base64
  }
  const screenshotDigest = async (base64: string) =>
    Array.from(
      new Uint8Array(
        await crypto.subtle.digest(
          'SHA-256',
          Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)),
        ),
      ),
      (b) => b.toString(16).padStart(2, '0'),
    ).join('')
  return {
    id: 'office-presentation-slide-duplication',
    systemPrompt:
      'Duplicate only through a confirmed durable original-page savepoint and journal. An uncertain insertion is never replayed; confirm recognition of one uniquely adjacent exact-package copy. Undo removes only an unchanged owned copy and retains the source. Package proof does not accept visual or professional QA.',
    tools,
    clear() {
      epoch++
      visualEpoch++
      captures.clear()
    },
    beginMutation() {
      visualEpoch++
      captures.clear()
    },
    endMutation() {
      visualEpoch++
      captures.clear()
    },
    async propose(
      slideIndex: number,
      explanation?: string,
      signal?: AbortSignal,
      toolName: 'duplicate_slide' | 'execute_office_js' = 'duplicate_slide',
      validateProduction?: (signal?: AbortSignal) => Promise<boolean>,
    ) {
      const token = epoch
      if (!options.available()) throw Error('presentation_existing_persistence_unavailable')
      if (
        !Number.isSafeInteger(slideIndex) ||
        slideIndex < 0 ||
        slideIndex > 100000 ||
        !['duplicate_slide', 'execute_office_js'].includes(toolName) ||
        (explanation !== undefined && (typeof explanation !== 'string' || explanation.length > 300))
      )
        throw Error('invalid_tool_input')
      if (!options.adapter.exportPresentationPagePackage) throw Error('office_api_unsupported')
      const documentId = await options.documentId()
      await guard(documentId, signal, token, true)
      const snap = structuredClone(await options.adapter.snapshotSlide(slideIndex, signal))
      await guard(documentId, signal, token, true)
      const page = structuredClone(
        await options.adapter.exportPresentationPagePackage(snap.slideId, signal),
      )
      await guard(documentId, signal, token, true)
      if (page.slideId !== snap.slideId || page.slideIds[slideIndex] !== snap.slideId)
        throw Error('presentation_baseline_changed')
      const metadata = await describePagePackageBackup(page.base64, signal)
      await guard(documentId, signal, token, true)
      const sourceSlideId = await sourceIdentity(page.base64, signal)
      await guard(documentId, signal, token, true)
      let r: PresentationSlideDuplicationBatch = {
        version: 4,
        kind: 'native_slide_duplicate',
        changeId: crypto.randomUUID(),
        documentId,
        baselineId: crypto.randomUUID(),
        baselineDigest: metadata.packageDigest,
        hostSlideId: page.slideId,
        slideIndex,
        sourceSlideId,
        beforeSlideIds: structuredClone(page.slideIds),
        scope: { slideIds: [page.slideId] },
        intent: explanation || 'Duplicate existing PowerPoint slide',
        preserved: ['Source page and original deck order'],
        validation: [
          'Exact source and copied page package proof',
          'Explicit insertion and removal receipts',
        ],
        risk: 'high',
        backups: [{ hostSlideId: page.slideId, backupId: crypto.randomUUID(), ...metadata }],
        operations: [{ op: 'duplicate_slide', slide_index: slideIndex }],
        nextIndex: 0,
        state: 'applying',
      }
      if (!validatePresentationExistingBatch(r)) throw Error('invalid_tool_input')
      const fresh = async (s?: AbortSignal) => {
        await guard(documentId, s, token, true)
        const exported = structuredClone(
          await options.adapter.exportPresentationPagePackage!(r.hostSlideId, s),
        )
        await guard(documentId, s, token, true)
        const digest = await presentationPackageDigest(exported.base64, s)
        await guard(documentId, s, token, true)
        if (
          exported.slideId !== r.hostSlideId ||
          !same(exported.slideIds, r.beforeSlideIds) ||
          digest !== r.baselineDigest
        )
          throw Error('presentation_baseline_changed')
      }
      const store = async (
        next: PresentationSlideDuplicationBatch,
        expected: PresentationSlideDuplicationBatch | undefined,
        writeSignal?: AbortSignal,
      ) => {
        await guard(documentId, writeSignal ?? signal, token, true)
        await options.writeExistingBatch(next, expected)
        r = structuredClone(next)
      }
      await fresh(signal)
      const proposal = options.proposals.propose({
        operation: toolName,
        toolName,
        title: r.intent,
        preview: {
          changeId: r.changeId,
          sourceSlideId: r.hostSlideId,
          slideIndex,
          packageSavepoint: true,
          operations: r.operations,
          qaScope: { basis: 'slide_duplication_savepoint', hostSlideIds: [r.hostSlideId] },
          qaPassed: false,
          uncertainInsertion:
            'Explicit package and order reconciliation; insertion is never replayed.',
        },
        impact: { host: 'powerpoint', targets: [r.hostSlideId], count: 1 },
        fingerprint: selectionFingerprint(JSON.stringify(r)),
        validate: async (s) => {
          try {
            await fresh(s)
            return !validateProduction || (await validateProduction(s))
          } catch {
            return false
          }
        },
        execute: async (s) => {
          await fresh(s)
          if (validateProduction && !(await validateProduction(s)))
            throw Error('office_concurrent_change')
          const scope = {
            request: options.request,
            documentId,
            hostSlideId: r.hostSlideId,
            slideIds: r.beforeSlideIds,
          }
          const b = r.backups[0]
          let retained: string
          try {
            const stored = await saveChartPackageBackup(
              { ...scope, backupId: b.backupId, base64: page.base64 },
              s,
            )
            await guard(documentId, s, token, true)
            if (stored.sha256 !== b.sha256 || stored.sizeBytes !== b.sizeBytes)
              throw Error('presentation_chart_backup_invalid')
            retained = await readChartPackageBackup(
              { ...scope, backup: b, expectedPackageDigest: r.baselineDigest },
              s,
            )
            await guard(documentId, s, token, true)
            const identity = await sourceIdentity(retained, s)
            await guard(documentId, s, token, true)
            if (identity !== r.sourceSlideId) throw Error('presentation_page_source_invalid')
            await fresh(s)
            await store(r, undefined, s)
          } catch (error) {
            let absent = false
            try {
              absent = options.readExistingBatch(r.changeId) === undefined
            } catch {
              // An unreadable intent may still own the savepoint.
            }
            if (absent)
              await Promise.allSettled([
                cleanupUncommittedChartPackageBackup({ ...scope, backup: b }),
              ])
            throw error
          }
          await current(r, s, token, true)
          await store({ ...r, inFlightIndex: 0 }, r, s)
          await current(r, s, token, true)
          await options.pageAdapter.stage(
            temporary(r),
            retained,
            async (newId) => {
              await current(r, s, token, true)
              await store({ ...r, insertedSlideId: newId }, r, s)
              await current(r, s, token, true)
            },
            () => current(r, s, token, true),
            s,
          )
          await current(r, s, token, true)
          const proof = await observe(r, s, token, true)
          if (proof.status !== 'inserted' || proof.insertedSlideId !== r.insertedSlideId)
            throw Error('presentation_slide_duplication_conflict')
          const { inFlightIndex: _flight, ...rest } = r
          await store({ ...rest, state: 'applied', nextIndex: 1 }, r, s)
          await current(r, s, token, true)
        },
        verify: async (s) => {
          const completed = saved(r.changeId)
          await current(completed, s, token, true)
          if (
            completed.state !== 'applied' ||
            (await observe(completed, s, token, true)).status !== 'inserted'
          )
            throw Error('office_verify_failed')
        },
      })
      return { proposalId: proposal.id, changeId: r.changeId, status: 'awaiting_confirmation' }
    },
    async executeTool(call: Parameters<AgentSkill['executeTool']>[0], signal?: AbortSignal) {
      const token = epoch
      try {
        call = { ...call, input: structuredClone(call.input) }
        const tool = tools.find((tool) => tool.name === call.name)
        const schema = tool?.inputSchema as
          { properties: Record<string, unknown>; required: string[] } | undefined
        if (
          !tool ||
          !schema ||
          call.inputError ||
          call.truncated ||
          Object.keys(call.input).some((key) => !Object.hasOwn(schema.properties, key)) ||
          schema.required.some((key) => !Object.hasOwn(call.input, key)) ||
          typeof call.input.change_id !== 'string' ||
          !/^[A-Za-z0-9_-]{1,128}$/.test(call.input.change_id)
        )
          throw Error('invalid_tool_input')
        let r = saved(call.input.change_id)
        await current(r, signal, token)
        if (call.name === 'inspect_slide_duplication') {
          let proof: { status: string; insertedSlideId?: string }
          try {
            proof = await observe(r, signal, token)
          } catch {
            await current(r, signal, token)
            proof = { status: 'unavailable' }
          }
          const expected =
            r.state === 'undone' ? 'baseline' : r.state === 'applied' ? 'inserted' : proof.status
          const matches =
            proof.status !== 'conflict' &&
            proof.status !== 'unavailable' &&
            proof.status === expected
          return {
            output: JSON.stringify({
              changeId: r.changeId,
              state: r.state,
              status: proof.status,
              hostStatus: proof.status === 'inserted' ? 'staged' : proof.status,
              insertedSlideId: r.insertedSlideId,
              observedInsertedSlideId: proof.insertedSlideId,
              currentPackageMatches: matches,
              currentHostVerified: matches,
              qaPassed: false,
              historicalOnly: true,
              inFlightIndex: r.inFlightIndex,
            }),
            mutated: false,
            summary: 'Inspected source and copied page packages without replay',
          }
        }
        const store = async (next: PresentationSlideDuplicationBatch, s?: AbortSignal) => {
          await current(r, s, token, true)
          await options.writeExistingBatch(next, r)
          r = structuredClone(next)
          await current(r, s, token, true)
        }
        if (call.name === 'reconcile_slide_duplication' || call.name === 'undo_slide_duplication') {
          const reconcile = call.name === 'reconcile_slide_duplication'
          if (reconcile ? r.state !== 'applying' : !['applied', 'undoing'].includes(r.state))
            throw Error('presentation_existing_batch_state_invalid')
          const originalProof = await observe(r, signal, token)
          if (
            reconcile
              ? (originalProof.status !== 'inserted' && originalProof.status !== 'baseline') ||
                (originalProof.status === 'inserted' &&
                  (!originalProof.insertedSlideId || r.inFlightIndex !== 0)) ||
                (originalProof.status === 'baseline' && r.insertedSlideId !== undefined)
              : originalProof.status !== 'inserted' &&
                (r.state !== 'undoing' || originalProof.status !== 'baseline')
          )
            throw Error('presentation_slide_duplication_conflict')
          await current(r, signal, token, true)
          const fresh = async (s?: AbortSignal) => {
            await current(r, s, token, true)
            const proof = await observe(r, s, token, true)
            if (!same(proof, originalProof)) throw Error('presentation_slide_duplication_conflict')
          }
          const proposal = options.proposals.propose({
            operation: call.name,
            toolName: call.name,
            title: reconcile ? 'Recognize owned copied slide' : 'Undo owned copied slide',
            preview: {
              changeId: r.changeId,
              sourceSlideId: r.hostSlideId,
              insertedSlideId: originalProof.insertedSlideId,
              receiptOnly: reconcile || originalProof.status === 'baseline',
              ...(reconcile
                ? {}
                : {
                    qaScope: {
                      basis: 'slide_duplication_savepoint',
                      hostSlideIds: [r.hostSlideId, r.insertedSlideId!],
                    },
                  }),
              qaPassed: false,
            },
            impact: {
              host: reconcile ? 'local_checkpoint' : 'powerpoint',
              targets: [
                r.hostSlideId,
                ...(originalProof.insertedSlideId ? [originalProof.insertedSlideId] : []),
              ],
              count: 1,
            },
            fingerprint: selectionFingerprint(JSON.stringify(r)),
            validate: async (s) => {
              try {
                await fresh(s)
                return true
              } catch {
                return false
              }
            },
            execute: async (s) => {
              await fresh(s)
              await backup(r, s, token, true)
              await fresh(s)
              if (reconcile) {
                const { inFlightIndex: _flight, ...rest } = r
                await store(
                  {
                    ...rest,
                    state: originalProof.status === 'baseline' ? 'undone' : 'applied',
                    nextIndex: originalProof.status === 'baseline' ? 0 : 1,
                    insertedSlideId:
                      originalProof.status === 'baseline'
                        ? undefined
                        : originalProof.insertedSlideId!,
                  },
                  s,
                )
                return
              }
              if (r.state === 'applied') await store({ ...r, state: 'undoing' }, s)
              if (originalProof.status === 'inserted') {
                await options.pageAdapter.discard(
                  temporary(r, 'discard_pending'),
                  () => current(r, s, token, true),
                  s,
                )
                await current(r, s, token, true)
              }
              if ((await observe(r, s, token, true)).status !== 'baseline')
                throw Error('presentation_slide_duplication_conflict')
              await store({ ...r, state: 'undone' }, s)
            },
            verify: async (s) => {
              const terminal = saved(r.changeId)
              await current(terminal, s, token, true)
              if (
                terminal.state !==
                  (reconcile && originalProof.status === 'inserted' ? 'applied' : 'undone') ||
                (await observe(terminal, s, token, true)).status !==
                  (reconcile ? originalProof.status : 'baseline')
              )
                throw Error('office_verify_failed')
            },
          })
          return {
            output: JSON.stringify({
              proposalId: proposal.id,
              changeId: r.changeId,
              status: 'awaiting_confirmation',
            }),
            mutated: false,
            summary: 'Prepared explicit copied-slide recovery confirmation',
          }
        }
        if (r.state !== 'applied' || !r.insertedSlideId || !options.adapter.inspectPresentationPage)
          throw Error('presentation_existing_batch_state_invalid')
        const reviewing = call.name === 'record_slide_duplication_page_review',
          visualToken = visualEpoch,
          capture = captures.get(r.changeId),
          input = call.input
        if (
          reviewing &&
          (!capture ||
            capture.visualEpoch !== visualToken ||
            capture.insertedSlideId !== r.insertedSlideId ||
            input.screenshot_digest !== capture.digest ||
            !['pass', 'fail'].includes(String(input.status)) ||
            typeof input.notes !== 'string' ||
            input.notes.length > 2000)
        )
          throw Error('presentation_slide_duplication_capture_stale')
        if ((await observe(r, signal, token)).status !== 'inserted')
          throw Error('presentation_slide_duplication_conflict')
        const shot = structuredClone(
          await options.adapter.inspectPresentationPage(r.insertedSlideId, signal),
        )
        await current(r, signal, token)
        if (
          shot.slideId !== r.insertedSlideId ||
          shot.shapesTruncated ||
          shot.screenshot.mime !== 'image/png'
        )
          throw Error('office_read_failed')
        const pngBase64 = validatePowerPointPageScreenshot(shot.screenshot.base64),
          digest = await screenshotDigest(pngBase64)
        await current(r, signal, token)
        if ((await observe(r, signal, token)).status !== 'inserted' || visualToken !== visualEpoch)
          throw Error('presentation_slide_duplication_capture_stale')
        if (!reviewing) {
          const capturedAt = new Date().toISOString()
          captures.set(r.changeId, {
            digest,
            capturedAt,
            visualEpoch: visualToken,
            insertedSlideId: r.insertedSlideId,
          })
          return {
            output: JSON.stringify({
              changeId: r.changeId,
              slideId: r.insertedSlideId,
              screenshotDigest: digest,
              capturedAt,
              screenshot: { mime: 'image/png', base64: pngBase64 },
              qaPassed: false,
              historicalOnly: true,
            }),
            mutated: false,
            summary: 'Captured the owned copied page for historical review',
          }
        }
        if (capture!.digest !== digest) throw Error('presentation_slide_duplication_capture_stale')
        const review = {
          hostSlideId: r.insertedSlideId,
          screenshotDigest: digest,
          capturedAt: capture!.capturedAt,
          reviewedAt: new Date().toISOString(),
          status: input.status as 'pass' | 'fail',
          notes: input.notes as string,
        }
        await store({ ...r, reviews: [review] }, signal)
        return {
          output: JSON.stringify({
            changeId: r.changeId,
            review,
            qaPassed: false,
            historicalOnly: true,
          }),
          mutated: false,
          summary: 'Recorded a historical copied-page assessment',
        }
      } catch (error) {
        return {
          output: error instanceof Error ? error.message : 'presentation_slide_duplication_failed',
          isError: true,
          mutated: false,
          summary: 'Copied-slide recovery stopped without replay',
        }
      }
    },
  }
}
