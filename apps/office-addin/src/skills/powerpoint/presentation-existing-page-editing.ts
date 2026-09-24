import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import JSZip from 'jszip'
import { selectionFingerprint, type ProposalPostWriteEvidence, type StructuredProposalController } from '../../agent/proposal-controller.js'
import { validatePowerPointPageScreenshot, type PowerPointPageInspection } from './browser-powerpoint-adapter.js'
import type { InMemoryVfs } from '../shared/vfs.js'
import type { PresentationBaselineSkill } from './presentation-baseline.js'
import type { PresentationPageReplacementAdapter } from './browser-presentation-page-replacement-adapter.js'
import type { PresentationPageReplacement } from './presentation-page-replacement-record.js'
import { presentationPackageDigest, MAX_PPTX_PACKAGE_BYTES } from './powerpoint-package.js'
import { validatePresentationExistingPageChange, type PresentationExistingPageChange } from './presentation-existing-page.js'

interface Options {
  baseline: PresentationBaselineSkill
  adapter: PresentationPageReplacementAdapter
  inspectPage(slideId: string, signal?: AbortSignal): Promise<Pick<PowerPointPageInspection, 'slideId' | 'shapesTruncated' | 'screenshot'>>
  exportAdapter: { exportPresentationPagePackage(slideId: string, signal?: AbortSignal): Promise<{slideId:string; slideIds:string[]; base64:string}> }
  vfs: InMemoryVfs
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  proposals: StructuredProposalController
  documentId(): Promise<string>
  readExistingPageChange(id: string): PresentationExistingPageChange | undefined
  writeExistingPageChange(record: PresentationExistingPageChange, expected: PresentationExistingPageChange | undefined): Promise<void>
  available(): boolean
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const id = (v: unknown): v is string => typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
const host = (v: unknown): v is string => typeof v === 'string' && v.length > 0 && v.length <= 256 && !Array.from(v).some((c) => c.charCodeAt(0) < 32)
const b64 = (bytes: Uint8Array) => btoa(Array.from(bytes, (x) => String.fromCharCode(x)).join(''))
const bytes = (base64: string) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
const sha = async (value: Uint8Array) => Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', Uint8Array.from(value).buffer)), (x) => x.toString(16).padStart(2, '0')).join('')
const output = (value: unknown) => JSON.stringify(value)
const CHUNK = 128 * 1024

function projected(r: PresentationExistingPageChange): PresentationPageReplacement {
  return {
    version: 1, changeId: r.changeId, documentId: r.documentId,
    projectId: 'existing', parentRequestId: 'baseline', requestId: r.changeId,
    pageId: 'existing', backupId: r.backup.backupId,
    parentArtifactDigest: r.baselineDigest, backupDigest: r.backup.sha256,
    originalPackageDigest: r.originalPackageDigest,
    replacementPackageDigest: r.replacementPackageDigest,
    sourceSlideId: r.sourceSlideId, oldSlideId: r.oldSlideId,
    beforeSlideIds: r.beforeSlideIds, state: r.state,
    ...(r.newSlideId ? { newSlideId: r.newSlideId } : {}),
    ...(r.restoredSlideId ? { restoredSlideId: r.restoredSlideId } : {}),
  }
}
async function oneSlide(bytesValue: Uint8Array, signal?: AbortSignal) {
  if (!bytesValue.length || bytesValue.length > MAX_PPTX_PACKAGE_BYTES) throw new Error('presentation_page_source_invalid')
  const base64 = b64(bytesValue)
  const digest = await presentationPackageDigest(base64, signal)
  const zip = await JSZip.loadAsync(bytesValue)
  const xml = await zip.file('ppt/presentation.xml')?.async('string')
  const ids = [...(xml ?? '').matchAll(/<p:sldId\b[^>]*\bid="([0-9]+)"[^>]*\/?\s*>/g)]
  const slidePaths = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  const relations = await zip.file('ppt/_rels/presentation.xml.rels')?.async('string')
  if (
    ids.length !== 1 ||
    Number(ids[0]![1]) < 256 ||
    Number(ids[0]![1]) > 4294967295 ||
    slidePaths.length !== 1 ||
    !relations?.includes(`Target="slides/${slidePaths[0]!.split('/').at(-1)}"`)
  )
    throw new Error('presentation_page_source_invalid')
  return { base64, digest, sourceSlideId: `${ids[0]![1]}#`, sha256: await sha(bytesValue), sizeBytes: bytesValue.length }
}
const names = ['stage', 'inspect', 'resume', 'commit', 'discard', 'undo'] as const
const tools: AgentToolDef[] = names.map((action) => ({
  name: `${action}_existing_presentation_page_change`,
  description: action === 'stage' ? 'Confirm backup and stage a one-slide PPTX after an existing native page; original remains.' : action === 'inspect' ? 'Inspect current native page order and content against a saved change.' : action === 'resume' ? 'Finish journal for a known inserted page without replaying insertion.' : action === 'commit' ? 'Confirm replacing the original with the verified staged page.' : action === 'discard' ? 'Confirm deleting the verified staged page while keeping the original.' : 'Confirm restoring the backed-up original page and removing the replacement.',
  inputSchema: { type: 'object', properties: action === 'stage'
    ? { baseline_id: { type: 'string' }, slide_id: { type: 'string' }, path: { type: 'string' }, explanation: { type: 'string' } }
    : { change_id: { type: 'string' } }, required: action === 'stage' ? ['baseline_id', 'slide_id', 'path'] : ['change_id'], additionalProperties: false },
}))

export function createPresentationExistingPageEditingSkill(options: Options): AgentSkill & { clear(): void } {
  let epoch = 0
  return {
    id: 'presentation-existing-page-editing',
    get tools() { return options.available() ? tools : tools.filter((t) => t.name.startsWith('inspect_')) },
    systemPrompt: 'Existing page rebuild uses a validated one-slide VFS PPTX. Stage retains the original. Inspect and resume interrupted insertion before further action; never replay unknown insertion. Commit and undo require separate confirmation. A saved record is not visual QA.',
    clear() { epoch++ },
    async executeTool(call, signal) {
      const token = epoch
      const active = () => { if (epoch !== token || signal?.aborted) throw new Error('cancelled') }
      try {
        const tool = tools.find((t) => t.name === call.name)
        if (!tool || call.inputError || call.truncated || !call.input || typeof call.input !== 'object') throw new Error('invalid_tool_input')
        const schema = tool.inputSchema as { properties: Record<string, unknown>; required: string[] }
        if (Object.keys(call.input).some((k) => !(k in schema.properties)) || schema.required.some((k) => !(k in call.input))) throw new Error('invalid_tool_input')
        const action = call.name.split('_')[0]
        if (action !== 'inspect' && !options.available()) throw new Error('presentation_page_backup_unavailable')
        const documentId = await options.documentId()
        const current = async () => { active(); if (await options.documentId() !== documentId) throw new Error('presentation_document_changed'); active() }
        const request = async (operation: string, data: Record<string, unknown>) => {
          await current()
          const response = await options.request({ operation, documentId, ...data }, signal)
          await current()
          const value = await response.json() as Record<string, unknown>
          if (!response.ok || !value || typeof value !== 'object' || 'error' in value) throw new Error('presentation_page_backup_failed')
          return value
        }
        let record: PresentationExistingPageChange
        let source: Awaited<ReturnType<typeof oneSlide>> | undefined
        let sourcePath: string | undefined
        let baseline = undefined as ReturnType<PresentationBaselineSkill['snapshot']>
        if (action === 'stage') {
          const input = call.input
          if (!id(input.baseline_id) || !host(input.slide_id) || typeof input.path !== 'string' || !input.path.endsWith('.pptx') || input.path.length > 1024 || (input.explanation !== undefined && (typeof input.explanation !== 'string' || input.explanation.length > 300))) throw new Error('invalid_tool_input')
          baseline = options.baseline.snapshot(input.baseline_id)
          if (!baseline || baseline.documentId !== documentId || !baseline.scope.slideIds.includes(input.slide_id) || baseline.scope.shapeIds?.length) throw new Error('presentation_existing_scope_mismatch')
          sourcePath = input.path
          source = await oneSlide(options.vfs.readBytes(sourcePath, { maxBytes: MAX_PPTX_PACKAGE_BYTES + 1 }), signal)
          const original = await options.exportAdapter.exportPresentationPagePackage(input.slide_id, signal)
          await current()
          if (original.slideId !== input.slide_id || !same(original.slideIds, baseline.context.slideIds)) throw new Error('presentation_baseline_changed')
          const originalPackageDigest = await presentationPackageDigest(original.base64, signal)
          record = { version: 1, changeId: crypto.randomUUID(), documentId, baselineId: baseline.baselineId, baselineDigest: baseline.contentDigest, scope: { slideIds: [...baseline.scope.slideIds] }, oldSlideId: input.slide_id, beforeSlideIds: [...baseline.context.slideIds], originalPackageDigest, replacementPackageDigest: source.digest, sourceSlideId: source.sourceSlideId, backup: { backupId: crypto.randomUUID(), sha256: '0'.repeat(64), sizeBytes: 1 }, state: 'pending' }
        } else {
          if (!id(call.input.change_id)) throw new Error('invalid_tool_input')
          const saved = options.readExistingPageChange(call.input.change_id)
          if (!saved || !validatePresentationExistingPageChange(saved) || saved.documentId !== documentId) throw new Error('presentation_existing_page_missing')
          record = structuredClone(saved)
        }
        let expected = action === 'stage' ? undefined : structuredClone(record)
        const saved = () => { if (!same(options.readExistingPageChange(record.changeId), expected)) throw new Error('presentation_existing_page_stale') }
        const assertCurrent = async () => { await current(); saved() }
        const store = async (next: PresentationExistingPageChange) => {
          await current(); saved(); await options.writeExistingPageChange(next, expected); await current()
          if (!same(options.readExistingPageChange(next.changeId), next)) throw new Error('office_state_uncertain')
          record = structuredClone(next); expected = structuredClone(next)
        }
        const baselineFresh = async () => {
          if (!baseline || !same(options.baseline.snapshot(baseline.baselineId), baseline)) return false
          const result = await options.baseline.executeTool({ id: 'page-check', name: 'check_presentation_baseline', input: { baseline_id: baseline.baselineId } }, signal)
          await current()
          return !result.isError && JSON.parse(result.output).unchanged === true && same(options.baseline.snapshot(baseline.baselineId), baseline)
        }
        const sourceFresh = async () => {
          if (!source || !sourcePath) return
          const value = options.vfs.readBytes(sourcePath, { maxBytes: MAX_PPTX_PACKAGE_BYTES + 1 })
          if (await sha(value) !== source.sha256 || value.length !== source.sizeBytes) throw new Error('proposal_stale')
          await current()
        }
        const inspect = async () => { await current(); saved(); const state = await options.adapter.inspect(projected(record), signal); await current(); saved(); return state }
        if (action === 'inspect') {
          const observed = await inspect()
          const expectedStatuses: Record<PresentationExistingPageChange['state'], string[]> = {
            pending: [],
            inserted: ['staged'],
            staged: ['staged'],
            commit_pending: ['staged', 'applied'],
            applied: ['applied'],
            discard_pending: ['staged', 'baseline'],
            discarded: ['baseline'],
            undo_pending: ['applied'],
            restore_inserted: ['restore_staged', 'undone'],
            undone: ['undone'],
          }
          const verified = expectedStatuses[record.state].includes(observed.status)
          return { output: output({ changeId: record.changeId, state: record.state, inspection: observed, currentHostVerified: verified, manualReview: !verified, qaPassed: false }), mutated: false, summary: '已检查现稿单页变更' }
        }
        const observed = action === 'stage' ? undefined : await inspect()
        if (action === 'stage') {
          if (!(await baselineFresh())) throw new Error('presentation_baseline_changed')
        } else if (action === 'resume') {
          if (!['inserted', 'commit_pending', 'discard_pending', 'restore_inserted'].includes(record.state)) throw new Error('presentation_existing_page_manual_review')
          if (!observed || observed.status === 'conflict') throw new Error('presentation_existing_page_manual_review')
        } else if (action === 'commit' && (record.state !== 'staged' && record.state !== 'commit_pending' || !observed || !['staged', 'applied'].includes(observed.status))) throw new Error('presentation_existing_page_conflict')
        else if (action === 'discard' && (record.state !== 'staged' && record.state !== 'discard_pending' || !observed || !['staged', 'baseline'].includes(observed.status))) throw new Error('presentation_existing_page_conflict')
        else if (action === 'undo' && (record.state !== 'applied' && record.state !== 'undo_pending' && record.state !== 'restore_inserted' || !observed || !['applied', 'restore_staged', 'undone'].includes(observed.status))) throw new Error('presentation_existing_page_conflict')
        const initial = observed?.status
        const loadBackup = async () => {
          const status = await request('existing_page_backup_status', { backupId: record.backup.backupId })
          if (status.status !== 'ready' || status.backupId !== record.backup.backupId || status.documentId !== documentId || status.hostSlideId !== record.oldSlideId || !same(status.slideIds, record.beforeSlideIds) || status.sha256 !== record.backup.sha256 || status.sizeBytes !== record.backup.sizeBytes || status.receivedBytes !== status.sizeBytes) throw new Error('presentation_page_backup_invalid')
          const content = new Uint8Array(record.backup.sizeBytes)
          for (let offset = 0; offset < content.length; offset += CHUNK) {
            const length = Math.min(CHUNK, content.length - offset)
            const part = await request('existing_page_backup_read', { backupId: record.backup.backupId, offset, length })
            if (part.backupId !== record.backup.backupId || part.offset !== offset || part.sizeBytes !== content.length || part.sha256 !== record.backup.sha256 || typeof part.base64 !== 'string') throw new Error('presentation_page_backup_invalid')
            const chunk = bytes(part.base64)
            if (chunk.length !== length || b64(chunk) !== part.base64) throw new Error('presentation_page_backup_invalid')
            content.set(chunk, offset)
          }
          if (await sha(content) !== record.backup.sha256 || await presentationPackageDigest(b64(content), signal) !== record.originalPackageDigest) throw new Error('presentation_page_backup_invalid')
          await current()
          return b64(content)
        }
        const proposal = options.proposals.propose({
          operation: call.name, toolName: call.name,
          title: action === 'stage' ? (call.input.explanation as string) || '暂存现稿单页重建' : `${action} 现稿单页重建`,
          preview: { oldSlideId: record.oldSlideId, newSlideId: record.newSlideId, originalDigest: record.originalPackageDigest, replacementDigest: record.replacementPackageDigest, stageKeepsOriginal: action === 'stage', qaPassed: false },
          impact: { host: 'powerpoint', targets: [record.oldSlideId], count: 1 },
          fingerprint: selectionFingerprint(output(record)),
          validate: async () => {
            try {
              await sourceFresh()
              if (action !== 'stage') return same((await inspect()).status, initial)
              if (!(await baselineFresh())) return false
              const page = await options.exportAdapter.exportPresentationPagePackage(record.oldSlideId, signal)
              await current()
              return page.slideId === record.oldSlideId && same(page.slideIds, record.beforeSlideIds) && await presentationPackageDigest(page.base64, signal) === record.originalPackageDigest
            } catch { return false }
          },
          execute: async () => {
            await sourceFresh()
            if (action === 'stage') {
              if (!(await baselineFresh())) throw new Error('proposal_stale')
              const exported = await options.exportAdapter.exportPresentationPagePackage(record.oldSlideId, signal)
              await current()
              if (exported.slideId !== record.oldSlideId || !same(exported.slideIds, record.beforeSlideIds)) throw new Error('proposal_stale')
              const originalBytes = bytes(exported.base64)
              if (!originalBytes.length || originalBytes.length > MAX_PPTX_PACKAGE_BYTES) throw new Error('presentation_page_backup_invalid')
              const originalPackageDigest = await presentationPackageDigest(exported.base64, signal)
              if (originalPackageDigest !== record.originalPackageDigest) throw new Error('proposal_stale')
              const backupSha = await sha(originalBytes)
              const match = (value: Record<string, unknown>) => value.backupId === record.backup.backupId && value.documentId === documentId && value.hostSlideId === record.oldSlideId && same(value.slideIds, record.beforeSlideIds) && value.sha256 === backupSha && value.sizeBytes === originalBytes.length
              let meta = await request('existing_page_backup_begin', { backupId: record.backup.backupId, hostSlideId: record.oldSlideId, slideIds: record.beforeSlideIds, sha256: backupSha, sizeBytes: originalBytes.length })
              if (!match(meta)) throw new Error('presentation_page_backup_invalid')
              while (meta.status !== 'ready' && Number(meta.receivedBytes) < originalBytes.length) {
                const offset = Number(meta.receivedBytes)
                if (!Number.isSafeInteger(offset) || offset < 0) throw new Error('presentation_page_backup_invalid')
                const chunk = originalBytes.subarray(offset, Math.min(offset + CHUNK, originalBytes.length))
                meta = await request('existing_page_backup_chunk', { backupId: record.backup.backupId, offset, base64: b64(chunk) })
                if (!match(meta) || Number(meta.receivedBytes) !== offset + chunk.length) throw new Error('presentation_page_backup_invalid')
              }
              if (meta.status !== 'ready') meta = await request('existing_page_backup_finish', { backupId: record.backup.backupId })
              if (!match(meta) || meta.status !== 'ready' || meta.receivedBytes !== originalBytes.length) throw new Error('presentation_page_backup_invalid')
              await sourceFresh()
              if (!(await baselineFresh())) throw new Error('proposal_stale')
              const beforeWrite = await options.exportAdapter.exportPresentationPagePackage(record.oldSlideId, signal)
              if (beforeWrite.slideId !== record.oldSlideId || !same(beforeWrite.slideIds, record.beforeSlideIds) || await presentationPackageDigest(beforeWrite.base64, signal) !== originalPackageDigest) throw new Error('proposal_stale')
              record = { ...record, originalPackageDigest, backup: { backupId: record.backup.backupId, sha256: backupSha, sizeBytes: originalBytes.length } }
              await store(record)
              await sourceFresh()
              if (!(await baselineFresh())) throw new Error('proposal_stale')
              await options.adapter.stage(projected(record), source!.base64, async (newSlideId) => { await store({ ...record, state: 'inserted', newSlideId }) }, assertCurrent, signal)
              await store({ ...record, state: 'staged' })
            } else if (action === 'resume') {
              if (record.state === 'inserted' && initial === 'staged') await store({ ...record, state: 'staged' })
              else if (record.state === 'commit_pending' && initial === 'applied') await store({ ...record, state: 'applied' })
              else if (record.state === 'discard_pending' && initial === 'baseline') await store({ ...record, state: 'discarded' })
              else if (record.state === 'restore_inserted' && initial === 'undone') await store({ ...record, state: 'undone' })
              else throw new Error('presentation_existing_page_manual_review')
            } else if (action === 'commit') {
              await loadBackup()
              if ((await inspect()).status !== initial) throw new Error('proposal_stale')
              if (record.state === 'staged') await store({ ...record, state: 'commit_pending' })
              await options.adapter.commit(projected(record), assertCurrent, signal)
              await store({ ...record, state: 'applied' })
            } else if (action === 'discard') {
              if (record.state === 'staged') await store({ ...record, state: 'discard_pending' })
              await options.adapter.discard(projected(record), assertCurrent, signal)
              await store({ ...record, state: 'discarded' })
            } else {
              const backup = await loadBackup()
              if ((await inspect()).status !== initial) throw new Error('proposal_stale')
              if (record.state === 'applied') await store({ ...record, state: 'undo_pending' })
              await options.adapter.undo(projected(record), backup, async (restoredSlideId) => { await store({ ...record, state: 'restore_inserted', restoredSlideId }) }, assertCurrent, signal)
              await store({ ...record, state: 'undone' })
            }
          },
          verify: async () => { saved(); const status = (await options.adapter.inspect(projected(record), signal)).status; if (status !== (action === 'stage' ? 'staged' : action === 'discard' ? 'baseline' : action === 'undo' ? 'undone' : action === 'resume' ? initial : 'applied')) throw new Error('office_state_uncertain') },
          postWrite: async (): Promise<ProposalPostWriteEvidence> => {
            const targets = record.state === 'staged'
              ? [record.oldSlideId, record.newSlideId!]
              : record.state === 'applied'
                ? [record.newSlideId!]
                : record.state === 'discarded'
                  ? [record.oldSlideId]
                  : record.state === 'undone'
                    ? [record.restoredSlideId!]
                    : []
            if (!targets.length) throw new Error('office_state_uncertain')
            const status = record.state === 'staged' ? 'staged' : record.state === 'applied' ? 'applied' : record.state === 'discarded' ? 'baseline' : 'undone'
            const check = async () => {
              await current(); saved()
              const observed = await options.adapter.inspect(projected(record))
              await current(); saved()
              if (observed.status !== status || !targets.every((id) => observed.slideIds.includes(id))) throw new Error('office_state_uncertain')
              return observed.slideIds
            }
            const before = await check()
            const pages = [] as {slideId:string;pngBase64:string;digest:string}[]
            for (const slideId of targets) {
              const shot = await options.inspectPage(slideId)
              if (shot.slideId !== slideId || shot.shapesTruncated || shot.screenshot.mime !== 'image/png') throw new Error('office_read_failed')
              const pngBase64 = validatePowerPointPageScreenshot(shot.screenshot.base64)
              pages.push({ slideId, pngBase64, digest: await sha(bytes(pngBase64)) })
              if (!same(await check(), before)) throw new Error('office_state_uncertain')
            }
            return { status: 'captured', pages }
          },
        })
        return { output: output({ proposalId: proposal.id, status: 'awaiting_confirmation', changeId: record.changeId, state: record.state, oldSlideId: record.oldSlideId, newSlideId: record.newSlideId }), mutated: false, summary: '已准备现稿单页变更，等待确认' }
      } catch (error) {
        const message = error instanceof Error ? error.message : ''
        return { output: /^(presentation_[a-z_]+|office_[a-z_]+|invalid_tool_input|proposal_stale|cancelled|vfs_[a-z_]+)$/.test(message) ? message : 'presentation_existing_page_failed', isError: true, mutated: false, summary: '现稿单页操作未完成' }
      }
    },
  }
}
