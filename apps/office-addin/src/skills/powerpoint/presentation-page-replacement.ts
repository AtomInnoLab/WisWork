import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import {
  selectionFingerprint,
  type StructuredProposalController,
} from '../../agent/proposal-controller.js'
import type {
  CompiledPresentationArtifact,
  PresentationImportRecord,
} from './presentation-delivery.js'
import {
  presentationArtifactContent,
  presentationImportKey,
  presentationPageMapping,
  validPresentationImportRecord,
} from './presentation-page-delivery.js'
import {
  parsePresentationPageArtifact,
  parsePresentationProductionStatus,
} from './presentation-production.js'
import type { PresentationPageBackupMetadata } from './presentation-page-backup.js'
import { presentationPackageDigest } from './powerpoint-package.js'
import {
  validatePresentationPageReplacement,
  type PresentationPageReplacement,
} from './presentation-page-replacement-record.js'
import type { PresentationPageReplacementAdapter } from './browser-presentation-page-replacement-adapter.js'

export interface PresentationPageReplacementOptions {
  available(): boolean
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  documentId(): Promise<string>
  artifact(projectId?: string): CompiledPresentationArtifact | undefined
  readReceipt(key: string): PresentationImportRecord | undefined
  readPageReplacement(): PresentationPageReplacement | undefined
  writePageReplacement(
    record: PresentationPageReplacement,
    expected: PresentationPageReplacement | undefined,
  ): Promise<void>
  loadBackup(
    projectId: string,
    backupId: string,
    signal?: AbortSignal,
  ): Promise<{ metadata: PresentationPageBackupMetadata; base64: string }>
  adapter: PresentationPageReplacementAdapter
  proposals: StructuredProposalController
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const id = (v: unknown, max: number): v is string =>
  typeof v === 'string' && new RegExp(`^[A-Za-z0-9_-]{1,${max}}$`).test(v)
async function digest(bytes: Uint8Array) {
  return Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes).buffer)),
    (b) => b.toString(16).padStart(2, '0'),
  ).join('')
}
const names = ['stage', 'inspect', 'resume', 'discard', 'commit', 'undo'] as const
const tools: AgentToolDef[] = names.map((action) => ({
  name: `${action}_presentation_page_replacement`,
  description:
    action === 'stage'
      ? 'Propose inserting a compiled revision after its backed-up original page, retaining the original and existing business mapping. Requires explicit confirmation. This stages a replacement; it does not finish replacement.'
      : action === 'inspect'
        ? 'Inspect a durable staged replacement without writing; uncertain pending insertion requires manual review.'
        : action === 'resume'
          ? 'Propose completing only the receipt for an already known, verified inserted page. Never repeats insertion.'
          : action === 'commit'
            ? 'Confirm replacing the backed-up original with its verified staged revision, atomically switching the business mapping. Reuse to recover commit_pending. Explicitly prepare the child request after completion.'
            : action === 'undo'
              ? 'Confirm restoring the verified original backup before removing the revision. Reuse to recover undo_pending/restore_inserted; explicitly prepare the parent request after completion.'
              : 'Propose removing only the unchanged owned staged page. Keeps the original page and backup; requires explicit confirmation.',
  inputSchema: {
    type: 'object',
    properties: {
      project_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,80}$' },
      change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
      ...(action === 'stage'
        ? {
            request_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
            page_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,80}$' },
            backup_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
          }
        : {}),
    },
    required:
      action === 'stage'
        ? ['project_id', 'change_id', 'request_id', 'page_id', 'backup_id']
        : ['project_id', 'change_id'],
    additionalProperties: false,
  },
}))
export function createPresentationPageReplacementSkill(
  options: PresentationPageReplacementOptions,
): AgentSkill & { clear(): void } {
  let epoch = 0
  return {
    id: 'office-presentation-page-replacement',
    get tools() {
      return options.available() ? tools : []
    },
    systemPrompt:
      'Stage a backed-up revision with explicit confirmation, retaining the original. Inspect the durable transaction; resume only completes a known staged insertion. Commit verifies both pages and the backup, deletes the original and switches the business mapping. Undo restores the original backup before deleting the revision. Reuse commit for commit_pending and undo for undo_pending/restore_inserted recovery; uncertain unknown inserted IDs require manual review. After commit explicitly prepare the child request; after undo explicitly prepare the parent request before editing or QA. Discard removes an unchanged staged page only. Never claim visual, content or round-trip QA acceptance from replacement success.',
    clear() {
      epoch++
    },
    async executeTool(call, signal) {
      const captured = epoch
      const check = (s?: AbortSignal) => {
        if (epoch !== captured || signal?.aborted || s?.aborted) throw new Error('cancelled')
        if (!options.available()) throw new Error('presentation_unavailable')
      }
      try {
        check()
        const action = names.find((a) => call.name === `${a}_presentation_page_replacement`)
        const input = structuredClone(call.input) as Record<string, unknown>
        const keys =
          action === 'stage'
            ? ['project_id', 'change_id', 'request_id', 'page_id', 'backup_id']
            : ['project_id', 'change_id']
        if (
          !action ||
          call.inputError ||
          call.truncated ||
          !input ||
          typeof input !== 'object' ||
          Array.isArray(input) ||
          Object.keys(input).length !== keys.length ||
          Object.keys(input).some((k) => !keys.includes(k)) ||
          !id(input.project_id, 80) ||
          !id(input.change_id, 128) ||
          (action === 'stage' &&
            (!id(input.request_id, 128) || !id(input.page_id, 80) || !id(input.backup_id, 128)))
        )
          throw new Error('invalid_tool_input')
        const projectId = input.project_id,
          changeId = input.change_id
        const documentId = await options.documentId()
        check()
        if (typeof documentId !== 'string' || !documentId || documentId.length > 4096)
          throw new Error('presentation_document_changed')
        const artifact = options.artifact(projectId)
        if (
          !artifact ||
          artifact.documentId !== documentId ||
          artifact.projectId !== projectId ||
          !artifact.pagePptxBase64
        )
          throw new Error('presentation_restore_required')
        const snapshot = JSON.stringify(artifact),
          key = presentationImportKey(artifact),
          receipt = options.readReceipt(key),
          receiptJson = JSON.stringify(receipt)
        let expectedReceipt = receiptJson
        const expectedOtherReceipts = new Map<string, string | undefined>()
        const receiptSnapshot = (receiptKey: string) => {
          try {
            return JSON.stringify(options.readReceipt(receiptKey))
          } catch (error) {
            if (error instanceof Error && error.message === 'presentation_import_superseded')
              return 'superseded'
            throw error
          }
        }
        if (
          !validPresentationImportRecord(receipt) ||
          receipt.state !== 'complete' ||
          receipt.documentId !== documentId ||
          receipt.checkpoint?.version !== 2 ||
          !same(
            receipt.checkpoint.sourceSlideIds,
            artifact.pages?.map((p) => p.sourceSlideId),
          )
        )
          throw new Error('presentation_page_binding_invalid')
        const current = async (s?: AbortSignal) => {
          check(s)
          if ((await options.documentId()) !== documentId)
            throw new Error('presentation_document_changed')
          check(s)
          if (
            options.artifact(projectId) !== artifact ||
            JSON.stringify(artifact) !== snapshot ||
            receiptSnapshot(key) !== expectedReceipt ||
            [...expectedOtherReceipts].some(([k, v]) => receiptSnapshot(k) !== v)
          )
            throw new Error('proposal_stale')
        }
        const parentArtifactDigest = await digest(
          new TextEncoder().encode(presentationArtifactContent(artifact)),
        )
        await current()
        if (parentArtifactDigest !== receipt.checkpoint.artifactDigest)
          throw new Error('presentation_page_binding_invalid')
        const read = () => {
          const value = options.readPageReplacement()
          if (value !== undefined && !validatePresentationPageReplacement(value))
            throw new Error('presentation_page_replacement_invalid')
          return value === undefined ? undefined : structuredClone(value)
        }
        let latest = read()
        let record: PresentationPageReplacement
        let base64 = ''
        let commitReceipts:
          Pick<PresentationPageReplacement, 'parentReceipt' | 'childReceipt'> | undefined
        const request = async (operation: string, body: Record<string, unknown>) => {
          await current()
          const response = await options.request(
            { operation, documentId, projectId, ...body },
            signal,
          )
          await current()
          const text = await response.text()
          await current()
          if (text.length > 15 * 1024 * 1024) throw new Error('presentation_response_invalid')
          const value: unknown = JSON.parse(text)
          if (!response.ok) throw new Error('presentation_response_invalid')
          return value
        }
        if (action === 'stage') {
          if (
            latest &&
            (!['discarded', 'undone'].includes(latest.state) || latest.changeId === changeId)
          )
            throw new Error('presentation_page_replacement_uncertain')
          const mapping = presentationPageMapping(artifact, receipt, input.page_id as string)
          if (!mapping) throw new Error('presentation_page_binding_invalid')
          const status = parsePresentationProductionStatus(
            await request('production_status', { requestId: input.request_id }),
          )
          if (
            status.projectId !== projectId ||
            status.requestId !== input.request_id ||
            status.status !== 'compiled' ||
            status.revision?.parentRequestId !== artifact.requestId ||
            status.revision.pageId !== input.page_id ||
            status.planRevision !== artifact.planRevision ||
            !same(
              status.pages.map((p) => p.id),
              artifact.pages?.map((p) => p.id),
            )
          )
            throw new Error('presentation_page_binding_invalid')
          const page = parsePresentationPageArtifact(
            await request('production_page', {
              requestId: input.request_id,
              pageId: input.page_id,
            }),
            projectId,
            input.request_id as string,
            input.page_id as string,
          )
          if (page.planRevision !== status.planRevision)
            throw new Error('presentation_page_binding_invalid')
          base64 = page.pptxBase64
          const backup = await options.loadBackup(projectId, input.backup_id as string, signal)
          await current()
          const m = backup.metadata
          if (
            m.status !== 'ready' ||
            m.documentId !== documentId ||
            m.projectId !== projectId ||
            m.backupId !== input.backup_id ||
            m.requestId !== status.requestId ||
            m.parentRequestId !== artifact.requestId ||
            m.pageId !== input.page_id ||
            m.hostSlideId !== mapping.slideId ||
            m.parentInputDigest !== status.revision.parentInputDigest ||
            m.receivedBytes !== m.sizeBytes
          )
            throw new Error('presentation_page_binding_invalid')
          const binary = atob(backup.base64),
            bytes = Uint8Array.from(binary, (c) => c.charCodeAt(0))
          if (
            btoa(binary) !== backup.base64 ||
            bytes.length !== m.sizeBytes ||
            (await digest(bytes)) !== m.sha256
          )
            throw new Error('presentation_page_binding_invalid')
          const originalPackageDigest = await presentationPackageDigest(backup.base64, signal),
            replacementPackageDigest = await presentationPackageDigest(base64, signal)
          await current()
          record = {
            version: 1,
            changeId,
            documentId,
            projectId,
            parentRequestId: artifact.requestId,
            requestId: status.requestId,
            pageId: input.page_id as string,
            backupId: m.backupId,
            parentArtifactDigest,
            backupDigest: m.sha256,
            originalPackageDigest,
            replacementPackageDigest,
            sourceSlideId: page.sourceSlideId,
            oldSlideId: m.hostSlideId,
            beforeSlideIds: [...m.slideIds],
            state: 'pending',
          }
          if (!validatePresentationPageReplacement(record))
            throw new Error('presentation_page_binding_invalid')
        } else {
          if (
            !latest ||
            latest.changeId !== changeId ||
            latest.projectId !== projectId ||
            latest.documentId !== documentId ||
            (action === 'undo' ||
            (action === 'inspect' &&
              ['applied', 'undo_pending', 'restore_inserted'].includes(latest.state))
              ? latest.requestId !== artifact.requestId ||
                latest.childReceipt?.checkpoint?.artifactDigest !== parentArtifactDigest ||
                presentationPageMapping(artifact, receipt, latest.pageId)?.slideId !==
                  latest.newSlideId
              : latest.parentRequestId !== artifact.requestId ||
                latest.parentArtifactDigest !== parentArtifactDigest ||
                presentationPageMapping(artifact, receipt, latest.pageId)?.slideId !==
                  (latest.state === 'undone' ? latest.restoredSlideId : latest.oldSlideId))
          )
            throw new Error('presentation_page_binding_invalid')
          record = structuredClone(latest)
        }
        let verifiedBackupMetadata: string | undefined
        // Recheck the durable PC backup for destructive transitions, not just staging.
        const verifyBackup = async (s?: AbortSignal) => {
          await current(s)
          const backup = await options.loadBackup(projectId, record.backupId, s)
          await current(s)
          const m = backup.metadata,
            binary = atob(backup.base64)
          if (verifiedBackupMetadata !== undefined && JSON.stringify(m) !== verifiedBackupMetadata)
            throw new Error('presentation_page_binding_invalid')
          if (
            m.status !== 'ready' ||
            m.documentId !== documentId ||
            m.projectId !== projectId ||
            m.backupId !== record.backupId ||
            m.parentRequestId !== record.parentRequestId ||
            m.requestId !== record.requestId ||
            m.pageId !== record.pageId ||
            m.hostSlideId !== record.oldSlideId ||
            !same(m.slideIds, record.beforeSlideIds) ||
            m.sha256 !== record.backupDigest ||
            m.receivedBytes !== m.sizeBytes ||
            binary.length !== m.sizeBytes ||
            btoa(binary) !== backup.base64 ||
            (await digest(Uint8Array.from(binary, (c) => c.charCodeAt(0)))) !==
              record.backupDigest ||
            (await presentationPackageDigest(backup.base64, s)) !== record.originalPackageDigest
          )
            throw new Error('presentation_page_binding_invalid')
          await current(s)
          verifiedBackupMetadata = JSON.stringify(m)
          return backup
        }
        if (action === 'commit' || action === 'undo') {
          const otherRequest = action === 'commit' ? record.requestId : record.parentRequestId
          const otherKey = `production/${projectId}/${otherRequest}`
          expectedOtherReceipts.set(otherKey, receiptSnapshot(otherKey))
          const backup = await verifyBackup(signal)
          base64 = backup.base64
          if (action === 'commit') {
            if (!['staged', 'commit_pending'].includes(record.state))
              throw new Error('presentation_page_replacement_uncertain')
            const status = parsePresentationProductionStatus(
              await request('production_status', { requestId: record.requestId }),
            )
            if (
              status.projectId !== projectId ||
              status.requestId !== record.requestId ||
              status.status !== 'compiled' ||
              status.revision?.parentRequestId !== record.parentRequestId ||
              status.revision.pageId !== record.pageId ||
              status.revision.parentInputDigest !== backup.metadata.parentInputDigest ||
              status.planRevision !== artifact.planRevision ||
              !same(
                status.pages.map((p) => ({ id: p.id, title: p.title })),
                artifact.pages?.map((p) => ({ id: p.id, title: p.title })),
              )
            )
              throw new Error('presentation_page_binding_invalid')
            const pages: NonNullable<CompiledPresentationArtifact['pages']> = [],
              pagePptxBase64: string[] = []
            let bytes = 0
            for (const planned of status.pages) {
              const page = parsePresentationPageArtifact(
                await request('production_page', {
                  requestId: record.requestId,
                  pageId: planned.id,
                }),
                projectId,
                record.requestId,
                planned.id,
              )
              bytes += page.binary.length
              if (bytes > 10 * 1024 * 1024) throw new Error('presentation_output_too_large')
              const index = pages.length
              if (
                page.planRevision !== status.planRevision ||
                (planned.id === record.pageId
                  ? page.sourceSlideId !== record.sourceSlideId ||
                    (await presentationPackageDigest(page.pptxBase64, signal)) !==
                      record.replacementPackageDigest
                  : page.pptxBase64 !== artifact.pagePptxBase64![index] ||
                    page.sourceSlideId !== artifact.pages![index].sourceSlideId)
              )
                throw new Error('presentation_page_binding_invalid')
              await current()
              pages.push({
                id: planned.id,
                title: planned.title,
                sourceSlideId: page.sourceSlideId,
              })
              pagePptxBase64.push(page.pptxBase64)
            }
            const childArtifact: CompiledPresentationArtifact = {
              documentId,
              projectId,
              requestId: record.requestId,
              slideCount: pages.length,
              pptxBase64: '',
              pages,
              pagePptxBase64,
              planRevision: status.planRevision,
            }
            const childDigest = await digest(
              new TextEncoder().encode(presentationArtifactContent(childArtifact)),
            )
            await current()
            const childReceipt = structuredClone(receipt)
            const index = pages.findIndex((p) => p.id === record.pageId)
            childReceipt.slideIds![index] = record.newSlideId!
            childReceipt.checkpoint!.artifactDigest = childDigest
            childReceipt.checkpoint!.sourceSlideIds = pages.map((p) => p.sourceSlideId)
            childReceipt.checkpoint!.completed[index] = {
              sourceSlideId: record.sourceSlideId,
              slideId: record.newSlideId!,
            }
            if (
              !validPresentationImportRecord(childReceipt) ||
              (record.state === 'commit_pending' &&
                (!same(record.parentReceipt, receipt) || !same(record.childReceipt, childReceipt)))
            )
              throw new Error('presentation_page_binding_invalid')
            commitReceipts = { parentReceipt: structuredClone(receipt), childReceipt }
          } else if (
            !['applied', 'undo_pending', 'restore_inserted'].includes(record.state) ||
            !same(record.childReceipt, receipt)
          ) {
            throw new Error('presentation_page_replacement_uncertain')
          }
        }
        const unchanged = async (s?: AbortSignal) => {
          await current(s)
          if (!same(read(), latest)) throw new Error('proposal_stale')
        }
        const inspect = async (s?: AbortSignal) => {
          await unchanged(s)
          const value = await options.adapter.inspect(record, s)
          await unchanged(s)
          return value
        }
        const initial = await inspect(signal)
        if (action === 'inspect')
          return {
            output: JSON.stringify({
              ...record,
              stateSource: 'persisted_checkpoint',
              inspection: initial,
              originalRetained: initial.slideIds.includes(record.oldSlideId),
              originalContentVerified: ['baseline', 'staged', 'restore_staged', 'undone'].includes(
                initial.status,
              ),
              businessMappingUpdated: [
                'applied',
                'undo_pending',
                'restore_inserted',
                'undone',
              ].includes(record.state),
            }),
            mutated: false,
            summary: '已核对当前页面；事务状态为持久检查点，内容验收需另行执行',
          }
        if (
          (action === 'stage' && initial.status !== 'baseline') ||
          (action === 'resume' && (record.state !== 'inserted' || initial.status !== 'staged')) ||
          (action === 'discard' &&
            (!['staged', 'discard_pending'].includes(record.state) ||
              (initial.status !== 'staged' &&
                !(record.state === 'discard_pending' && initial.status === 'baseline'))))
        )
          throw new Error('presentation_page_replacement_uncertain')
        if (
          (action === 'commit' &&
            !(
              initial.status === 'staged' ||
              (record.state === 'commit_pending' && initial.status === 'applied')
            )) ||
          (action === 'undo' &&
            !(record.state === 'restore_inserted'
              ? ['restore_staged', 'undone'].includes(initial.status)
              : initial.status === 'applied'))
        )
          throw new Error('presentation_page_replacement_uncertain')
        const stable = async (s?: AbortSignal) => {
          const value = await inspect(s)
          if (!same(value, initial)) throw new Error('proposal_stale')
        }
        const save = async (next: PresentationPageReplacement, s?: AbortSignal) => {
          await unchanged(s)
          await options.writePageReplacement(next, latest)
          // Only the exact mapping published by our successful terminal CAS may change.
          if (next.state === 'applied' || next.state === 'undone') {
            expectedReceipt = 'superseded'
            const otherKey = `production/${projectId}/${next.state === 'applied' ? next.requestId : next.parentRequestId}`
            const other = structuredClone(
              next.state === 'applied' ? next.childReceipt! : next.parentReceipt!,
            )
            if (next.state === 'undone') {
              const index = other.checkpoint!.completed.findIndex(
                (p) => p.slideId === next.oldSlideId,
              )
              other.slideIds![index] = next.restoredSlideId!
              other.checkpoint!.completed[index].slideId = next.restoredSlideId!
            }
            expectedOtherReceipts.set(otherKey, JSON.stringify(other))
          }
          await current(s)
          if (!same(read(), next)) throw new Error('office_state_uncertain')
          latest = structuredClone(next)
        }
        const proposal = options.proposals.propose({
          operation: call.name,
          toolName: call.name,
          title:
            action === 'stage'
              ? '暂存修订页：插在原页之后，保留原页'
              : action === 'resume'
                ? '恢复已插入修订页的暂存回执'
                : action === 'commit'
                  ? '提交单页替换：删除原页并切换页面映射'
                  : action === 'undo'
                    ? '撤销单页替换：先恢复备份，再删除修订页'
                    : '撤回暂存修订页，保留原页',
          preview: {
            projectId,
            changeId,
            pageId: record.pageId,
            oldSlideId: record.oldSlideId,
            newSlideId: record.newSlideId,
            backupId: record.backupId,
            state: record.state,
            inspection: initial.status,
            originalRetained: action !== 'commit',
            businessMappingUpdated: action === 'commit' || action === 'undo',
          },
          impact: { host: 'powerpoint', targets: [record.oldSlideId], count: 1 },
          fingerprint: selectionFingerprint(JSON.stringify([record, initial])),
          validate: async (s) => {
            try {
              await stable(s)
              return true
            } catch {
              return false
            }
          },
          execute: async (s) => {
            await stable(s)
            if (action === 'stage') {
              await save(record, s)
              await options.adapter.stage(
                record,
                base64,
                async (newSlideId) => {
                  await save({ ...record, state: 'inserted', newSlideId }, s)
                },
                () => unchanged(s),
                s,
              )
              await unchanged(s)
              if (latest?.state !== 'inserted' || !latest.newSlideId)
                throw new Error('office_state_uncertain')
              const proof = await options.adapter.inspect(latest, s)
              await unchanged(s)
              if (proof.status !== 'staged') throw new Error('office_state_uncertain')
              await save({ ...latest, state: 'staged' }, s)
            } else if (action === 'resume') {
              await save({ ...record, state: 'staged' }, s)
            } else if (action === 'commit' || action === 'undo') {
              // Backup availability and bytes are checked again after confirmation.
              const backup = await verifyBackup(s)
              await unchanged(s)
              if (action === 'commit') {
                if (record.state === 'staged')
                  await save({ ...record, ...commitReceipts, state: 'commit_pending' }, s)
                await options.adapter.commit(latest!, () => unchanged(s), s)
                await unchanged(s)
                const proof = await options.adapter.inspect(latest!, s)
                await unchanged(s)
                if (proof.status !== 'applied') throw new Error('office_state_uncertain')
                await save({ ...latest!, state: 'applied' }, s)
              } else {
                if (record.state === 'applied') await save({ ...record, state: 'undo_pending' }, s)
                await options.adapter.undo(
                  latest!,
                  backup.base64,
                  async (restoredSlideId) => {
                    await save({ ...latest!, state: 'restore_inserted', restoredSlideId }, s)
                  },
                  () => unchanged(s),
                  s,
                )
                await unchanged(s)
                if (latest?.state !== 'restore_inserted') throw new Error('office_state_uncertain')
                const proof = await options.adapter.inspect(latest, s)
                await unchanged(s)
                if (proof.status !== 'undone') throw new Error('office_state_uncertain')
                await save({ ...latest, state: 'undone' }, s)
              }
            } else {
              if (record.state === 'staged') await save({ ...record, state: 'discard_pending' }, s)
              if (initial.status === 'staged')
                await options.adapter.discard(latest!, () => unchanged(s), s)
              await unchanged(s)
              const proof = await options.adapter.inspect(latest!, s)
              await unchanged(s)
              if (proof.status !== 'baseline') throw new Error('office_state_uncertain')
              await save({ ...latest!, state: 'discarded' }, s)
            }
          },
        })
        return {
          output: JSON.stringify({
            status: 'awaiting_confirmation',
            proposalId: proposal.id,
            changeId,
            originalRetained: action !== 'commit',
            businessMappingUpdated: false,
            ...(action === 'commit' || action === 'undo'
              ? {
                  afterConfirmation: {
                    prepareRequestId:
                      action === 'commit' ? record.requestId : record.parentRequestId,
                    tool: 'prepare_presentation_production_import',
                  },
                }
              : {}),
          }),
          mutated: false,
          summary:
            action === 'commit' || action === 'undo'
              ? '页面替换操作等待确认；完成后显式准备对应任务，再编辑或验收'
              : '页面暂存操作等待确认；原页保留，尚未完成替换',
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : ''
        return {
          output:
            /^(presentation_[a-z_]+|office_[a-z_]+|proposal_stale|invalid_tool_input|cancelled)$/.test(
              message,
            )
              ? message
              : 'presentation_page_replacement_failed',
          isError: true,
          mutated: false,
          summary: '页面替换操作未完成；请检查持久事务后再处理',
        }
      }
    },
  }
}
