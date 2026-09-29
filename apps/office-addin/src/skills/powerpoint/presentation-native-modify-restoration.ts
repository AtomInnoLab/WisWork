import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import {
  validatePresentationExistingBatch,
  type PresentationNativeModifyBatch,
  type PresentationExistingBatch,
} from './presentation-existing-batch.js'
import {
  validatePresentationExistingPageChange,
  type PresentationExistingPageChange,
} from './presentation-existing-page.js'
import { readChartPackageBackup } from './presentation-chart-backup.js'
import { presentationPackageDigest } from './powerpoint-package.js'
interface Dependencies {
  documentId(): Promise<string>
  readExistingBatch(id: string): PresentationExistingBatch | undefined
  writeExistingBatch(
    next: PresentationExistingBatch,
    expected: PresentationExistingBatch | undefined,
  ): Promise<void>
  readExistingPageChange(id: string): PresentationExistingPageChange | undefined
  listExistingPageChanges(): PresentationExistingPageChange[]
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  exportPresentationPagePackage(
    slideId: string,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; slideIds: string[]; base64: string }>
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const conflict = (): never => {
  throw Error('presentation_native_modify_restore_conflict')
}
/** Proves separately committed original-page restorations before closing the mutation journal. No host writes. */
export function createPresentationNativeModifyRestoration(options: Dependencies) {
  let epoch = 0
  return {
    clear() {
      epoch++
    },
    async finalize(
      changeId: string,
      inputRestorationIds: string[],
      signal?: AbortSignal,
    ): Promise<PresentationNativeModifyBatch> {
      const restorationIds = structuredClone(inputRestorationIds)
      const token = epoch
      const abort = () => {
        if (signal?.aborted || token !== epoch) throw Error('cancelled')
      }
      abort()
      const documentId = await options.documentId()
      abort()
      const saved = options.readExistingBatch(changeId)
      if (
        !validatePresentationExistingBatch(saved) ||
        saved.version !== 3 ||
        saved.changeId !== changeId ||
        saved.documentId !== documentId ||
        saved.backupReleasedAt ||
        !Array.isArray(restorationIds) ||
        restorationIds.length !== saved.pages.length ||
        new Set(restorationIds).size !== restorationIds.length
      )
        return conflict()
      let record = structuredClone(saved)
      const candidates = () =>
        options
          .listExistingPageChanges()
          .filter(
            (page) =>
              page.documentId === documentId &&
              page.restores?.sourceKind === 'batch' &&
              page.restores.sourceChangeId === changeId &&
              page.state !== 'discarded',
          )
      const pages = restorationIds.map((id) => options.readExistingPageChange(id))
      if (pages.some((page) => !validatePresentationExistingPageChange(page))) return conflict()
      const restores = pages as PresentationExistingPageChange[]
      const active = candidates()
      if (
        active.length !== restores.length ||
        active.some((page) => !restores.some((expected) => same(page, expected)))
      )
        return conflict()
      const restored = Object.fromEntries(
        record.scope.slideIds.map((id) => [
          id,
          restores.find((page) => page.oldSlideId === id)?.newSlideId,
        ]),
      ) as Record<string, string>
      if (
        new Set(restores.map((page) => page.oldSlideId)).size !== restores.length ||
        new Set(Object.values(restored)).size !== restores.length ||
        record.scope.slideIds.some((id) => !restored[id]) ||
        Object.values(restored).some((id) => record.beforeSlideIds.includes(id))
      )
        return conflict()
      for (const page of restores) {
        const backup = record.backups.find((b) => b.hostSlideId === page.oldSlideId)
        if (
          !backup ||
          page.state !== 'applied' ||
          !page.newSlideId ||
          page.backupReleasedAt ||
          page.reapplies ||
          page.restores?.sourceHostSlideId !== page.oldSlideId ||
          page.restores.originalBackupId !== backup.backupId ||
          page.restores.originalPackageDigest !== backup.packageDigest ||
          page.replacementPackageDigest !== backup.packageDigest ||
          !page.sourceBackup ||
          page.sourceBackup.sha256 !== backup.sha256 ||
          page.sourceBackup.sizeBytes !== backup.sizeBytes
        )
          return conflict()
      }
      // Each committed restore replaces one identity in the preceding exact deck order.
      let order = [...record.beforeSlideIds]
      const remaining = [...restores]
      while (remaining.length) {
        const matches = remaining.filter((page) => same(page.beforeSlideIds, order))
        if (matches.length !== 1) return conflict()
        const page = matches[0]!
        order = order.map((id) => (id === page.oldSlideId ? page.newSlideId! : id))
        remaining.splice(remaining.indexOf(page), 1)
      }
      const guard = async () => {
        abort()
        if ((await options.documentId()) !== documentId) return conflict()
        abort()
        if (
          !same(options.readExistingBatch(changeId), record) ||
          restores.some((page) => !same(options.readExistingPageChange(page.changeId), page)) ||
          !same(candidates(), active)
        )
          throw Error('presentation_existing_batch_stale')
      }
      const request: Dependencies['request'] = async (body, requestSignal) => {
        await guard()
        const response = await options.request(body, requestSignal)
        await guard()
        return response
      }
      const prove = async () => {
        for (const page of restores) {
          const backup = record.backups.find((b) => b.hostSlideId === page.oldSlideId)!
          await readChartPackageBackup(
            {
              request,
              documentId,
              hostSlideId: page.oldSlideId,
              slideIds: record.beforeSlideIds,
              backup,
              expectedPackageDigest: backup.packageDigest,
            },
            signal,
          )
          await readChartPackageBackup(
            {
              request,
              documentId,
              hostSlideId: page.oldSlideId,
              slideIds: page.beforeSlideIds,
              backup: page.sourceBackup!,
              expectedPackageDigest: backup.packageDigest,
            },
            signal,
          )
          for (let i = 0; i < 2; i++) {
            await guard()
            const exported = await options.exportPresentationPagePackage(page.newSlideId!, signal)
            await guard()
            const digest = await presentationPackageDigest(exported.base64, signal)
            await guard()
            if (
              exported.slideId !== page.newSlideId ||
              !same(exported.slideIds, order) ||
              digest !== backup.packageDigest
            )
              return conflict()
          }
        }
      }
      await guard()
      if (record.state === 'undone') {
        if (!same(record.restoredSlideIds, restored)) return conflict()
        await prove()
        return structuredClone(record)
      }
      await prove()
      const store = async (next: PresentationNativeModifyBatch) => {
        await guard()
        await options.writeExistingBatch(next, record)
        record = next
        await guard()
      }
      if (record.state !== 'undoing') {
        const { inFlightIndex: _flight, ...identity } = record
        await store({ ...identity, state: 'undoing' })
      }
      await prove()
      await store({ ...record, state: 'undone', restoredSlideIds: restored })
      return structuredClone(record)
    },
  }
}

export function createPresentationNativeModifyRestorationSkill(
  options: Dependencies,
): AgentSkill & { clear(): void } {
  const controller = createPresentationNativeModifyRestoration(options)
  const tool: AgentToolDef = {
    name: 'finalize_native_modify_restore',
    description:
      'Close a generic edit savepoint only after every affected original page was separately restored and committed. Requires the exact applied restoration change IDs. Verifies original packages, restored pages and deck order; never writes the host.',
    inputSchema: {
      type: 'object',
      properties: {
        change_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
        restoration_change_ids: {
          type: 'array',
          minItems: 1,
          maxItems: 8,
          uniqueItems: true,
          items: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
        },
      },
      required: ['change_id', 'restoration_change_ids'],
      additionalProperties: false,
    },
  }
  return {
    id: 'presentation-native-modify-restoration',
    tools: [tool],
    systemPrompt:
      'For an uncertain generic write, prepare each original page restore with source_kind batch, stage and separately confirm commit. Finalize only after all affected originals have been restored; never replay an uncertain operation.',
    clear: () => controller.clear(),
    async executeTool(call, signal) {
      try {
        const input = structuredClone(call.input) as Record<string, unknown>
        if (
          call.inputError ||
          call.truncated ||
          call.name !== tool.name ||
          !input ||
          typeof input !== 'object' ||
          Array.isArray(input) ||
          Object.keys(input).sort().join(',') !== 'change_id,restoration_change_ids' ||
          typeof input.change_id !== 'string' ||
          !/^[A-Za-z0-9_-]{1,128}$/.test(input.change_id) ||
          !Array.isArray(input.restoration_change_ids) ||
          input.restoration_change_ids.length < 1 ||
          input.restoration_change_ids.length > 8 ||
          input.restoration_change_ids.some(
            (id) => typeof id !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(id),
          )
        )
          throw Error('invalid_tool_input')
        const record = await controller.finalize(
          input.change_id,
          input.restoration_change_ids as string[],
          signal,
        )
        return {
          output: JSON.stringify({
            changeId: record.changeId,
            state: record.state,
            restoredSlideIds: record.restoredSlideIds,
            historicalOnly: true,
            visualQaVerified: false,
          }),
          mutated: false,
          summary: 'Verified generic edit original-page restoration',
        }
      } catch (error) {
        return {
          output:
            error instanceof Error ? error.message : 'presentation_native_modify_restore_conflict',
          isError: true,
          mutated: false,
          summary: 'Generic restore verification failed',
        }
      }
    },
  }
}
