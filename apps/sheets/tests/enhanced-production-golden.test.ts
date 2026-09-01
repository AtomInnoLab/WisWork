import { describe, expect, it } from 'vitest'
import { runEnhancedGolden } from '../../../packages/agent-runtime/src/production-golden'
import { createHostGoldenBridge } from '../../../packages/agent-runtime/tests/host-golden-bridge'
import { InMemoryWorkbookAdapter } from '../src/domain/in-memory-workbook'
import type { WorkbookSnapshot } from '../src/domain/workbook.types'
import { createWorkbookSkill } from '../src/renderer/ai/workbook-skill'

describe('Sheets production Enhanced golden', () => {
  it('runs the production workbook skill through apply, readback receipt and undo', async () => {
    const adapter = new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [{ id: 'sheet-1', name: 'Sheet1', cells: { A1: { value: 'before' } } }],
    } satisfies WorkbookSnapshot)
    const skill = createWorkbookSkill({
      getActiveSheetInfo: () => ({
        mode: 'demo',
        sheetId: 'sheet-1',
        sheetName: 'Sheet1',
        revision: adapter.getSnapshot().revision,
        knownAddresses: ['A1'],
        sheets: [{ id: 'sheet-1', name: 'Sheet1' }],
      }),
      readCells: (addresses) =>
        Object.fromEntries(
          addresses.map((address) => [
            address,
            { value: adapter.getSnapshot().sheets[0]?.cells[address]?.value ?? null },
          ]),
        ),
      readFormats: () => ({}),
      readSheetFeatures: () => '',
      proposeOperations: (operations, summary) => {
        const plan = adapter.plan({
          dslVersion: 1,
          transactionId: 'enhanced-sheets-golden',
          baseRevision: adapter.getSnapshot().revision,
          summary,
          operations: [...operations],
        })
        adapter.apply(plan)
        return { ok: true, plan }
      },
    })
    const call = {
      id: 'sheets-golden-call',
      name: 'propose_operations',
      input: {
        operations: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'after' }],
        summary: 'Update A1',
      },
    }
    const result = await runEnhancedGolden('sheets', {
      documentId: 'sheets-golden-document',
      generation: 1,
      instruction: 'Update A1',
      bridge: createHostGoldenBridge({
        documentId: 'sheets-golden-document',
        generation: 1,
        call,
      }),
      captureSnapshot: () => adapter.getSnapshot(),
      skill,
      confirm: async () => ({ mutationReceiptId: 'enhanced-sheets-golden' }),
      readback: async () => ({
        status:
          adapter.getSnapshot().sheets[0]?.cells.A1?.value === 'after' ? 'verified' : 'failed',
      }),
      rollback: async () => {
        adapter.undo()
        return { status: 'restored' as const }
      },
    })
    expect(result.verification).toEqual({ status: 'verified' })
    expect(result.rollback).toEqual({ status: 'restored' })
    expect(adapter.getSnapshot().sheets[0]?.cells.A1?.value).toBe('before')
    console.log(
      'ENHANCED_GOLDEN_REPORT',
      JSON.stringify({
        host: 'sheets',
        verification: result.verification.status,
        rollback: result.rollback.status,
      }),
    )
  })
})
