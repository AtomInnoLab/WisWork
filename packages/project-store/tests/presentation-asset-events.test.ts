import { parsePresentationAssetLedger } from '../src/presentation-asset-events.js'
import { createHash } from 'node:crypto'
import { mkdtempSync, rmSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import { PresentationStore } from '../src/presentation-store.js'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../pptx-engine/tests/fixtures/presentation-plan.js'
it('persists exact asset resolution attempts and restores their identity without permitting events for unrelated pages', () => {
  const root = mkdtempSync(join(tmpdir(), 'asset-events-'))
  try {
    const store = new PresentationStore(root),
      plan = benchmarkPlan(),
      deck = benchmarkPlannedDeck()
    const saved = store.beginProduction(plan.projectId, 'doc', 'run', deck, { revision: 1, plan })
    const page = deck.slides.find((page) =>
      page.elements.some((element) => element.kind === 'image'),
    )!
    const assetId = page.elements.find((element) => element.kind === 'image')!.assetId
    expect(store.productionAssets(plan.projectId, 'doc', 'run').events).toEqual([])
    expect(() =>
      store.appendProductionAsset(plan.projectId, 'doc', 'run', {
        type: 'asset.fetching',
        pageId: page.id,
        assetId,
        attempt: 1,
      }),
    ).toThrow('invalid_state')
    store.updateProductionPage(saved, page.id, { state: 'building', attempt: 1 })
    store.appendProductionAsset(plan.projectId, 'doc', 'run', {
      type: 'asset.fetching',
      pageId: page.id,
      assetId,
      attempt: 1,
    })
    store.appendProductionAsset(plan.projectId, 'doc', 'run', {
      type: 'asset.ready',
      pageId: page.id,
      assetId,
      attempt: 1,
    })
    const ledger = new PresentationStore(root).productionAssets(plan.projectId, 'doc', 'run')
    expect(ledger).toMatchObject({
      scope: 'production_asset_resolution',
      requestId: 'run',
      inputDigest: saved.inputDigest,
      planDigest: saved.planDigest,
      revision: 2,
    })
    expect(ledger.events.map((event) => event.type)).toEqual(['asset.fetching', 'asset.ready'])
    expect(() =>
      store.appendProductionAsset(plan.projectId, 'doc', 'run', {
        type: 'asset.ready',
        pageId: page.id,
        assetId,
        attempt: 1,
      }),
    ).toThrow('invalid_state')
    expect(() =>
      store.appendProductionAsset(plan.projectId, 'doc', 'run', {
        type: 'asset.fetching',
        pageId: deck.slides[0]!.id,
        assetId,
        attempt: 1,
      }),
    ).toThrow('invalid_state')
    expect(() => store.productionAssets(plan.projectId, 'foreign', 'run')).toThrow(
      'document_mismatch',
    )
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
it('keeps a bounded replay window over retries and rejects corrupted persisted events', () => {
  const root = mkdtempSync(join(tmpdir(), 'asset-events-'))
  try {
    const store = new PresentationStore(root),
      plan = benchmarkPlan(),
      deck = benchmarkPlannedDeck()
    let record = store.beginProduction(plan.projectId, 'doc', 'run', deck, { revision: 1, plan })
    const page = deck.slides.find((page) =>
      page.elements.some((element) => element.kind === 'image'),
    )!
    const assetId = page.elements.find((element) => element.kind === 'image')!.assetId
    for (let attempt = 1; attempt <= 65; attempt++) {
      record = store.updateProductionPage(record, page.id, { state: 'building', attempt })
      store.appendProductionAsset(plan.projectId, 'doc', 'run', {
        type: 'asset.fetching',
        pageId: page.id,
        assetId,
        attempt,
      })
      store.appendProductionAsset(plan.projectId, 'doc', 'run', {
        type: 'asset.rejected',
        pageId: page.id,
        assetId,
        attempt,
        error: 'asset_unavailable',
      })
      record = store.updateProductionPage(record, page.id, {
        state: 'failed',
        attempt,
        error: 'asset_unavailable',
      })
    }
    const ledger = new PresentationStore(root).productionAssets(plan.projectId, 'doc', 'run')
    expect(ledger).toMatchObject({ revision: 130 })
    expect(ledger.events).toHaveLength(128)
    expect(ledger.events[0]!.sequence).toBe(3)
    expect(ledger.events.at(-1)!.attempt).toBe(65)
    const file = join(
      root,
      'projects',
      'presentations',
      createHash('sha256').update(plan.projectId).digest('hex'),
      `asset-events-${createHash('sha256').update('run').digest('hex')}.json`,
    )
    const stored = JSON.parse(readFileSync(file, 'utf8'))
    stored.events.at(-1).error = 'aborted'
    writeFileSync(file, JSON.stringify(stored))
    expect(() => store.productionAssets(plan.projectId, 'doc', 'run')).toThrow('invalid_state')
    expect(
      store
        .production(plan.projectId, 'doc', 'run')!
        .pages.find((value) => value.pageId === page.id)!.attempt,
    ).toBe(65)
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
it('rejects decreasing attempts and unexpected private fields in public history', () => {
  const base = {
    version: 1,
    scope: 'production_asset_resolution',
    projectId: 'project',
    documentId: 'doc',
    requestId: 'run',
    inputDigest: 'a'.repeat(64),
    planDigest: 'b'.repeat(64),
    revision: 3,
    events: [
      {
        type: 'asset.fetching',
        pageId: 'page',
        assetId: 'image',
        attempt: 2,
        sequence: 1,
        createdAt: '2026-09-29T00:00:00.000Z',
      },
      {
        type: 'asset.ready',
        pageId: 'page',
        assetId: 'image',
        attempt: 2,
        sequence: 2,
        createdAt: '2026-09-29T00:00:01.000Z',
      },
      {
        type: 'asset.fetching',
        pageId: 'page',
        assetId: 'image',
        attempt: 1,
        sequence: 3,
        createdAt: '2026-09-29T00:00:02.000Z',
      },
    ],
  }
  expect(() => parsePresentationAssetLedger(base)).toThrow('invalid_state')
  base.events[2]!.attempt = 3
  expect(parsePresentationAssetLedger(base).revision).toBe(3)
  expect(() => parsePresentationAssetLedger({ ...base, base64: 'private' })).toThrow(
    'invalid_state',
  )
})
