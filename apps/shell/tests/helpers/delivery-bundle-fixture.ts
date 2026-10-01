import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import JSZip from 'jszip'
import { PresentationStore } from '@wiswork/project-store'
import { presentationDeliveryBundleFiles } from '@wiswork/project-store/presentation-delivery-bundle'
import { buildPresentationDeliveryReport } from '@wiswork/pptx-engine/presentation-delivery-report'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationDeliveryBundleService } from '../../src/main/presentation-delivery-bundles'
const roots: string[] = []
export function cleanupDeliveryBundleFixtures() {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
}
const hash = (v: string | Uint8Array) => createHash('sha256').update(v).digest('hex')
export async function deliveryBundleFixture(
  modify?: (files: Map<string, Buffer>) => void,
  configure?: (
    root: string,
    store: PresentationStore,
  ) => Omit<Parameters<typeof createPresentationDeliveryBundleService>[0], 'userDataPath'>,
) {
  const root = mkdtempSync(join(tmpdir(), 'delivery-bundle-'))
  roots.push(root)
  const store = new PresentationStore(root)
  const plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  const production = store.beginProduction(plan.projectId, 'doc', 'req', deck, {
    revision: 1,
    plan,
  })
  const report = await buildPresentationDeliveryReport({
    plan,
    deck,
    metadata: {
      projectId: plan.projectId,
      documentId: 'doc',
      requestId: 'req',
      planRevision: 1,
      inputDigest: production.inputDigest,
      planDigest: production.planDigest,
    },
    pageStates: production.pages.map((p) => ({ pageId: p.pageId, state: p.state })),
    reviews: [],
    issueLedger: store.issueActions(plan.projectId, 'doc', 'req'),
  })
  const files = new Map(
    presentationDeliveryBundleFiles.map(
      (name) =>
        [
          name,
          Buffer.from(
            name === 'presentation.pptx'
              ? 'PK\u0003\u0004host current bytes'
              : name === 'evidence.json'
                ? JSON.stringify(report)
                : name === 'claims.json'
                  ? JSON.stringify(plan.claims)
                  : name === 'sources.json'
                    ? JSON.stringify(plan.sources)
                    : 'historical not verified',
          ),
        ] as [string, Buffer],
    ),
  )
  modify?.(files)
  const manifest = {
    version: 1 as const,
    scope: 'current_office_document' as const,
    documentId: 'doc',
    projectId: plan.projectId,
    requestId: 'req',
    planRevision: 1,
    inputDigest: production.inputDigest,
    planDigest: production.planDigest,
    createdAt: new Date().toISOString(),
    files: [...files].map(([name, data]) => ({ name, sizeBytes: data.length, sha256: hash(data) })),
    checks: {
      completion: 'not_verified' as const,
      sourceAuthority: 'not_verified' as const,
      timeliness: 'not_verified' as const,
      roundTrip: 'not_run' as const,
      hostQa: 'not_checked' as const,
      pdf: 'not_requested' as const,
      ...([...files.keys()].some((name) => /^page-\d+\.png$/.test(name))
        ? { pageScreenshots: 'captured_unreviewed' as const }
        : {}),
    },
  }
  const zip = new JSZip()
  for (const [name, data] of files) zip.file(name, data)
  zip.file('manifest.json', JSON.stringify(manifest))
  const raw = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  const base = {
    documentId: 'doc',
    projectId: plan.projectId,
    requestId: 'req',
    bundleId: hash(raw),
  }
  const service = createPresentationDeliveryBundleService({
    userDataPath: root,
    ...configure?.(root, store),
  })
  const call = (operation: string, fields: Record<string, unknown> = {}) =>
    service(
      { ...base, operation: 'delivery_bundle_' + operation, ...fields },
      new AbortController().signal,
    )
  const begin = () => call('begin', { sha256: hash(raw), sizeBytes: raw.length, manifest })
  const upload = async () => {
    await begin()
    await call('chunk', { offset: 0, base64: raw.toString('base64') })
  }
  return { root, store, manifest, raw, call, begin, upload, base, service }
}
