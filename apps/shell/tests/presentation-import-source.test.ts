import { mkdtempSync, rmSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { expect, it } from 'vitest'
import { PresentationStore } from '@wiswork/project-store'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import { presentationArtifactContent } from '../../office-addin/src/skills/powerpoint/presentation-page-delivery.js'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationService } from '../src/main/presentation-service.js'

it('reads exact immutable source metadata for both request namespaces without transferring PPTX bodies', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-import-source-'))
  try {
    let service = createPresentationService({ userDataPath })
    const plan = benchmarkPlan(),
      deck = benchmarkPlannedDeck()
    const call = async (operation: string, fields: Record<string, unknown> = {}) =>
      JSON.parse(
        Buffer.from(
          await service(
            { operation, documentId: 'doc', projectId: plan.projectId, ...fields },
            new AbortController().signal,
          ),
        ).toString('utf8'),
      )
    await call('save_plan', { expectedRevision: 0, plan })
    const pending = await call('read_import_source', { source: 'production', requestId: 'same' })
    expect(pending).toEqual({ error: 'not_found' })
    const compiled = await call('compile', { requestId: 'same', planRevision: 1, deck })
    expect(compiled).not.toHaveProperty('error')
    await call('production_begin', { requestId: 'same', planRevision: 1, deck })
    expect(await call('read_import_source', { source: 'production', requestId: 'same' })).toEqual({
      error: 'page_not_ready',
    })
    await call('production_run', { requestId: 'same' })
    await call('save_plan', { expectedRevision: 1, plan: { ...plan, title: 'new current plan' } })
    service = createPresentationService({ userDataPath })
    const hash = (input: string) => createHash('sha256').update(input).digest('hex')
    const whole = await call('read_import_source', { source: 'compiled', requestId: 'same' })
    expect(whole).toEqual({
      version: 1,
      documentId: 'doc',
      projectId: plan.projectId,
      requestId: 'same',
      source: 'compiled',
      planRevision: 1,
      artifactDigest: hash(compiled.pptxBase64),
      pages: compiled.pages,
    })
    const source = await call('read_import_source', { source: 'production', requestId: 'same' })
    const production = new PresentationStore(userDataPath).production(
      plan.projectId,
      'doc',
      'same',
    )!
    const pages = deck.slides.map((page, index) => ({
      id: page.id,
      title: page.title,
      sourceSlideId: production.pages[index]!.result!.sourceSlideId,
    }))
    expect(source).toEqual({
      version: 1,
      documentId: 'doc',
      projectId: plan.projectId,
      requestId: 'same',
      source: 'production',
      planRevision: 1,
      artifactDigest: hash(
        presentationArtifactContent({
          documentId: 'doc',
          projectId: plan.projectId,
          requestId: 'same',
          planRevision: 1,
          slideCount: 8,
          pptxBase64: '',
          pages,
          pagePptxBase64: production.pages.map((page) => page.result!.pptxBase64),
        }),
      ),
      pages,
    })
    expect(JSON.stringify(source).length).toBeLessThan(10_000)
    expect(source).not.toHaveProperty('pptxBase64')
    expect(source).not.toHaveProperty('pagePptxBase64')
    expect(
      await call('read_import_source', {
        source: 'compiled',
        requestId: 'same',
        documentId: 'foreign',
      }),
    ).toEqual({ error: 'document_mismatch' })
    expect(await call('read_import_source', { source: 'invalid', requestId: 'same' })).toEqual({
      error: 'invalid_request',
    })
    const store = new PresentationStore(userDataPath)
    const legacy = store.begin(plan.projectId, 'doc', 'legacy', deck)
    store.complete(legacy, { pptxBase64: compiled.pptxBase64, report: compiled.report })
    expect(await call('read_import_source', { source: 'compiled', requestId: 'legacy' })).toEqual({
      version: 1,
      documentId: 'doc',
      projectId: plan.projectId,
      requestId: 'legacy',
      source: 'compiled',
      artifactDigest: hash(compiled.pptxBase64),
      pages: deck.slides.map(({ id, title }) => ({ id, title })),
    })
    const path = join(
      userDataPath,
      'projects',
      'presentations',
      hash(plan.projectId),
      `${hash('same')}.json`,
    )
    const corrupted = JSON.parse(readFileSync(path, 'utf8'))
    corrupted.result.pages[0].id = 'another-page'
    corrupted.resultDigest = hash(canonicalPresentationValue(corrupted.result))
    writeFileSync(path, JSON.stringify(corrupted))
    expect(await call('read_import_source', { source: 'compiled', requestId: 'same' })).toEqual({
      error: 'invalid_state',
    })
  } finally {
    rmSync(userDataPath, { recursive: true, force: true })
  }
})
