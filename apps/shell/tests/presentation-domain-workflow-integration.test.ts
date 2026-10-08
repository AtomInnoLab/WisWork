import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import { PRESENTATION_DOMAIN_PROFILES } from '@wiswork/pptx-engine/presentation-plan'
import {
  parsePresentationDeliveryReport,
  presentationDeliveryMarkdown,
} from '@wiswork/pptx-engine/presentation-delivery-report'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'

it.each(['pitch', 'report', 'training', 'research', 'sales'] as const)(
  'keeps the %s workflow on its real frozen production after plan changes and PC restart',
  async (domain) => {
    const root = mkdtempSync(join(tmpdir(), 'wiswork-domain-workflow-'))
    try {
      let service = createPresentationService({ userDataPath: root })
      const plan = benchmarkPlan(),
        deck = benchmarkPlannedDeck()
      plan.domain = domain
      const profile = PRESENTATION_DOMAIN_PROFILES[domain]
      plan.slides.forEach((slide, index) => {
        slide.domainSection = profile.sections[index % profile.sections.length]
      })
      const call = async (operation: string, extra: Record<string, unknown> = {}) =>
        JSON.parse(
          Buffer.from(
            await service(
              { operation, documentId: 'doc', projectId: plan.projectId, ...extra },
              new AbortController().signal,
            ),
          ).toString('utf8'),
        )
      expect(await call('save_plan', { expectedRevision: 0, plan })).toMatchObject({ revision: 1 })
      expect(
        await call('production_begin', { requestId: 'frozen', planRevision: 1, deck }),
      ).not.toHaveProperty('error')
      const first = await call('production_delivery_report', { requestId: 'frozen' })
      expect(first).not.toHaveProperty('error')
      expect(first.domainWorkflow?.domain).toBe(domain)
      expect(first.domainWorkflow.sections.map((section: { id: string }) => section.id)).toEqual(
        profile.sections,
      )
      expect(first.checks).toMatchObject({
        content: 'needs_review',
        sourceAuthority: 'not_verified',
        host: 'not_checked',
      })
      expect(first).not.toHaveProperty('professionalWorkflow')
      expect(await call('production_run', { requestId: 'frozen' })).toMatchObject({
        status: 'compiled',
      })
      const compiled = await call('production_delivery_report', { requestId: 'frozen' })
      expect(
        compiled.pages.every(
          (page: { productionState: string }) => page.productionState === 'compiled',
        ),
      ).toBe(true)
      const changed = structuredClone(plan)
      delete changed.domain
      changed.slides.forEach((slide) => {
        delete slide.domainSection
      })
      expect(await call('save_plan', { expectedRevision: 1, plan: changed })).toMatchObject({
        revision: 2,
      })
      service = createPresentationService({ userDataPath: root })
      const restored = parsePresentationDeliveryReport(
        await call('production_delivery_report', { requestId: 'frozen' }),
      )
      expect(restored.planRevision).toBe(1)
      expect(restored.plan.domain).toBe(domain)
      expect(restored.domainWorkflow).toEqual(first.domainWorkflow)
      expect(restored.issueLedger.actions).toEqual([])
      const markdown = presentationDeliveryMarkdown(restored)
      expect(markdown).toContain('NOT VERIFIED')
      const readable = markdown.replace(/&#(\d+);/g, (_, code) =>
        String.fromCodePoint(Number(code)),
      )
      for (const section of first.domainWorkflow.sections)
        expect(readable).toContain(section.instruction)
      expect(restored.checks).toEqual(first.checks)
    } finally {
      rmSync(root, { recursive: true, force: true })
    }
  },
)
