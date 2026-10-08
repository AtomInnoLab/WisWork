import { afterAll, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { masterXmlFixture, cleanupMasterXmlFixtures } from './helpers/master-xml-fixture.js'
import type { StructuredProposalRequest } from '../src/agent/proposal-controller.js'

describe('durable master XML complete proposal preimage', () => {
  let f: Awaited<ReturnType<typeof masterXmlFixture>>, request: StructuredProposalRequest
  let originalOrder: string[]
  beforeAll(async () => {
    f = await masterXmlFixture(600)
    originalOrder = [...f.order]
    const propose = f.proposals.propose.bind(f.proposals)
    vi.spyOn(f.proposals, 'propose').mockImplementation((value) => {
      request = value
      return propose(value)
    })
    await f.propose()
  }, 120_000)
  afterAll(cleanupMasterXmlFixtures)
  beforeEach(() => {
    f.order.splice(0, f.order.length, ...originalOrder)
    for (const id of originalOrder) f.packages.set(id, f.original)
    f.adapter.stage.mockClear()
    f.adapter.remove.mockClear()
    f.adapter.applyLayout.mockClear()
  })
  const noWrites = () => {
    expect(f.adapter.stage).not.toHaveBeenCalled()
    expect(f.adapter.remove).not.toHaveBeenCalled()
    expect(f.adapter.applyLayout).not.toHaveBeenCalled()
    expect(f.data.size).toBe(0)
  }
  it('checks all 600 pages before confirmation without uploading savepoints', async () => {
    expect(await request.validate()).toBe(true)
    expect(f.order).toEqual(Array.from({ length: 600 }, (_, i) => `s${i}`))
    expect(f.adapter.readPage.mock.calls.map(([id]) => id)).toEqual(originalOrder)
    const completed = f.request.mock.calls
      .map(([body]) => body as Record<string, unknown>)
      .filter((body) => body.operation === 'package_backup_finish')
    expect(completed).toHaveLength(0)
    expect(request.impact).toMatchObject({ count: 600, targets: originalOrder })
    noWrites()
  })
  it.each(['identity', 'unrelated', 'structural', 'order'] as const)(
    'refuses %s drift at the last page before any SDK mutation',
    async (drift) => {
      if (drift === 'identity') {
        f.order[599] = 'foreign-page'
        f.packages.set('foreign-page', f.original)
        f.deps.set('foreign-page', { slideId: 'foreign-page', masterId: 'm1', layoutId: 'l1' })
      } else if (drift === 'order') {
        ;[f.order[598], f.order[599]] = [f.order[599]!, f.order[598]!]
      } else {
        const zip = await JSZip.loadAsync(f.original, { base64: true })
        zip.file(
          drift === 'unrelated' ? 'docProps/core.xml' : 'ppt/slides/_rels/slide1.xml.rels',
          drift === 'unrelated'
            ? '<core value="foreign"/>'
            : '<Relationships><Relationship Id="foreign" Target="other.xml"/></Relationships>',
        )
        f.packages.set('s599', await zip.generateAsync({ type: 'base64' }))
      }
      expect(await request.validate()).toBe(false)
      await expect(request.execute()).rejects.toThrow('drift')
      noWrites()
    },
  )
  it('freezes host observations before later document guards mutate returned aliases', async () => {
    const inspect = f.adapter.inspect.getMockImplementation()!
    f.adapter.inspect.mockImplementationOnce(async (ids) => {
      const observed = await inspect(ids)
      f.readHook(() => {
        f.order[599] = 'foreign-page'
        f.packages.set('foreign-page', f.original)
        f.deps.set('foreign-page', { slideId: 'foreign-page', masterId: 'm1', layoutId: 'l1' })
        observed.slideIds[599] = 'foreign-page'
        observed.pages[599]!.slideId = 'foreign-page'
        observed.dependencies[599]!.slideId = 'foreign-page'
        f.readHook(() => {})
      })
      return observed
    })
    expect(await request.validate()).toBe(true)
    await expect(request.execute()).rejects.toThrow('drift')
    noWrites()
  })
})
