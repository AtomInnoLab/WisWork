import { afterEach, describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { BrowserPresentationImageAdapter } from '../src/skills/powerpoint/browser-presentation-image-adapter'
const png =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII='
const next =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='
type Shape = {
  id: string
  type: string
  left: number
  top: number
  width: number
  height: number
  rotation: number
  name: string
  altTextTitle: string
  altTextDescription: string
  media: string
  readonly zOrderPosition: number
  load: ReturnType<typeof vi.fn>
  delete: ReturnType<typeof vi.fn>
  setZOrder: ReturnType<typeof vi.fn>
}
afterEach(() => vi.unstubAllGlobals())
async function packageFor(shapes: Shape[]) {
  const zip = new JSZip()
  let pics = '',
    rels = ''
  for (const shape of shapes) {
    const n = shape.id
    pics += `<p:pic><p:nvPicPr><p:cNvPr id="${n}" name="${shape.name}" descr="${shape.altTextDescription}" title="${shape.altTextTitle}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="r${n}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm rot="${shape.rotation * 60000}"><a:off x="${shape.left}" y="${shape.top}"/><a:ext cx="${shape.width}" cy="${shape.height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
    rels += `<Relationship Id="r${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${n}.png"/>`
    zip.file(`ppt/media/image${n}.png`, shape.media, { base64: true })
  }
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>${pics}</p:spTree></p:cSld></p:sld>`,
  )
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<Relationships>${rels}</Relationships>`)
  return zip.generateAsync({ type: 'base64' })
}
function setup() {
  const shapes: Shape[] = []
  let additions = 0,
    syncs = 0
  let onSync: (() => void) | undefined
  let setterFails = false
  let setterThrowSyncs = -1
  function create(id: string, media = png): Shape {
    const shape = {
      id,
      type: 'Image',
      left: 10,
      top: 20,
      width: 100,
      height: 80,
      rotation: 0,
      name: `Image${id}`,
      altTextTitle: 'title',
      altTextDescription: 'description',
      media,
      get zOrderPosition() {
        return shapes.indexOf(shape)
      },
      load: vi.fn(),
      delete: vi.fn(() => {
        shapes.splice(shapes.indexOf(shape), 1)
      }),
      setZOrder: vi.fn((action: string) => {
        const pos = shapes.indexOf(shape)
        shapes.splice(pos, 1)
        shapes.splice(action === 'SendToBack' ? 0 : Math.min(pos + 1, shapes.length), 0, shape)
      }),
    } as Shape
    return shape
  }
  shapes.push(create('1'), create('2'))
  const collection = {
    get items() {
      return shapes
    },
    load: vi.fn(),
    getItem: vi.fn((id: string) => {
      const shape = shapes.find((s) => s.id === id)
      if (!shape) throw new Error('ItemNotFound')
      return shape
    }),
    addImage: vi.fn((base64: string, geometry: Record<string, number>) => {
      additions++
      const shape = create('3', base64)
      Object.assign(shape, geometry)
      if (setterFails)
        Object.defineProperty(shape, 'name', {
          get: () => 'Image3',
          set: () => {
            setterThrowSyncs = syncs
            throw new Error('setter failed')
          },
        })
      shapes.push(shape)
      return shape
    }),
  }
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } },
  })
  vi.stubGlobal('PowerPoint', {
    run: async (callback: (context: unknown) => Promise<unknown>) => {
      const exported: { value: string }[] = []
      const page = {
        id: 'host-30',
        load: vi.fn(),
        shapes: collection,
        exportAsBase64: () => {
          const result = { value: '' }
          exported.push(result)
          return result
        },
      }
      return callback({
        presentation: {
          slides: {
            getItem: (id: string) => {
              if (id !== 'host-30') throw new Error('wrong ID')
              return page
            },
          },
        },
        sync: async () => {
          syncs++
          onSync?.()
          for (const item of exported) item.value = await packageFor(shapes)
        },
      })
    },
  })
  return {
    adapter: new BrowserPresentationImageAdapter(),
    shapes,
    collection,
    additions: () => additions,
    syncs: () => syncs,
    setterThrowSyncs: () => setterThrowSyncs,
    setOnSync: (fn: () => void) => {
      onSync = fn
    },
    failSetter: () => {
      setterFails = true
    },
  }
}
describe('ordinary image replacement transaction', () => {
  it('persists the verified candidate before deletion and preserves properties and object order', async () => {
    const { adapter, shapes, additions } = setup()
    shapes[0]!.rotation = 15
    const expected = await adapter.inspect('host-30', '1')
    const inserted = vi.fn(async (id: string) => {
      expect(id).toBe('3')
      expect(shapes.map((shape) => shape.id)).toEqual(['1', '3', '2'])
      expect(shapes[1]).toMatchObject({ name: 'Image1', rotation: 15, altTextTitle: 'title' })
    })
    expect(await adapter.replace('host-30', '1', next, expected, inserted)).toEqual({
      shapeId: '3',
    })
    expect(shapes.map((shape) => shape.id)).toEqual(['3', '2'])
    expect(additions()).toBe(1)
    expect(inserted).toHaveBeenCalledTimes(1)
  })
  it('rejects a stale snapshot without inserting', async () => {
    const { adapter, shapes, additions } = setup()
    const expected = await adapter.inspect('host-30', '1')
    shapes[0]!.left++
    await expect(adapter.replace('host-30', '1', next, expected, async () => {})).rejects.toThrow(
      'office_concurrent_change',
    )
    expect(additions()).toBe(0)
  })
  it('keeps the old image if candidate persistence fails or the user changes it during persistence', async () => {
    for (const mode of ['save-fails', 'user-edit']) {
      const { adapter, shapes, additions } = setup()
      const expected = await adapter.inspect('host-30', '1')
      await expect(
        adapter.replace('host-30', '1', next, expected, async () => {
          if (mode === 'save-fails') throw new Error('settings save failed')
          shapes.find((shape) => shape.id === '1')!.name = 'UserEdit'
        }),
      ).rejects.toThrow()
      expect(shapes.some((shape) => shape.id === '1')).toBe(true)
      expect(additions()).toBe(1)
    }
  })
  it('does not delete the original after cancellation following insertion', async () => {
    const { adapter, shapes, setOnSync, additions } = setup()
    const expected = await adapter.inspect('host-30', '1')
    const controller = new AbortController()
    setOnSync(() => {
      if (additions()) controller.abort()
    })
    await expect(
      adapter.replace('host-30', '1', next, expected, async () => {}, controller.signal),
    ).rejects.toThrow('cancelled')
    expect(shapes.map((shape) => shape.id)).toEqual(['1', '2', '3'])
    expect(additions()).toBe(1)
  })
  it('does not flush a metadata setter error or remove either image', async () => {
    const { adapter, shapes, failSetter, syncs, setterThrowSyncs } = setup()
    const expected = await adapter.inspect('host-30', '1')
    failSetter()
    await expect(adapter.replace('host-30', '1', next, expected, async () => {})).rejects.toThrow(
      'office_state_uncertain',
    )
    expect(syncs()).toBe(setterThrowSyncs())
    expect(shapes.some((shape) => shape.id === '1')).toBe(true)
    expect(shapes.some((shape) => shape.id === '3')).toBe(true)
  })
  it('never deletes the original if the inserted content differs or add sync rejects', async () => {
    for (const mode of ['wrong-media', 'rejected-sync']) {
      const { adapter, shapes, setOnSync, additions } = setup()
      const expected = await adapter.inspect('host-30', '1')
      setOnSync(() => {
        if (additions()) {
          if (mode === 'rejected-sync') throw new Error('rejected')
          shapes.find((shape) => shape.id === '3')!.media = png
        }
      })
      await expect(adapter.replace('host-30', '1', next, expected, async () => {})).rejects.toThrow(
        'office_state_uncertain',
      )
      expect(shapes.some((shape) => shape.id === '1')).toBe(true)
      expect(additions()).toBe(1)
    }
  })
  it('requires final readback even when deletion sync rejects after commit', async () => {
    const { adapter, shapes, setOnSync } = setup()
    const expected = await adapter.inspect('host-30', '1')
    let rejected = false
    setOnSync(() => {
      if (!shapes.some((shape) => shape.id === '1') && !rejected) {
        rejected = true
        throw new Error('delete sync rejected')
      }
    })
    await expect(adapter.replace('host-30', '1', next, expected, async () => {})).resolves.toEqual({
      shapeId: '3',
    })
    expect(rejected).toBe(true)
  })
  it('refuses success if deletion did not remove the original', async () => {
    const { adapter, shapes } = setup()
    const expected = await adapter.inspect('host-30', '1')
    shapes[0]!.delete.mockImplementation(() => {})
    await expect(adapter.replace('host-30', '1', next, expected, async () => {})).rejects.toThrow(
      'office_state_uncertain',
    )
    expect(shapes.some((shape) => shape.id === '1')).toBe(true)
  })
})

async function pendingRecovery() {
  const host = setup()
  const baseline = await host.adapter.inspect('host-30', '1')
  await expect(
    host.adapter.replace('host-30', '1', next, baseline, async () => {
      throw new Error('interrupted')
    }),
  ).rejects.toThrow('interrupted')
  const candidate = await host.adapter.inspect('host-30', '3')
  const record = {
    version: 1 as const,
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    pageId: 'page',
    hostSlideId: 'host-30',
    oldShapeId: '1',
    assetDigest: candidate.mediaDigest,
    state: 'pending' as const,
    newShapeId: '3',
    baseline,
  }
  return { ...host, record }
}
describe('interrupted image replacement recovery', () => {
  it('inspects both verified objects and completes by deleting only the old image', async () => {
    const { adapter, record, shapes, additions } = await pendingRecovery()
    expect(await adapter.inspectRecovery(record)).toEqual({ status: 'ready_to_finish' })
    await expect(adapter.finishRecovery(record, 'ready_to_finish')).resolves.toEqual({
      shapeId: '3',
    })
    expect(shapes.map((shape) => shape.id)).toEqual(['3', '2'])
    expect(additions()).toBe(1)
    expect(await adapter.inspectRecovery(record)).toEqual({ status: 'already_applied' })
    const candidate = shapes[0]!,
      calls = candidate.setZOrder.mock.calls.length
    await expect(adapter.finishRecovery(record, 'already_applied')).resolves.toEqual({
      shapeId: '3',
    })
    expect(candidate.setZOrder).toHaveBeenCalledTimes(calls)
    expect(additions()).toBe(1)
  })
  it.each(['missing', 'old-drift', 'candidate-drift', 'order-drift'] as const)(
    'requires manual review for %s',
    async (state) => {
      const { adapter, record, shapes, additions } = await pendingRecovery()
      if (state === 'missing')
        shapes.splice(
          shapes.findIndex((shape) => shape.id === '3'),
          1,
        )
      if (state === 'old-drift') shapes.find((shape) => shape.id === '1')!.left++
      if (state === 'candidate-drift') shapes.find((shape) => shape.id === '3')!.media = png
      if (state === 'order-drift') shapes.reverse()
      expect(await adapter.inspectRecovery(record)).toMatchObject({ status: 'manual_review' })
      await expect(adapter.finishRecovery(record, 'ready_to_finish')).rejects.toThrow(
        'office_concurrent_change',
      )
      expect(shapes.some((shape) => shape.id === '1')).toBe(true)
      expect(additions()).toBe(1)
    },
  )
  it('does not infer absence from host read errors or missing recovery evidence', async () => {
    const { adapter, record, setOnSync, shapes } = await pendingRecovery()
    expect(await adapter.inspectRecovery({ ...record, baseline: undefined })).toMatchObject({
      status: 'manual_review',
    })
    expect(await adapter.inspectRecovery({ ...record, newShapeId: undefined })).toMatchObject({
      status: 'manual_review',
    })
    expect(await adapter.inspectRecovery({ ...record, state: 'complete' })).toMatchObject({
      status: 'manual_review',
    })
    setOnSync(() => {
      throw new Error('read failed')
    })
    await expect(adapter.inspectRecovery(record)).rejects.toThrow('read failed')
    expect(shapes.some((shape) => shape.id === '1')).toBe(true)
  })
  it('rejects stale expected recovery state before deletion', async () => {
    const { adapter, record, shapes } = await pendingRecovery()
    await expect(adapter.finishRecovery(record, 'already_applied')).rejects.toThrow(
      'office_concurrent_change',
    )
    expect(shapes.some((shape) => shape.id === '1')).toBe(true)
  })
  it('does not flush a queued deletion when delete throws', async () => {
    const { adapter, record, shapes, syncs } = await pendingRecovery()
    let atThrow = -1
    shapes
      .find((shape) => shape.id === '1')!
      .delete.mockImplementation(() => {
        atThrow = syncs()
        throw new Error('queue failure')
      })
    await expect(adapter.finishRecovery(record, 'ready_to_finish')).rejects.toThrow(
      'office_state_uncertain',
    )
    expect(syncs()).toBe(atThrow)
    expect(shapes.some((shape) => shape.id === '1')).toBe(true)
  })
  it('cancels before writes and reconciles cancellation after a committed deletion', async () => {
    const { adapter, record, shapes, setOnSync } = await pendingRecovery()
    const before = new AbortController()
    before.abort()
    await expect(adapter.finishRecovery(record, 'ready_to_finish', before.signal)).rejects.toThrow(
      'cancelled',
    )
    expect(shapes.some((shape) => shape.id === '1')).toBe(true)
    const after = new AbortController()
    setOnSync(() => {
      if (!shapes.some((shape) => shape.id === '1')) after.abort()
    })
    await expect(adapter.finishRecovery(record, 'ready_to_finish', after.signal)).resolves.toEqual({
      shapeId: '3',
    })
  })
  it('treats unsupported picture content as manual review but propagates unsupported host reads', async () => {
    const { adapter, record, shapes } = await pendingRecovery()
    shapes.find((shape) => shape.id === '3')!.type = 'TextBox'
    expect(await adapter.inspectRecovery(record)).toMatchObject({
      status: 'manual_review',
      reason: 'unsupported_picture',
    })
    shapes.find((shape) => shape.id === '3')!.type = 'Image'
    vi.spyOn(adapter, 'inspect').mockRejectedValueOnce(new Error('office_api_unsupported'))
    await expect(adapter.inspectRecovery(record)).rejects.toThrow('office_api_unsupported')
  })
  it('requires proof of completed deletion and does not convert final read failures to success', async () => {
    for (const mode of ['not-deleted', 'read-failed']) {
      const { adapter, record, shapes, setOnSync } = await pendingRecovery()
      if (mode === 'not-deleted')
        shapes.find((shape) => shape.id === '1')!.delete.mockImplementation(() => {})
      else
        setOnSync(() => {
          if (!shapes.some((shape) => shape.id === '1')) throw new Error('read failed')
        })
      await expect(adapter.finishRecovery(record, 'ready_to_finish')).rejects.toThrow(
        'office_state_uncertain',
      )
    }
  })
})
