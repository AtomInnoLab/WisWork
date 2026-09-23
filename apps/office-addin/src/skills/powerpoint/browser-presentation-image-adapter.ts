import { inspectPowerPointPicturePackage } from './powerpoint-package.js'
import type { PresentationPageGeometry } from './browser-powerpoint-adapter.js'

type Runtime = Record<string, any>
export interface PictureSnapshot {
  slideId: string
  shapeId: string
  geometry: PresentationPageGeometry
  rotation: number
  name: string
  altTextTitle: string
  altTextDescription: string
  zOrderPosition: number
  shapeIds: string[]
  pictureFingerprint: string
  mediaDigest: string
}
function check(signal?: AbortSignal) {
  if (signal?.aborted) throw new Error('cancelled')
}
function id(value: unknown): asserts value is string {
  if (typeof value !== 'string' || !value.length || value.length > 256)
    throw new Error('invalid_tool_input')
}
function runtime(): Runtime {
  const root = globalThis as Runtime
  if (
    root.Office?.context?.host !== 'PowerPoint' ||
    !root.Office?.context?.requirements?.isSetSupported?.('PowerPointApi', '1.10') ||
    typeof root.PowerPoint?.run !== 'function'
  )
    throw new Error('office_api_unsupported')
  return root.PowerPoint
}
const near = (a: number, b: number) => Math.abs(a - b) <= 0.01
function placement(a: PictureSnapshot, b: PictureSnapshot) {
  return (
    (['left', 'top', 'width', 'height'] as const).every((key) =>
      near(a.geometry[key], b.geometry[key]),
    ) &&
    near(a.rotation, b.rotation) &&
    a.name === b.name &&
    a.altTextTitle === b.altTextTitle &&
    a.altTextDescription === b.altTextDescription
  )
}
function same(a: unknown, b: unknown) {
  return JSON.stringify(a) === JSON.stringify(b)
}
async function slide(context: Runtime, slideId: string, signal?: AbortSignal): Promise<Runtime> {
  const item = context.presentation?.slides?.getItem?.(slideId)
  if (typeof item?.load !== 'function') throw new Error('office_api_unsupported')
  item.load('id')
  await context.sync()
  check(signal)
  if (item.id !== slideId) throw new Error('office_read_failed')
  return item
}
export class BrowserPresentationImageAdapter {
  async inspect(slideId: string, shapeId: string, signal?: AbortSignal): Promise<PictureSnapshot> {
    check(signal)
    id(slideId)
    id(shapeId)
    return runtime().run(async (context: Runtime) => {
      const page = await slide(context, slideId, signal)
      if (typeof page.shapes?.load !== 'function' || typeof page.exportAsBase64 !== 'function')
        throw new Error('office_api_unsupported')
      page.shapes.load({
        $top: 101,
        id: true,
        type: true,
        left: true,
        top: true,
        width: true,
        height: true,
        rotation: true,
        name: true,
        altTextTitle: true,
        altTextDescription: true,
        zOrderPosition: true,
      })
      const exported = page.exportAsBase64()
      await context.sync()
      check(signal)
      const shapes = page.shapes.items as Runtime[]
      if (!Array.isArray(shapes) || shapes.length > 100) throw new Error('office_read_failed')
      const ordered = [...shapes].sort((a, b) => a.zOrderPosition - b.zOrderPosition)
      if (
        ordered.some(
          (shape, i) =>
            typeof shape.id !== 'string' ||
            !shape.id.length ||
            shape.id.length > 256 ||
            shape.zOrderPosition !== i,
        ) ||
        new Set(ordered.map((shape) => shape.id)).size !== ordered.length
      )
        throw new Error('office_read_failed')
      const shape = ordered.find((value) => value.id === shapeId)
      if (!shape || shape.type !== 'Image') throw new Error('office_api_unsupported')
      const geometry = {
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
      }
      if (
        Object.values(geometry).some(
          (value) =>
            typeof value !== 'number' || !Number.isFinite(value) || Math.abs(value) > 100000,
        ) ||
        geometry.width < 0 ||
        geometry.height < 0 ||
        typeof shape.rotation !== 'number' ||
        !Number.isFinite(shape.rotation) ||
        Math.abs(shape.rotation) > 360 ||
        ['name', 'altTextTitle', 'altTextDescription'].some(
          (key) => typeof shape[key] !== 'string' || shape[key].length > 12000,
        )
      )
        throw new Error('office_read_failed')
      const proof = await inspectPowerPointPicturePackage(exported.value, shapeId, signal)
      if (
        !same(
          proof.shapeIds,
          ordered.map((value) => value.id),
        )
      )
        throw new Error('office_read_failed')
      return {
        slideId,
        shapeId,
        geometry,
        rotation: shape.rotation,
        name: shape.name,
        altTextTitle: shape.altTextTitle,
        altTextDescription: shape.altTextDescription,
        zOrderPosition: shape.zOrderPosition,
        ...proof,
      }
    })
  }

  async replace(
    slideId: string,
    shapeId: string,
    base64: string,
    expected: PictureSnapshot,
    onInserted: (newId: string) => Promise<void>,
    signal?: AbortSignal,
  ): Promise<{ shapeId: string }> {
    check(signal)
    id(slideId)
    id(shapeId)
    if (typeof base64 !== 'string' || base64.length > Math.ceil((2 * 1024 * 1024) / 3) * 4)
      throw new Error('invalid_tool_input')
    let binary: string
    try {
      binary = atob(base64)
    } catch {
      throw new Error('invalid_tool_input')
    }
    if (
      !binary.length ||
      binary.length > 2 * 1024 * 1024 ||
      btoa(binary) !== base64 ||
      !(binary.startsWith('\x89PNG\r\n\x1a\n') || binary.startsWith('\xff\xd8'))
    )
      throw new Error('invalid_tool_input')
    const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0))
    const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)))
      .map((byte) => byte.toString(16).padStart(2, '0'))
      .join('')
    const current = await this.inspect(slideId, shapeId, signal)
    if (!same(current, expected) || current.shapeIds.length >= 100)
      throw new Error('office_concurrent_change')
    check(signal)
    // Never infer ownership by diffing shape collections. Only the addImage return proxy owns it.
    const newId: string = await runtime().run(async (context: Runtime) => {
      const page = await slide(context, slideId, signal)
      if (typeof page.shapes?.addImage !== 'function') throw new Error('office_api_unsupported')
      let created: Runtime
      try {
        created = page.shapes.addImage(base64, current.geometry)
        created.load('id')
      } catch {
        throw new Error('office_state_uncertain')
      }
      if (signal?.aborted) throw new Error('office_state_uncertain')
      try {
        await context.sync()
      } catch {
        throw new Error('office_state_uncertain')
      }
      if (
        typeof created.id !== 'string' ||
        !created.id.length ||
        created.id.length > 256 ||
        current.shapeIds.includes(created.id)
      )
        throw new Error('office_state_uncertain')
      return created.id
    })
    check(signal)
    const candidate = await this.inspect(slideId, newId, signal)
    const old = await this.inspect(slideId, shapeId, signal)
    const appended = [...current.shapeIds, newId]
    if (
      candidate.mediaDigest !== digest ||
      !same(candidate.shapeIds, appended) ||
      !same(old.shapeIds, appended) ||
      old.pictureFingerprint !== current.pictureFingerprint ||
      !placement(old, current) ||
      (Object.keys(current.geometry) as Array<keyof PresentationPageGeometry>).some(
        (key) => !near(candidate.geometry[key], current.geometry[key]),
      )
    )
      throw new Error('office_state_uncertain')
    check(signal)
    await runtime().run(async (context: Runtime) => {
      const page = await slide(context, slideId, signal)
      const image = page.shapes.getItem(newId)
      if (typeof image?.setZOrder !== 'function') throw new Error('office_api_unsupported')
      try {
        image.rotation = current.rotation
        image.name = current.name
        image.altTextTitle = current.altTextTitle
        image.altTextDescription = current.altTextDescription
        image.setZOrder('SendToBack')
        for (let i = 0; i <= current.zOrderPosition; i++) image.setZOrder('BringForward')
      } catch {
        throw new Error('office_state_uncertain')
      }
      if (signal?.aborted) throw new Error('office_state_uncertain')
      try {
        await context.sync()
      } catch {
        throw new Error('office_state_uncertain')
      }
    })
    check(signal)
    const order = [...current.shapeIds]
    order.splice(current.zOrderPosition + 1, 0, newId)
    const verifyPair = async () => {
      const freshOld = await this.inspect(slideId, shapeId, signal)
      const freshNew = await this.inspect(slideId, newId, signal)
      if (
        freshOld.pictureFingerprint !== current.pictureFingerprint ||
        !placement(freshOld, current) ||
        freshNew.mediaDigest !== digest ||
        !placement(freshNew, current) ||
        !same(freshOld.shapeIds, order) ||
        !same(freshNew.shapeIds, order)
      )
        throw new Error('office_state_uncertain')
    }
    await verifyPair()
    check(signal)
    await onInserted(newId)
    check(signal)
    await verifyPair()
    check(signal)
    // The candidate is verified and durably identified. No cancellation after insertion may
    // cross this deletion boundary; errors leave both images for explicit user inspection.
    await runtime().run(async (context: Runtime) => {
      const page = await slide(context, slideId, signal)
      const original = page.shapes.getItem(shapeId)
      if (typeof original?.delete !== 'function') throw new Error('office_api_unsupported')
      check(signal)
      try {
        original.delete()
      } catch {
        throw new Error('office_state_uncertain')
      }
      if (signal?.aborted) throw new Error('office_state_uncertain')
      try {
        await context.sync()
      } catch {
        /* Reconcile using a fresh read-only context below. */
      }
    })
    let final: PictureSnapshot
    try {
      final = await this.inspect(slideId, newId)
    } catch {
      throw new Error('office_state_uncertain')
    }
    const finalOrder = current.shapeIds.map((value) => (value === shapeId ? newId : value))
    if (
      final.mediaDigest !== digest ||
      !placement(final, current) ||
      final.zOrderPosition !== current.zOrderPosition ||
      !same(final.shapeIds, finalOrder)
    )
      throw new Error('office_state_uncertain')
    return { shapeId: newId }
  }
}
