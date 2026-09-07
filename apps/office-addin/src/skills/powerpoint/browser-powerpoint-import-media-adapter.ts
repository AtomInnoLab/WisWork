import type { ImageGeometry, PowerPointImageAdapter } from './powerpoint-import-media.js'
import type { PowerPointAdapter } from './browser-powerpoint-adapter.js'

type Runtime = Record<string, any>
function cancelled(signal?: AbortSignal) {
  if (signal?.aborted) throw new Error('cancelled')
}
function runtime(): Runtime {
  const root = globalThis as Runtime
  const requirements = root.Office?.context?.requirements
  if (
    root.Office?.context?.host !== 'PowerPoint' ||
    !requirements?.isSetSupported?.('PowerPointApi', '1.5') ||
    !requirements?.isSetSupported?.('ImageCoercion', '1.1') ||
    typeof root.Office?.context?.document?.setSelectedDataAsync !== 'function' ||
    typeof root.PowerPoint?.run !== 'function'
  )
    throw new Error('office_api_unsupported')
  return root.PowerPoint
}
async function sync(context: Runtime, signal?: AbortSignal) {
  cancelled(signal)
  await context.sync()
  cancelled(signal)
}
async function slide(context: Runtime, index: number, signal?: AbortSignal): Promise<Runtime> {
  const slides = context.presentation.slides
  if (typeof slides.getItemAt !== 'function') throw new Error('office_api_unsupported')
  const item = slides.getItemAt(index)
  item.load('id')
  await sync(context, signal)
  return item
}
export function supportsPowerPointImportMedia(): boolean {
  try {
    runtime()
    return true
  } catch {
    return false
  }
}

export class BrowserPowerPointImportMediaAdapter implements PowerPointImageAdapter {
  constructor(private readonly base: Pick<PowerPointAdapter, 'snapshotSlide'>) {}
  snapshotSlide(index: number, signal?: AbortSignal) {
    return this.base.snapshotSlide(index, signal)
  }
  async insertImage(
    index: number,
    base64: string,
    geometry: ImageGeometry,
    signal?: AbortSignal,
  ): Promise<{ id: string }> {
    const powerpoint = runtime()
    const before = await powerpoint.run(async (context: Runtime) => {
      const item = await slide(context, index, signal)
      if (typeof context.presentation?.setSelectedSlides !== 'function')
        throw new Error('office_api_unsupported')
      item.shapes.load('items/id')
      await sync(context, signal)
      context.presentation.setSelectedSlides([String(item.id)])
      await sync(context, signal)
      return new Set((item.shapes.items as Runtime[]).map((shape) => String(shape.id)))
    })
    cancelled(signal)
    const root = globalThis as Runtime
    await new Promise<void>((resolve, reject) => {
      root.Office.context.document.setSelectedDataAsync(
        base64,
        {
          coercionType: root.Office.CoercionType.Image,
          imageLeft: geometry.left,
          imageTop: geometry.top,
          imageWidth: geometry.width,
          imageHeight: geometry.height,
        },
        (result: Runtime) => {
          if (result?.error || String(result?.status).toLowerCase() === 'failed')
            reject(result?.error ?? new Error('office_write_failed'))
          else resolve()
        },
      )
    })
    cancelled(signal)
    return powerpoint.run(async (context: Runtime) => {
      const item = await slide(context, index, signal)
      item.shapes.load('items/id')
      await sync(context, signal)
      const created = (item.shapes.items as Runtime[]).find(
        (shape) => !before.has(String(shape.id)),
      )
      if (!created?.id) throw new Error('office_write_failed')
      created.name = 'WisWork picture'
      await sync(context, signal)
      return { id: String(created.id) }
    })
  }
  async verifyImage(
    index: number,
    id: string,
    geometry: ImageGeometry,
    signal?: AbortSignal,
  ): Promise<boolean> {
    return runtime().run(async (context: Runtime) => {
      const item = await slide(context, index, signal)
      if (typeof item.shapes?.getItem !== 'function') throw new Error('office_api_unsupported')
      const shape = item.shapes.getItem(id)
      shape.load('id,left,top,width,height,type')
      if (typeof shape.fill?.load !== 'function') throw new Error('office_api_unsupported')
      shape.fill.load('type')
      await sync(context, signal)
      return (
        String(shape.id) === id &&
        String(shape.fill.type).toLowerCase().includes('picture') &&
        shape.left === geometry.left &&
        shape.top === geometry.top &&
        shape.width === geometry.width &&
        shape.height === geometry.height
      )
    })
  }
  async removeImage(index: number, id: string): Promise<void> {
    await runtime().run(async (context: Runtime) => {
      const item = await slide(context, index)
      const shape = item.shapes.getItem(id)
      if (typeof shape.delete !== 'function') throw new Error('office_api_unsupported')
      shape.delete()
      await sync(context)
    })
  }
  async verifyImageAbsent(index: number, id: string): Promise<boolean> {
    return runtime().run(async (context: Runtime) => {
      const item = await slide(context, index)
      item.shapes.load('items/id')
      await sync(context)
      return !(item.shapes.items as Runtime[]).some((shape) => String(shape.id) === id)
    })
  }
}
