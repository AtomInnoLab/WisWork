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
    !requirements?.isSetSupported?.('PowerPointApi', '1.8') ||
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
    return runtime().run(async (context: Runtime) => {
      const item = await slide(context, index, signal)
      if (typeof item.shapes?.addImage !== 'function') throw new Error('office_api_unsupported')
      cancelled(signal)
      const shape = item.shapes.addImage(base64, geometry)
      shape.load('id')
      await sync(context, signal)
      if (!shape.id) throw new Error('office_write_failed')
      return { id: String(shape.id) }
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
      await sync(context, signal)
      return (
        String(shape.id) === id &&
        String(shape.type).toLowerCase().includes('image') &&
        shape.left === geometry.left &&
        shape.top === geometry.top &&
        shape.width === geometry.width &&
        shape.height === geometry.height
      )
    })
  }
}
