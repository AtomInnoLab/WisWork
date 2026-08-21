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
      item.shapes.load('items/id')
      await sync(context, signal)
      const beforeIds = new Set((item.shapes.items as Runtime[]).map((shape) => String(shape.id)))
      cancelled(signal)
      let created: Runtime | undefined
      try {
        created = item.shapes.addImage(base64, geometry)
        if (typeof created?.delete !== 'function') throw new Error('office_api_unsupported')
        created.load('id')
        await sync(context, signal)
        if (!created.id) throw new Error('office_write_failed')
        return { id: String(created.id) }
      } catch (writeError) {
        try {
          if (created) {
            created.delete()
            await sync(context)
          }
          item.shapes.load('items/id')
          await sync(context)
          const recovered = new Set(
            (item.shapes.items as Runtime[]).map((shape) => String(shape.id)),
          )
          if (recovered.size !== beforeIds.size || [...recovered].some((id) => !beforeIds.has(id)))
            throw new Error('office_recovery_failed')
        } catch (recoveryError) {
          throw new Error('office_recovery_failed', {
            cause: new Error(
              recoveryError instanceof Error && recoveryError.message === 'office_recovery_failed'
                ? 'office_recovery_failed'
                : 'office_host_error',
            ),
          })
        }
        throw new Error('office_write_failed', {
          cause: new Error(
            writeError instanceof Error && writeError.message === 'cancelled'
              ? 'cancelled'
              : 'office_host_error',
          ),
        })
      }
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
