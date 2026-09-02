import type {
  ElevatedOfficeAdapter,
  ElevatedOfficeAuthority,
  ElevatedOfficeProgram,
  ElevatedOfficeSnapshot,
} from '../shared/elevated-office-program.js'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import type {
  PowerPointAdapter,
  PowerPointDeclarativeOperation,
} from './browser-powerpoint-adapter.js'
import {
  editPowerPointPackage,
  verifyPowerPointPackageInputs,
  type PackageEditResult,
} from './powerpoint-package.js'

export interface PowerPointElevatedAuthority {
  captureAuthority(): ElevatedOfficeAuthority
  snapshot(program: ElevatedOfficeProgram, signal?: AbortSignal): Promise<ElevatedOfficeSnapshot>
  validateSnapshot(
    program: ElevatedOfficeProgram,
    snapshot: ElevatedOfficeSnapshot,
    signal?: AbortSignal,
  ): Promise<boolean>
  execute(
    program: ElevatedOfficeProgram,
    snapshot: ElevatedOfficeSnapshot,
    signal?: AbortSignal,
    lifecycle?: Readonly<{ markStarted(): void; markApplied(): void }>,
  ): Promise<void>
  readback(
    program: ElevatedOfficeProgram,
    snapshot: ElevatedOfficeSnapshot,
    signal?: AbortSignal,
  ): Promise<{ verified: boolean; output?: unknown }>
  rollback(snapshot: ElevatedOfficeSnapshot, signal?: AbortSignal): Promise<void>
}

export const createPowerPointElevatedAdapter = (
  authority: PowerPointElevatedAuthority,
): ElevatedOfficeAdapter => Object.freeze({ host: 'powerpoint', ...authority })

type SlideState = { slideIndex: number; slideId: string; fingerprint: string; base64: string }
type PowerPointState = {
  slides: SlideState[]
  packageEdit?: PackageEditResult
  createdShapeIds?: string[]
}
const pptState = (snapshot: ElevatedOfficeSnapshot) => snapshot.state as PowerPointState
const slideIndexes = (program: ElevatedOfficeProgram) => [
  ...new Set(
    program.kind === 'office_js_ast'
      ? program.operations.map((operation) => Number(operation.args.slideIndex))
      : program.patches.map((patch) => Number(/slide(\d+)\.xml$/.exec(patch.part)![1]) - 1),
  ),
]
function pptOperations(program: ElevatedOfficeProgram): PowerPointDeclarativeOperation[] {
  if (program.kind !== 'office_js_ast') throw new Error('office_api_unsupported')
  return program.operations.map(({ call, args }) => {
    if (call === 'shape.setText')
      return {
        op: 'set_shape_text',
        slide_index: args.slideIndex as number,
        shape_id: args.shapeId as string,
        text: args.text as string,
      }
    if (call === 'shape.setGeometry')
      return {
        op: 'set_shape_geometry',
        slide_index: args.slideIndex as number,
        shape_id: args.shapeId as string,
        left: args.left as number,
        top: args.top as number,
        width: args.width as number,
        height: args.height as number,
        ...(args.style as {
          color: string
          fontFamily: string
          fontSize: number
          bold: boolean
          italic: boolean
        }),
      }
    if (call === 'slide.addTextBox')
      return {
        op: 'add_text_box',
        slide_index: args.slideIndex as number,
        name: 'WisWork Raw Text',
        text: args.text as string,
        left: args.left as number,
        top: args.top as number,
        width: args.width as number,
        height: args.height as number,
      }
    return {
      op: 'delete_shape',
      slide_index: args.slideIndex as number,
      shape_id: args.shapeId as string,
    }
  })
}

export function createBrowserPowerPointElevatedAdapter(options: {
  adapter: PowerPointAdapter
  authority(): ElevatedOfficeAuthority
}): ElevatedOfficeAdapter {
  return createPowerPointElevatedAdapter({
    captureAuthority: options.authority,
    async snapshot(program, signal) {
      const slides: SlideState[] = []
      for (const slideIndex of slideIndexes(program)) {
        const exported = await options.adapter.exportSlidePackage(slideIndex, signal)
        slides.push({ slideIndex, ...exported })
      }
      return {
        id: `history_${selectionFingerprint(JSON.stringify(slides.map(({ slideIndex, slideId, fingerprint }) => ({ slideIndex, slideId, fingerprint })))).replace(':', '_')}`,
        state: { slides } satisfies PowerPointState,
      }
    },
    async validateSnapshot(_program, snapshot, signal) {
      for (const slide of pptState(snapshot).slides) {
        const current = await options.adapter.snapshotSlide(slide.slideIndex, signal)
        if (current.slideId !== slide.slideId || current.fingerprint !== slide.fingerprint)
          return false
      }
      return true
    },
    async execute(program, snapshot, signal, lifecycle) {
      lifecycle?.markStarted()
      if (program.kind === 'office_js_ast') {
        const result = await options.adapter.executeDeclarative(pptOperations(program), signal)
        pptState(snapshot).createdShapeIds = result.createdShapeIds
        lifecycle?.markApplied()
        return
      }
      if (program.patches.length !== 1) throw new Error('office_api_unsupported')
      const slideIndex = slideIndexes(program)[0]
      const before = pptState(snapshot).slides.find((slide) => slide.slideIndex === slideIndex)
      if (!before) throw new Error('proposal_stale')
      const edit = await editPowerPointPackage(
        before.base64,
        'slide',
        [{ path: 'ppt/slides/slide1.xml', xml: program.patches[0].xml }],
        signal,
      )
      pptState(snapshot).packageEdit = edit
      await options.adapter.replaceSlidePackage(slideIndex, edit.base64, false, edit, signal)
      lifecycle?.markApplied()
    },
    async readback(program, snapshot, signal) {
      if (program.kind === 'ooxml_patch') {
        const edit = pptState(snapshot).packageEdit
        if (!edit) return { verified: false }
        const current = await options.adapter.exportSlidePackage(slideIndexes(program)[0], signal)
        return {
          verified: await verifyPowerPointPackageInputs(current.base64, edit.afterHashes, signal),
        }
      }
      let createdIndex = 0
      for (const operation of program.operations) {
        if (operation.call === 'shape.setText') {
          const current = await options.adapter.readSlideText(
            operation.args.slideIndex as number,
            operation.args.shapeId as string,
            signal,
          )
          if (current.text !== operation.args.text) return { verified: false }
        } else if (operation.call === 'slide.addTextBox') {
          const shapeId = pptState(snapshot).createdShapeIds?.[createdIndex++]
          if (!shapeId) return { verified: false }
          if (!options.adapter.readShapeTextStyle) return { verified: false }
          const [text, shapes, style] = await Promise.all([
            options.adapter.readSlideText(operation.args.slideIndex as number, shapeId, signal),
            options.adapter.listSlideShapes(operation.args.slideIndex as number, signal),
            options.adapter.readShapeTextStyle(
              operation.args.slideIndex as number,
              shapeId,
              signal,
            ),
          ])
          const shape = shapes.shapes.find((candidate) => candidate.id === shapeId)
          const expectedStyle = operation.args.style as Record<string, unknown>
          if (
            text.text !== operation.args.text ||
            !shape ||
            ['left', 'top', 'width', 'height'].some(
              (key) =>
                Math.abs(Number(shape[key as keyof typeof shape]) - Number(operation.args[key])) >
                0.01,
            ) ||
            style.color?.toUpperCase() !== String(expectedStyle.color).toUpperCase() ||
            style.fontFamily !== expectedStyle.fontFamily ||
            Math.abs(Number(style.fontSize) - Number(expectedStyle.fontSize)) > 0.01 ||
            style.bold !== expectedStyle.bold ||
            style.italic !== expectedStyle.italic
          )
            return { verified: false }
        } else if (operation.call === 'shape.setGeometry' || operation.call === 'shape.delete') {
          const current = await options.adapter.listSlideShapes(
            operation.args.slideIndex as number,
            signal,
          )
          const shape = current.shapes.find((candidate) => candidate.id === operation.args.shapeId)
          if (
            operation.call === 'shape.delete'
              ? shape
              : !shape ||
                ['left', 'top', 'width', 'height'].some(
                  (key) =>
                    Math.abs(
                      Number(shape[key as keyof typeof shape]) - Number(operation.args[key]),
                    ) > 0.01,
                )
          )
            return { verified: false }
        }
      }
      return { verified: true }
    },
    async rollback(snapshot, signal) {
      for (const slide of pptState(snapshot).slides)
        await options.adapter.replaceSlidePackage(
          slide.slideIndex,
          slide.base64,
          false,
          undefined,
          signal,
        )
    },
  })
}
