import {
  presentationSlideSourceLabels,
  type PresentationClaim,
  type PresentationStyle,
  type SlideIR,
} from '@wiswork/pptx-engine/presentation'
import type { PowerPointDeclarativeOperation } from './browser-powerpoint-adapter.js'

type AddOperation = Extract<
  PowerPointDeclarativeOperation,
  { op: 'add_text_box' | 'add_geometric_shape' | 'add_native_table' }
>

/** Translate a complete supported SlideIR page into bounded native Office operations. */
export function officeOperationsForSlideIR(
  slide: SlideIR,
  style: PresentationStyle,
  slideIndex: number,
  claims: PresentationClaim[] = [],
  resolvedFontFace?: string,
): AddOperation[] {
  if (
    !Number.isSafeInteger(slideIndex) ||
    slideIndex < 0 ||
    slideIndex > 31 ||
    !Array.isArray(slide?.elements) ||
    !slide.elements.length ||
    slide.elements.length > 32 ||
    !style?.fontFace ||
    (style.fontFallbacks?.length && !resolvedFontFace) ||
    (resolvedFontFace !== undefined &&
      resolvedFontFace !== style.fontFace &&
      !style.fontFallbacks?.includes(resolvedFontFace)) ||
    style.fontFace.length > 128 ||
    !/^[0-9A-Fa-f]{6}$/.test(style.textColor) ||
    !/^[0-9A-Fa-f]{6}$/.test(style.accentColor)
  )
    throw new Error('invalid_tool_input')
  const fontFace = resolvedFontFace ?? style.fontFace
  const names = new Set<string>()
  const labels = presentationSlideSourceLabels(slide, claims)
  const sourceText = labels.join('；')
  if (
    slide.elements.length + Number(labels.length > 0) > 32 ||
    slide.elements.some((element) => element.id === 'source-attribution') ||
    sourceText.length > 500
  )
    throw new Error('invalid_tool_input')
  const operations: AddOperation[] = slide.elements.map((element): AddOperation => {
    if (!element.id || element.id.length > 256 || names.has(element.id))
      throw new Error('invalid_tool_input')
    names.add(element.id)
    if (
      [element.x, element.y, element.w, element.h].some((value) => !Number.isFinite(value)) ||
      element.w <= 0 ||
      element.h <= 0
    )
      throw new Error('invalid_tool_input')
    const box = {
      slide_index: slideIndex,
      name: element.id,
      left: element.x * 72,
      top: element.y * 72,
      width: element.w * 72,
      height: element.h * 72,
    }
    if (element.kind === 'text') {
      if (
        typeof element.text !== 'string' ||
        element.text.length > 12_000 ||
        (element.fontSize !== undefined &&
          (!Number.isFinite(element.fontSize) || element.fontSize < 6 || element.fontSize > 96)) ||
        (element.color !== undefined && !/^[0-9A-Fa-f]{6}$/.test(element.color))
      )
        throw new Error('invalid_tool_input')
      return {
        op: 'add_text_box',
        ...box,
        text: element.text,
        fontFace,
        fontSize: element.fontSize ?? 20,
        color: element.color ?? style.textColor,
        bold: element.bold ?? false,
        align: element.align ?? 'left',
        margin: 0,
        verticalAlignment: 'top',
      }
    }
    if (element.kind === 'shape') {
      if (
        !['rect', 'ellipse', 'roundRect'].includes(element.shape) ||
        (element.fill !== undefined && !/^[0-9A-Fa-f]{6}$/.test(element.fill)) ||
        (element.lineColor !== undefined && !/^[0-9A-Fa-f]{6}$/.test(element.lineColor))
      )
        throw new Error('invalid_tool_input')
      return {
        op: 'add_geometric_shape',
        ...box,
        shape: element.shape,
        fill: element.fill ?? style.accentColor,
        lineColor: element.lineColor ?? element.fill ?? style.accentColor,
      }
    }
    if (element.kind === 'table') {
      const rows = element.rows
      const width = rows?.[0]?.length
      if (
        !Array.isArray(rows) ||
        rows.length < 1 ||
        rows.length > 20 ||
        !width ||
        width > 12 ||
        rows.length * width > 128 ||
        rows.some(
          (row) =>
            !Array.isArray(row) ||
            row.length !== width ||
            row.some((cell) => typeof cell !== 'string' || cell.length > 256),
        ) ||
        JSON.stringify(rows).length > 12_000 ||
        (element.fontSize !== undefined &&
          (!Number.isFinite(element.fontSize) || element.fontSize < 6 || element.fontSize > 48))
      )
        throw new Error('invalid_tool_input')
      return {
        op: 'add_native_table',
        ...box,
        rows,
        fontFace,
        fontSize: element.fontSize ?? 16,
        color: style.textColor,
        borderColor: style.accentColor,
        cellMargin: 0.04 * 72,
      }
    }
    throw new Error('office_api_unsupported')
  })
  if (labels.length)
    operations.push({
      op: 'add_text_box',
      slide_index: slideIndex,
      name: 'source-attribution',
      text: sourceText,
      left: 36,
      top: 507.6,
      width: 885.6,
      height: 21.6,
      fontFace,
      fontSize: 8,
      color: style.textColor,
      bold: false,
      align: 'left',
      margin: 0,
      verticalAlignment: 'top',
    })
  return operations
}
