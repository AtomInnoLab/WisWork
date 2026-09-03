import { writeFile } from 'node:fs/promises'
import {
  addElement,
  createBlankPptx,
  duplicateSlide,
  openPptx,
  savePptx,
} from '@wiswork/pptx-engine'
import type { RenderSlide } from '@wiswork/pptx-render'

const EMU_PER_PIXEL = 9_525

export async function writePresentationE2eArtifact(
  slides: readonly RenderSlide[],
  path: string,
): Promise<{ readonly slideCount: number; readonly text: string }> {
  if (slides.length === 0) throw new Error('presentation_e2e_empty_deck')
  const opened = await openPptx(await createBlankPptx())
  for (let index = 1; index < slides.length; index += 1) {
    if (!duplicateSlide(opened, index - 1)) throw new Error('presentation_e2e_duplicate_failed')
  }
  for (const [slideIndex, rendered] of slides.entries()) {
    const target = opened.deck.slides[slideIndex]
    if (!target) throw new Error('presentation_e2e_slide_missing')
    for (const node of rendered.nodes) {
      if ((node.type !== 'text' && node.type !== 'shape') || !node.text) continue
      const paragraphs = node.text.lines
        .map((line) => ({
          runs: line.runs.map((run) => ({ text: run.text })).filter((run) => run.text.length > 0),
        }))
        .filter((paragraph) => paragraph.runs.length > 0)
      if (paragraphs.length === 0) continue
      addElement(target, {
        kind: 'textbox',
        offset: {
          x: Math.max(0, Math.round(node.box.x * EMU_PER_PIXEL)),
          y: Math.max(0, Math.round(node.box.y * EMU_PER_PIXEL)),
          cx: Math.max(EMU_PER_PIXEL, Math.round(node.box.w * EMU_PER_PIXEL)),
          cy: Math.max(EMU_PER_PIXEL, Math.round(node.box.h * EMU_PER_PIXEL)),
        },
        paragraphs,
      })
    }
  }
  const bytes = await savePptx(opened)
  await writeFile(path, bytes)
  const reopened = await openPptx(bytes)
  return {
    slideCount: reopened.deck.slides.length,
    text: reopened.deck.slides
      .flatMap((slide) => slide.elements)
      .flatMap((element) =>
        element.type === 'text' || element.type === 'shape'
          ? (element.text?.paragraphs ?? []).flatMap((paragraph) =>
              paragraph.runs.map((run) => run.text),
            )
          : [],
      )
      .join('\n'),
  }
}
