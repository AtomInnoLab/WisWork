import { readFile } from 'node:fs/promises'
import { extname } from 'node:path'
import { docxToText } from './docx'
import { pdfToPagesWithImageCoverage } from './pdf'
import { pptxToText } from './pptx'
import { xlsxToText } from './xlsx'
import { decodeHtmlBytes, htmlToText } from './html'

export type ParsedFileKind = 'text' | 'image' | 'unsupported'

export interface ParsedFile {
  ok: boolean
  text?: string
  kind: ParsedFileKind
  mime?: string
  error?: string
  sections?: { locator: string; start: number; end: number }[]
  pagesWithFullPageImage?: number[]
}

export function paragraphSections(text: string): NonNullable<ParsedFile['sections']> {
  const sections: NonNullable<ParsedFile['sections']> = []
  let start = 0
  for (const [index, paragraph] of text.split('\n').entries()) {
    sections.push({ locator: `第 ${index + 1} 段`, start, end: start + paragraph.length })
    start += paragraph.length + 1
  }
  return sections
}

/** No text extraction for images: callers read raw bytes and go multimodal (see @wiswork/ai-provider images support) */
const IMAGE_MIMES: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
}

const TEXT_EXTS = new Set(['txt', 'md', 'markdown', 'csv', 'tsv', 'json', 'xml', 'log'])

/** parse an attachment into plain text (or flag it as image / unsupported) */
export async function parseFileToText(filePath: string): Promise<ParsedFile> {
  const ext = extname(filePath).slice(1).toLowerCase()
  const imageMime = IMAGE_MIMES[ext]
  if (imageMime) return { ok: true, kind: 'image', mime: imageMime }
  try {
    if (ext === 'html' || ext === 'htm') {
      const text = htmlToText(decodeHtmlBytes(await readFile(filePath)))
      return { ok: true, kind: 'text', text, sections: paragraphSections(text) }
    }
    if (TEXT_EXTS.has(ext)) {
      return { ok: true, kind: 'text', text: await readFile(filePath, 'utf-8') }
    }
    switch (ext) {
      case 'docx': {
        const text = await docxToText(await readFile(filePath))
        return { ok: true, kind: 'text', text, sections: paragraphSections(text) }
      }
      case 'pptx':
        return { ok: true, kind: 'text', text: await pptxToText(await readFile(filePath)) }
      case 'xlsx':
        return { ok: true, kind: 'text', text: await xlsxToText(await readFile(filePath)) }
      case 'pdf': {
        const { pages, pagesWithFullPageImage } = await pdfToPagesWithImageCoverage(
          await readFile(filePath),
        )
        if (pages.every((page) => !page.trim()))
          return { ok: false, kind: 'text', error: 'pdf_no_extractable_text' }
        const sections: NonNullable<ParsedFile['sections']> = []
        let offset = 0
        for (const [index, page] of pages.entries()) {
          sections.push({ locator: `第 ${index + 1} 页`, start: offset, end: offset + page.length })
          offset += page.length + (index < pages.length - 1 ? 2 : 0)
        }
        return {
          ok: true,
          kind: 'text',
          text: pages.join('\n\n'),
          sections,
          ...(pagesWithFullPageImage.length ? { pagesWithFullPageImage } : {}),
        }
      }
    }
  } catch (e) {
    return { ok: false, kind: 'text', error: e instanceof Error ? e.message : String(e) }
  }
  return { ok: false, kind: 'unsupported', error: `Unsupported file type: .${ext || 'unknown'}` }
}
