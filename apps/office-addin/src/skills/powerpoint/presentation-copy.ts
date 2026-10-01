import {
  exportPowerPointDocument,
  supportsPowerPointDocumentExport,
} from './presentation-document-export.js'

export function supportsPowerPointPresentationCopy(): boolean {
  return (
    supportsPowerPointDocumentExport('pptx') &&
    typeof PowerPoint !== 'undefined' &&
    typeof PowerPoint.createPresentation === 'function' &&
    Office.context.requirements.isSetSupported('PowerPointApi', '1.1')
  )
}

export async function openPowerPointPresentationCopy(): Promise<void> {
  if (!supportsPowerPointPresentationCopy()) throw new Error('presentation_copy_unavailable')
  const bytes = await exportPowerPointDocument('pptx')
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000))
  }
  await PowerPoint.createPresentation(btoa(binary))
}
