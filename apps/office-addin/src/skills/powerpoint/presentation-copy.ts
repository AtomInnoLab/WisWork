import {
  exportPowerPointDocument,
  supportsPowerPointDocumentExport,
} from './presentation-document-export.js'
import { createBrowserPresentationDocumentBinding } from './presentation-document.js'

export function supportsPowerPointPresentationCopy(): boolean {
  return (
    supportsPowerPointDocumentExport('pptx') &&
    typeof PowerPoint !== 'undefined' &&
    typeof PowerPoint.createPresentation === 'function' &&
    Office.context.requirements.isSetSupported('PowerPointApi', '1.1')
  )
}

export async function openPowerPointPresentationCopy(
  documentId: () => Promise<string> = () => createBrowserPresentationDocumentBinding().documentId(),
): Promise<void> {
  if (!supportsPowerPointPresentationCopy()) throw new Error('presentation_copy_unavailable')
  const sourceId = await documentId()
  // Office can copy document.settings. Two unsaved documents may otherwise
  // share both binding fields until one receives a distinct URL.
  if (!JSON.parse(sourceId)[1]) throw new Error('presentation_copy_save_source_first')
  const bytes = await exportPowerPointDocument('pptx')
  if ((await documentId()) !== sourceId) throw new Error('presentation_document_changed')
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000))
  }
  await PowerPoint.createPresentation(btoa(binary))
}
