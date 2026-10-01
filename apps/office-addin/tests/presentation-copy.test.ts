import { beforeEach, describe, expect, it, vi } from 'vitest'

const { exportDocument, supportsExport } = vi.hoisted(() => ({
  exportDocument: vi.fn(),
  supportsExport: vi.fn(),
}))
vi.mock('../src/skills/powerpoint/presentation-document-export.js', () => ({
  exportPowerPointDocument: exportDocument,
  supportsPowerPointDocumentExport: supportsExport,
}))

import {
  openPowerPointPresentationCopy,
  supportsPowerPointPresentationCopy,
} from '../src/skills/powerpoint/presentation-copy.js'

describe('PowerPoint copy workflow', () => {
  const createPresentation = vi.fn()
  beforeEach(() => {
    vi.clearAllMocks()
    supportsExport.mockReturnValue(true)
    exportDocument.mockResolvedValue(new Uint8Array([80, 75, 3, 4]))
    createPresentation.mockResolvedValue(undefined)
    vi.stubGlobal('Office', {
      context: { requirements: { isSetSupported: () => true } },
    })
    vi.stubGlobal('PowerPoint', { createPresentation })
  })

  it('opens an exported presentation as a new document', async () => {
    await openPowerPointPresentationCopy(async () => '["document-a","file:///source.pptx"]')
    expect(exportDocument).toHaveBeenCalledWith('pptx')
    expect(createPresentation).toHaveBeenCalledWith('UEsDBA==')
  })

  it('does not export when the host lacks the copy API', async () => {
    vi.stubGlobal('PowerPoint', undefined)
    expect(supportsPowerPointPresentationCopy()).toBe(false)
    await expect(
      openPowerPointPresentationCopy(async () => '["document-a","file:///source.pptx"]'),
    ).rejects.toThrow('presentation_copy_unavailable')
    expect(exportDocument).not.toHaveBeenCalled()
  })

  it('does not open a copy when the active document changes during export', async () => {
    const documentId = vi
      .fn()
      .mockResolvedValueOnce('["document-a","file:///source.pptx"]')
      .mockResolvedValueOnce('["document-a","file:///copy.pptx"]')
    await expect(openPowerPointPresentationCopy(documentId)).rejects.toThrow(
      'presentation_document_changed',
    )
    expect(createPresentation).not.toHaveBeenCalled()
  })

  it('requires saving an unsaved source before making a copy', async () => {
    await expect(openPowerPointPresentationCopy(async () => '["document-a",""]')).rejects.toThrow(
      'presentation_copy_save_source_first',
    )
    expect(exportDocument).not.toHaveBeenCalled()
    expect(createPresentation).not.toHaveBeenCalled()
  })

  it('uses the real document binding to reject Save As during export', async () => {
    let url = 'file:///source.pptx'
    const values = new Map<string, string>()
    vi.stubGlobal('Office', {
      AsyncResultStatus: { Succeeded: 'succeeded' },
      context: {
        document: {
          get url() {
            return url
          },
          settings: {
            get: (key: string) => values.get(key),
            set: (key: string, value: string) => values.set(key, value),
            saveAsync: (done: (result: { status: string }) => void) =>
              done({ status: 'succeeded' }),
          },
        },
        requirements: { isSetSupported: () => true },
      },
    })
    exportDocument.mockImplementation(async () => {
      url = 'file:///copy.pptx'
      return new Uint8Array([80, 75, 3, 4])
    })
    await expect(openPowerPointPresentationCopy()).rejects.toThrow('presentation_document_changed')
    expect(createPresentation).not.toHaveBeenCalled()
  })
})
