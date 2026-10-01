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
    await openPowerPointPresentationCopy()
    expect(exportDocument).toHaveBeenCalledWith('pptx')
    expect(createPresentation).toHaveBeenCalledWith('UEsDBA==')
  })

  it('does not export when the host lacks the copy API', async () => {
    vi.stubGlobal('PowerPoint', undefined)
    expect(supportsPowerPointPresentationCopy()).toBe(false)
    await expect(openPowerPointPresentationCopy()).rejects.toThrow('presentation_copy_unavailable')
    expect(exportDocument).not.toHaveBeenCalled()
  })
})
