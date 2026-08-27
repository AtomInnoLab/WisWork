import { readFile } from 'node:fs/promises'
import { atomicWriteFile } from '@wiswork/latex-project'
import type { ExportPdfResult } from '../shared/ipc.js'

export interface SavePdfDialog {
  showSaveDialog(options: {
    title: string
    defaultPath: string
    filters: Array<{ name: string; extensions: string[] }>
  }): Promise<{ canceled: boolean; filePath?: string }>
}

export async function exportPublishedPdf(
  dialog: SavePdfDialog,
  sourcePath: string,
  suggestedName: string,
): Promise<ExportPdfResult> {
  const selection = await dialog.showSaveDialog({
    title: 'Export PDF',
    defaultPath: suggestedName,
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  })
  if (selection.canceled || !selection.filePath) return { state: 'cancelled' }

  const bytes = await readFile(sourcePath)
  await atomicWriteFile(selection.filePath, bytes)
  return { state: 'written', path: selection.filePath }
}
