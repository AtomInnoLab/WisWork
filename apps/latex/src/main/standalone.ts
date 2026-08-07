import { join } from 'node:path'
import { app, BrowserWindow, protocol } from 'electron'
import {
  configureLatexRuntime,
  createLatexView,
  getLatexSessionRegistry,
  registerLatexPdfProtocol,
} from './latex-main.js'
import { assertStandaloneDevelopment } from './standalone-guard.js'

assertStandaloneDevelopment(app.isPackaged)

protocol.registerSchemesAsPrivileged([
  {
    scheme: 'wiswork-latex-pdf',
    privileges: { secure: true, standard: true, supportFetchAPI: true },
  },
])

void app.whenReady().then(() => {
  const projectPath = process.env.WISWORK_LATEX_PROJECT
  const tectonicPath = process.env.WISWORK_TECTONIC_PATH
  if (!projectPath || !tectonicPath)
    throw new Error('Standalone LaTeX requires project and Tectonic paths')
  configureLatexRuntime({
    preloadPath: join(import.meta.dirname, '../preload/index.cjs'),
    rendererUrl: process.env.LATEX_RENDERER_URL ?? 'http://localhost:5177',
    tectonicPath,
    userDataPath: app.getPath('userData'),
  })
  registerLatexPdfProtocol(protocol, getLatexSessionRegistry())
  const window = new BrowserWindow({ width: 1400, height: 900 })
  const view = createLatexView(projectPath)
  window.contentView.addChildView(view)
  const resize = () => view.setBounds(window.getContentBounds())
  window.on('resize', resize)
  resize()
})
