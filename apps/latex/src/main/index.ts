export {
  configureLatexRuntime,
  createLatexView,
  getLatexSessionRegistry,
  latexQueryDirty,
  notifyLatexProjectRenamed,
  requestLatexEditFlush,
  releaseLatexEditFlush,
  registerLatexPdfProtocol,
  requestLatexClose,
} from './latex-main.js'

// Standalone mode is deliberately opt-in and development-only. Shell owns production app,
// authentication, protocol registration, navigation policy, and global lifecycle.
if (process.env.WISWORK_LATEX_STANDALONE === '1' && process.env.NODE_ENV !== 'production') {
  void import('./standalone.js')
}
