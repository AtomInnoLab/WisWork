const en = {
  appTitle: 'WisWork LaTeX',
  files: 'Files',
  compile: 'Compile',
  compiling: 'Compiling…',
  cancel: 'Cancel',
  save: 'Save',
  newFile: 'New file',
  rename: 'Rename',
  preview: 'PDF preview',
  previewStale: 'Preview is out of date. Compile to refresh it.',
  logs: 'Compiler log',
  diagnostics: 'Diagnostics',
  noDiagnostics: 'No diagnostics',
  externalConflict: 'This file changed on disk. Your local edits were preserved.',
  loading: 'Loading…',
  compileFailed: 'Compilation failed',
  projectUnavailable: 'No LaTeX project is open.',
  previousPage: 'Previous page',
  nextPage: 'Next page',
  zoomOut: 'Zoom out',
  zoomIn: 'Zoom in',
  toolbar: 'LaTeX toolbar',
  showFiles: 'Show files',
  hideFiles: 'Hide files',
  showPreview: 'Show PDF preview',
  hidePreview: 'Hide PDF preview',
  showAi: 'Show WisWork AI',
  hideAi: 'Hide WisWork AI',
  unsavedChanges: 'Unsaved changes',
} as const

export type LatexStringKey = keyof typeof en
export type LatexStrings = Record<LatexStringKey, string>

const zhCN: LatexStrings = {
  appTitle: 'WisWork LaTeX',
  files: '文件',
  compile: '编译',
  compiling: '编译中…',
  cancel: '取消',
  save: '保存',
  newFile: '新建文件',
  rename: '重命名',
  preview: 'PDF 预览',
  previewStale: '预览已过期，请重新编译。',
  logs: '编译日志',
  diagnostics: '问题',
  noDiagnostics: '没有问题',
  externalConflict: '文件已在磁盘上更改，你的本地编辑已保留。',
  loading: '加载中…',
  compileFailed: '编译失败',
  projectUnavailable: '当前没有打开 LaTeX 项目。',
  previousPage: '上一页',
  nextPage: '下一页',
  zoomOut: '缩小',
  zoomIn: '放大',
  toolbar: 'LaTeX 工具栏',
  showFiles: '显示文件',
  hideFiles: '隐藏文件',
  showPreview: '显示 PDF 预览',
  hidePreview: '隐藏 PDF 预览',
  showAi: '显示 WisWork AI',
  hideAi: '隐藏 WisWork AI',
  unsavedChanges: '未保存的更改',
}

export const SUPPORTED_LATEX_LOCALES = [
  'en',
  'zh-CN',
  'zh-TW',
  'ja',
  'ko',
  'de',
  'fr',
  'es',
  'pt-BR',
  'it',
  'ru',
  'ar',
  'hi',
  'id',
  'tr',
  'vi',
] as const

export type LatexLocale = (typeof SUPPORTED_LATEX_LOCALES)[number]

export const LATEX_STRINGS: Record<LatexLocale, LatexStrings> = Object.fromEntries(
  SUPPORTED_LATEX_LOCALES.map((locale) => [locale, locale === 'zh-CN' ? zhCN : { ...en }]),
) as Record<LatexLocale, LatexStrings>

export function resolveLatexLocale(locale: string): LatexLocale {
  const exact = SUPPORTED_LATEX_LOCALES.find(
    (candidate) => candidate.toLowerCase() === locale.toLowerCase(),
  )
  if (exact) return exact
  const language = locale.split('-')[0]?.toLowerCase()
  return SUPPORTED_LATEX_LOCALES.find((candidate) => candidate === language) ?? 'en'
}
