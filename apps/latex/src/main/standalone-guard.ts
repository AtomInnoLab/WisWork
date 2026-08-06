export function assertStandaloneDevelopment(isPackaged: boolean): void {
  if (isPackaged) throw new Error('Standalone LaTeX is disabled in packaged builds')
}
