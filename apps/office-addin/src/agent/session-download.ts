import type { InMemoryVfs } from '../skills/shared/vfs.js'

export function downloadSessionFile(vfs: InMemoryVfs, path: string): void {
  if (!vfs.list('/home/user').includes(path)) throw new Error('vfs_path_denied')
  const bytes = vfs.readBytes(path)
  const mime = path.endsWith('.pptx')
    ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    : path.endsWith('.pdf')
      ? 'application/pdf'
      : path.endsWith('.json')
        ? 'application/json'
        : 'application/octet-stream'
  const url = URL.createObjectURL(new Blob([bytes.slice().buffer as ArrayBuffer], { type: mime }))
  const link = document.createElement('a')
  link.href = url
  link.download = path.split('/').at(-1) ?? 'download'
  document.body.append(link)
  link.click()
  link.remove()
  setTimeout(() => URL.revokeObjectURL(url), 10_000)
}
