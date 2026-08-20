export function isAiSensitivePath(path: string): boolean {
  return path
    .replaceAll('\\', '/')
    .split('/')
    .some((part) => {
      const lower = part.toLowerCase()
      return (
        lower === '.env' ||
        lower.startsWith('.env.') ||
        lower.includes('secret') ||
        lower.includes('credential') ||
        lower.includes('private-key') ||
        lower.includes('private_key')
      )
    })
}
