export interface StreamingProgressEntry {
  role: 'user' | 'assistant'
  text: string
  streaming?: boolean
}

export function shouldShowStreamingProgress(entry: StreamingProgressEntry): boolean {
  return entry.role === 'assistant' && entry.streaming === true
}
