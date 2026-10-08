export interface StreamingProgressEntry {
  role: 'user' | 'assistant'
  text: string
  streaming?: boolean
  tools?: readonly unknown[]
}

export type PresentationTimelineBlock = 'message' | 'tools'

export function presentationTimelineBlockOrder(
  entry: StreamingProgressEntry & { readonly hasMessageChrome?: boolean },
): readonly PresentationTimelineBlock[] {
  const blocks: PresentationTimelineBlock[] = []
  if (
    entry.role === 'user' ||
    !!entry.text ||
    shouldShowStreamingProgress(entry) ||
    entry.hasMessageChrome
  ) {
    blocks.push('message')
  }
  if (entry.tools?.length) blocks.push('tools')
  return blocks
}

export function shouldShowStreamingProgress(entry: StreamingProgressEntry): boolean {
  return entry.role === 'assistant' && entry.streaming === true && !entry.tools?.length
}
