export type ResponsesInput = string | ResponsesInputItem[]

export type ResponsesInputItem =
  | {
      type?: 'message'
      id?: string
      role: 'user' | 'assistant' | 'developer'
      content: Array<
        | { type: 'input_text' | 'output_text'; text: string }
        | { type: 'input_image'; image_url: string }
      >
    }
  | { type: 'custom_tool_call'; call_id: string; name: 'exec'; input: string }
  | { type: 'custom_tool_call_output'; call_id: string; output: string }
  | {
      type: 'additional_tools'
      role: 'developer'
      tools: Array<Record<string, unknown>>
    }

export interface ResponsesRequest {
  model: 'gpt-5.6-sol'
  input: ResponsesInput
  instructions?: string
  tools?: []
  tool_choice?: 'auto'
  max_output_tokens?: number
  parallel_tool_calls?: boolean
  reasoning?: { effort: 'medium'; context: 'all_turns' }
  store?: false
  stream?: true
  include?: ['reasoning.encrypted_content']
  prompt_cache_key?: string
  text?: { verbosity: 'low' }
  client_metadata?: Record<string, unknown>
}

export interface MessagesRequest {
  model: 'openai/gpt-5.6-sol'
  system?: string
  messages: Array<{
    role: 'user' | 'assistant'
    content: Array<Record<string, unknown>>
  }>
  tools?: Array<{
    name: string
    description?: string
    input_schema: Record<string, unknown>
  }>
  tool_choice?: { type: 'auto'; disable_parallel_tool_use: true }
  max_tokens: number
  stream: true
}

export interface PreparedResponsesTurn {
  readonly messagesRequest: MessagesRequest
  readonly messagesStreamToResponses: (
    chunks: AsyncIterable<string | Uint8Array>,
  ) => AsyncGenerator<string>
}

export interface ProtocolLimits {
  maxRequestItems: number
  maxRequestBytes: number
  maxRequestNodes: number
  maxNestingDepth: number
  maxContentParts: number
  maxTools: number
  maxStringLength: number
  maxDescriptionLength: number
  maxSchemaBytes: number
  maxOutputTokens: number
  maxPromptCacheKeyLength: number
  maxClientMetadataBytes: number
  maxWorkspaces: number
  maxSseFrameBytes: number
  maxSseBufferBytes: number
  maxSseFrames: number
  maxBlocks: number
  maxAccumulatedText: number
  maxToolArguments: number
  maxTotalOutput: number
}
