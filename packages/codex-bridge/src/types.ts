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
  | { type: 'function_call'; call_id: string; name: string; arguments: string }
  | { type: 'function_call_output'; call_id: string; output: string }
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
  tools?: Array<{
    type: 'function'
    name: string
    description?: string
    parameters: Record<string, unknown>
  }>
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

export type AdvertisedToolKind = 'function' | 'custom'

export interface StreamConversionContext {
  advertisedTools: Readonly<Record<string, AdvertisedToolKind>>
  usedCallIds: readonly string[]
  allowedExecMethods: readonly string[]
}

export interface RequestConversionResult {
  request: MessagesRequest
  context: StreamConversionContext
}

export interface ProtocolLimits {
  maxRequestItems: number
  maxContentParts: number
  maxTools: number
  maxStringLength: number
  maxDescriptionLength: number
  maxSchemaBytes: number
  maxOutputTokens: number
  maxPromptCacheKeyLength: number
  maxClientMetadataBytes: number
  maxSseFrameBytes: number
  maxSseBufferBytes: number
  maxSseFrames: number
  maxBlocks: number
  maxAccumulatedText: number
  maxToolArguments: number
  maxTotalOutput: number
}
