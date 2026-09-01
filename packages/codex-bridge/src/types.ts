export const ALLOWED_ENHANCED_CAPABILITIES = [
  'semantic-read',
  'transaction-proposal',
  'bounded-render-facts',
  'raw-office-proposal',
] as const

export const DENIED_ENHANCED_CAPABILITIES = [
  'shell',
  'arbitrary-filesystem',
  'git',
  'browser-control',
  'free-network',
  'direct-document-write',
] as const

export type EnhancedCapability = (typeof ALLOWED_ENHANCED_CAPABILITIES)[number]
export type CapabilityDeclaration = Readonly<{ capabilities: readonly EnhancedCapability[] }>

export const SAFE_RUNTIME_ERROR_CODES = [
  'runtime_unavailable',
  'runtime_blocked',
  'runtime_protocol_error',
  'runtime_input_rejected',
] as const
export type SafeRuntimeError = Readonly<{ code: (typeof SAFE_RUNTIME_ERROR_CODES)[number] }>

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
  | {
      type: 'custom_tool_call'
      id?: string
      call_id: string
      name: 'exec'
      input: string
      status?: 'completed'
    }
  | {
      type: 'custom_tool_call_output'
      id?: string
      call_id: string
      output: string | Array<{ type: 'input_text'; text: string }>
      status?: 'completed'
    }
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

declare const documentCarrierHandleBrand: unique symbol
export interface DocumentCarrierHandle {
  readonly [documentCarrierHandleBrand]: true
}

export interface DocumentCarrierIssuerContext {
  readonly host:
    'latex' | 'slides' | 'docs' | 'sheets' | 'office-word' | 'office-excel' | 'office-powerpoint'
  readonly documentId: string
  readonly sessionId: string
  readonly generation: number
}

export interface DocumentCarrierTurnContext {
  readonly turnId: string
  readonly sourceNonce: string
  readonly capability: unknown
  readonly method: string
  readonly toolName: string
  readonly schemaDigest: string
}

export interface DocumentCarrierIssuer {
  readonly issueForTurn: (context: unknown) => DocumentCarrierHandle
  readonly prepareTurn: (
    input: unknown,
    limitOverrides: Partial<ProtocolLimits>,
    handle: DocumentCarrierHandle,
  ) => PreparedResponsesTurn
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
