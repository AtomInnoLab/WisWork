// GENERATED SUPPORTED SUBSET — DO NOT EDIT BY HAND.
// Source: codex-cli 0.147.0 stable app-server schemas.
// Regenerate/check with `npm run schema:check -w @wiswork/codex-bridge -- /absolute/path/to/codex`.

export const CODEX_CLI_VERSION = 'codex-cli 0.147.0' as const

export const CODEX_SCHEMA_SHA256 = Object.freeze({
  protocol: 'f72b2caa3cbfa4298de9e85c62dda6dfbaf2266ffeb916fed30615ca69ff8c74',
  protocolV2: 'f3dec1e031d99a420b137b903f02196d4325eece57620c925bb7130b25f168d2',
  initialize: '6f0094be9a65242ec779a40794cbd4fdfa32fca1e45084a16adfb50501d33ea2',
  initializeResponse: '62ad689c2cb6379913c1d72749cfd8de5089d35760214123518eb92eef11acc9',
  threadStart: '792e2f32e37cece971bd616664ea2053741acbed4e9c92e9d1766427718f2ecd',
  threadStartResponse: '86d37438580ac2da25fa144ee9b4dda1269953cd92c9f9c94d7a2904a28a9517',
  turnStart: 'ff2e7e0796fbe2ad99e5ec7d489cc8c8630b75f2ab8f17857711107587e3197d',
  turnStartResponse: '888a437b6c818b10fb2ec3c255cb71c82293b378d9ce47e7be9c0028ed9e8a34',
  turnInterrupt: '6dff382dae73d1dbc58406ed045605f647e7a49660e2540fbd2c6c24d60c5f2b',
  turnInterruptResponse: '531de6be06fe979b5963f249bab82498a175e614bf65ac12fb2e849dfe60bcf1',
  serverNotification: '558aab11da20cd0d278ec6f9067a46444f0566be7e3953f443dcc5d7cb23736f',
})

export interface InitializeParams {
  clientInfo: { name: string; version: string }
  capabilities: null
}

export interface InitializeResponse {
  userAgent: string
  codexHome: string
  platformFamily: string
  platformOs: string
}

export interface ThreadStartParams {
  model: 'gpt-5.6-sol'
  modelProvider: 'wiswork'
  cwd: string
  approvalPolicy: 'never'
  sandbox: 'read-only'
  developerInstructions: string
  ephemeral: true
}

export interface ThreadStartResponse {
  thread: { id: string; [key: string]: unknown }
  [key: string]: unknown
}

export interface TurnStartParams {
  threadId: string
  input: [{ type: 'text'; text: string; text_elements: [] }]
  effort: 'medium'
}

export interface TurnStartResponse {
  turn: { id: string; [key: string]: unknown }
}

export interface TurnInterruptParams {
  threadId: string
  turnId: string
}

export type TurnInterruptResponse = Record<string, never>

export const KNOWN_SERVER_NOTIFICATION_METHODS = Object.freeze([
  'account/login/completed',
  'account/rateLimits/updated',
  'account/updated',
  'app/list/updated',
  'command/exec/outputDelta',
  'configWarning',
  'deprecationNotice',
  'error',
  'externalAgentConfig/import/completed',
  'externalAgentConfig/import/progress',
  'fs/changed',
  'fuzzyFileSearch/sessionCompleted',
  'fuzzyFileSearch/sessionUpdated',
  'guardianWarning',
  'hook/completed',
  'hook/started',
  'item/agentMessage/delta',
  'item/autoApprovalReview/completed',
  'item/autoApprovalReview/started',
  'item/commandExecution/outputDelta',
  'item/commandExecution/terminalInteraction',
  'item/completed',
  'item/fileChange/outputDelta',
  'item/fileChange/patchUpdated',
  'item/mcpToolCall/progress',
  'item/plan/delta',
  'item/reasoning/summaryPartAdded',
  'item/reasoning/summaryTextDelta',
  'item/reasoning/textDelta',
  'item/started',
  'mcpServer/oauthLogin/completed',
  'mcpServer/startupStatus/updated',
  'model/rerouted',
  'model/safetyBuffering/updated',
  'model/verification',
  'process/exited',
  'process/outputDelta',
  'remoteControl/status/changed',
  'serverRequest/resolved',
  'skills/changed',
  'thread/archived',
  'thread/closed',
  'thread/compacted',
  'thread/deleted',
  'thread/environment/connected',
  'thread/environment/disconnected',
  'thread/goal/cleared',
  'thread/goal/updated',
  'thread/name/updated',
  'thread/realtime/closed',
  'thread/realtime/error',
  'thread/realtime/itemAdded',
  'thread/realtime/outputAudio/delta',
  'thread/realtime/sdp',
  'thread/realtime/started',
  'thread/realtime/transcript/delta',
  'thread/realtime/transcript/done',
  'thread/settings/updated',
  'thread/started',
  'thread/status/changed',
  'thread/tokenUsage/updated',
  'thread/unarchived',
  'turn/completed',
  'turn/diff/updated',
  'turn/moderationMetadata',
  'turn/plan/updated',
  'turn/started',
  'warning',
  'windows/worldWritableWarning',
  'windowsSandbox/setupCompleted',
] as const)

export type KnownServerNotificationMethod = (typeof KNOWN_SERVER_NOTIFICATION_METHODS)[number]

interface Notification<TMethod extends KnownServerNotificationMethod, TParams> {
  method: TMethod
  params: TParams
  emittedAtMs?: number
}

export type CodexAppServerNotification =
  | Notification<
      'item/agentMessage/delta',
      { threadId: string; turnId: string; itemId: string; delta: string }
    >
  | Notification<'turn/started', { threadId: string; turn: { id: string; [key: string]: unknown } }>
  | Notification<
      'turn/completed',
      { threadId: string; turn: { id: string; [key: string]: unknown } }
    >
  | Notification<
      Exclude<
        KnownServerNotificationMethod,
        'item/agentMessage/delta' | 'turn/started' | 'turn/completed'
      >,
      unknown
    >
