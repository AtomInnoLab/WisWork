export { ProjectStore } from './store.js'
export type {
  ChatMessage,
  ChatMeta,
  ProjectData,
  ProjectIndex,
  ProjectInfo,
  ProjectSummary,
  TimelineEntry,
  ToolActivity,
} from './types.js'
export type {
  AppendChatArgs,
  LoadChatArgs,
  ProjectApi,
  RebindChatArgs,
  ResolveChatArgs,
  ResolveChatResult,
} from './ipc.js'
export { PresentationStore, assertPresentationId } from './presentation-store.js'
export type {
  PresentationClaimReviewRecord,
  PresentationReceipt,
  PresentationProductionPage,
  PresentationProductionRecord,
  PresentationPlanBinding,
  PresentationPlanRecord,
  PresentationPlanRevisionSnapshot,
} from './presentation-store.js'
export { parsePresentationProductionJob } from './presentation-job.js'
export type {
  PresentationProductionJob,
  PresentationProductionJobEvent,
  PresentationProductionJobEventInput,
  PresentationProductionJobState,
} from './presentation-job.js'

export {
  parsePresentationIssueLedger,
  parsePresentationIssueActionInput,
} from './presentation-issue.js'
export type {
  PresentationIssueActionInput,
  PresentationIssueAction,
  PresentationIssueLedger,
} from './presentation-issue.js'
export {
  PresentationLifecycleStore,
  DEFAULT_PRESENTATION_RETENTION_POLICY,
  PRESENTATION_LIFECYCLE_RESOURCE_KINDS,
  MAX_PRESENTATION_LIFECYCLE_BYTES,
  parsePresentationLifecycle,
} from './presentation-lifecycle.js'
export type {
  PresentationRetentionPolicy,
  PresentationLifecycleScope,
  PresentationLifecycleRecord,
  PresentationLifecycleResource,
  PresentationLifecycleResourceKind,
  PresentationLifecycleResourceStatus,
  PresentationLifecycleResultCode,
  PresentationDeletionIntent,
  PresentationDeletionResult,
  PresentationLifecycleAuditEvent,
} from './presentation-lifecycle.js'
