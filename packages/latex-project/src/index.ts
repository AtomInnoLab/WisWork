export { AtomicWriteCommittedError } from './atomic-write.js'
export type { AtomicFileIdentity } from './atomic-write.js'
export { discoverMainFile } from './main-file.js'
export type { MainFileProjectReader } from './main-file.js'
export { DEFAULT_IMPORT_LIMITS, importLatexProject } from './import.js'
export type { ImportLimits, ImportResult } from './import.js'
export { ProjectPathPolicy, ProjectWriteConflictError } from './path-policy.js'
export { createLatexProject, LatexProject, openLatexProject } from './project.js'
export { ProposalStore } from './proposal.js'
export type {
  AppliedProposal,
  EditProposal,
  ProposalFile,
  ProposalStoreOptions,
} from './proposal.js'
export { SnapshotStore } from './snapshot.js'
export type { SnapshotRestoreOptions, SnapshotStoreOptions, SnapshotSummary } from './snapshot.js'
export {
  DEFAULT_MAX_TEXT_BYTES,
  type DeleteTextOptions,
  type LatexProjectOptions,
  type MainFileDiscovery,
  type MainFileDiscoveryOptions,
  type MainFileSource,
  type SavedText,
  type SaveTextOptions,
} from './types.js'
