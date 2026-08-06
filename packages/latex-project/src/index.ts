export { AtomicWriteCommittedError } from './atomic-write.js'
export type { AtomicFileIdentity } from './atomic-write.js'
export { discoverMainFile } from './main-file.js'
export type { MainFileProjectReader } from './main-file.js'
export { ProjectPathPolicy } from './path-policy.js'
export { createLatexProject, LatexProject, openLatexProject } from './project.js'
export {
  DEFAULT_MAX_TEXT_BYTES,
  type LatexProjectOptions,
  type MainFileDiscovery,
  type MainFileDiscoveryOptions,
  type MainFileSource,
  type SavedText,
} from './types.js'
