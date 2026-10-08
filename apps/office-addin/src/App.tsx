import { PresentationProjectGovernanceCard } from './agent/presentation-project-governance-card.js'
import { createPresentationGovernanceStorage } from './agent/presentation-project-governance-storage.js'
import type { createPresentationProjectGovernanceController } from './agent/presentation-project-governance.js'
import { createBrowserAuthClient, createMemorySessionStore } from '@wiswork/auth/browser'
import { officeTeamAuthConfig } from './agent/team-auth-config.js'
import { createOfficeTeamLoginDialog } from './agent/team-login-dialog.js'
import { createOfficeTeamConnection, type OfficeTeamConnection } from './agent/team-connection.js'
import { PresentationTeamCard } from './agent/presentation-team-card.js'
import type { PresentationTeamController } from './agent/presentation-team-controller.js'
import { createPresentationResearchAbandonPersistence } from './agent/presentation-research-recovery-storage.js'
import { createPresentationResearchDeletePersistence } from './agent/presentation-research-cleanup-storage.js'
import { PresentationResearchCard } from './agent/presentation-research-card.js'
import type { PresentationResearchController } from './agent/presentation-research.js'
import type { PresentationAcquisitionHistory } from '@wiswork/project-store/presentation-acquisition'
import { PresentationAcquisitionHistoryCard } from './agent/presentation-acquisition-history.js'
import { PresentationChangesCard } from './agent/presentation-changes-card.js'
import type { PresentationChangesController } from './agent/presentation-changes.js'
import { validatePresentationQaRecord } from './skills/powerpoint/presentation-qa.js'
import { PresentationQaCard, type PresentationQaController } from './agent/presentation-qa-card.js'
import {
  PresentationImportProgressCard,
  type PresentationImportProgressController,
} from './agent/presentation-import-progress.js'
import {
  MAX_PRESENTATION_ATTACHMENT_BYTES,
  MAX_PRESENTATION_IMAGE_BYTES,
  isPresentationImage,
  supportsPresentationAttachment,
} from './skills/powerpoint/presentation-attachments.js'
import { PresentationProjectCard } from './agent/presentation-project-card.js'
import {
  openPowerPointPresentationCopy,
  supportsPowerPointPresentationCopy,
} from './skills/powerpoint/presentation-copy.js'
import { PresentationWorkflowCard } from './agent/presentation-workflow-card.js'
import { PresentationStageCard } from './agent/presentation-stage-card.js'
import { presentationStageTimeline } from './agent/presentation-stage-timeline.js'
import type { PresentationProjectController } from './skills/powerpoint/presentation-project.js'
import {
  createBrowserPresentationDocumentBinding,
  createPresentationAgentRunCheckpoint,
  preparePresentationAgentRunRecovery,
} from './skills/powerpoint/presentation-document.js'
import { downloadSessionFile } from './agent/session-download.js'
import { useEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react'
import { deployedBuildId, resolveBuildVersion, type BuildVersionState } from './build-version.js'
import {
  AiTypingIndicator,
  IconEnter,
  IconPaperclip,
  Markdown,
  PresentationActivityGroup,
  PresentationEmptyState,
  PresentationMessage,
} from '@wiswork/ui'
import { extractPresentationDesignDocument } from '@wiswork/agent-core'
import {
  normalizeLang,
  translatePresentationVerification,
  translateRawOfficeConfirmation,
} from '@wiswork/i18n'
import { createOfficeWebSkill } from './skills/shared/web-skill.js'
import { createOfficeHostRuntime, type OfficeHostRuntime } from './agent/host-runtime.js'
import type { PresentationAttachmentMetadata } from './skills/powerpoint/presentation-attachments.js'
import {
  officeCapabilityFlags,
  officeDiagnosticSamplePercent,
  officeRemoteDiagnosticsEnabled,
  officePresentationRolloutPercent,
  officeWorkspaceMode,
  presentationRolloutEnabled,
} from '../build-config.js'
import { officePresentationVerificationFlags } from './agent/presentation-flags.js'
import {
  createOfficeDiagnostics,
  officeDiagnosticEnvironment,
  type OfficeDiagnostics,
} from './diagnostics/office-diagnostics.js'
import type { OfficeProposal, StructuredProposal } from './agent/proposal-controller.js'
import { MAX_SKILL_BYTES } from './skills/shared/skill-registry.js'
import { MAX_VFS_FILE_BYTES, MAX_VFS_TOTAL_BYTES } from './skills/shared/vfs.js'
import { createPcBridgeAgentTransport } from './agent/transport.js'
import {
  createOfficeAgentSession,
  useOfficeAgent,
  type OfficeAgentSession,
} from './agent/use-office-agent.js'
import {
  presentationProgressLabel,
  type OfficePresentationEvent,
  OfficeClarificationQuestion,
  OfficePresentationTimeline,
  ProposalPresentationEvent,
} from './agent/presentation-state.js'
import { createPcBridgeSession, type PcBridgeSession } from './pc-bridge/session.js'
import {
  createOfficeRelaySession,
  officeTransportMode,
  type OfficeRelaySession,
  type OfficeRelaySnapshot,
  type OfficeRelayStatus,
} from './relay/session.js'
import type { PresentationVerificationStringKey } from '@wiswork/i18n'
import { rawOfficeCapabilities } from './agent/enhanced-session.js'
import { OfficeDesignPanel, type OfficeDesignRequest } from './OfficeDesignPanel.js'
import { createOfficeDesignRequest } from './relay/design-document.js'

export const officePresentationText = (
  locale: string | null | undefined,
  key: PresentationVerificationStringKey,
) => translatePresentationVerification(normalizeLang(locale), key)
import {
  createBrowserOfficeRuntime,
  createOfficeDocumentClient,
  type OfficeDocumentClient,
  type OfficeHost,
} from './office-document.js'

const hostLabels: Record<OfficeHost, string> = {
  word: 'Microsoft Word',
  excel: 'Microsoft Excel',
  powerpoint: 'Microsoft PowerPoint',
  unknown: 'Office',
}

const sessionAttachmentLabel = (file: string) =>
  file.startsWith('/home/user/generated/')
    ? file.slice('/home/user/generated/'.length).replaceAll('/', ' / ')
    : (file.split('/').at(-1) ?? file)

const agentProductLabels: Record<OfficeHost, string> = {
  word: 'AI Word',
  excel: 'AI Sheets',
  powerpoint: 'AI Slides',
  unknown: 'WisWork AI',
}

export function officeRuntimeModeForTaskpane(
  host: OfficeHost,
  snapshot: {
    readonly status: string
    readonly enhanced?: { readonly host: string; readonly expires_at: number }
  },
  now = Date.now(),
): 'standard' | 'enhanced' {
  return snapshot.status === 'connected' &&
    host !== 'unknown' &&
    snapshot.enhanced?.host === `office-${host}` &&
    snapshot.enhanced.expires_at > now
    ? 'enhanced'
    : 'standard'
}

export function relayConnectionPresentation(
  status: OfficeRelayStatus | 'signed_out',
  verificationCode?: string,
) {
  const detail = {
    offline: 'Connect again to create a new secure pairing with WisWork PC.',
    connecting: 'Connecting securely to the WisWork Office Relay…',
    reconnecting: 'Reconnecting to WisWork PC…',
    signed_out: 'Sign in to WisWork PC first.',
    pending: verificationCode
      ? `Enter code ${verificationCode} in WisWork PC, then approve the matching request.`
      : 'Enter the pairing code in WisWork PC.',
    waiting_for_pc: verificationCode
      ? `Enter code ${verificationCode} in WisWork PC to continue.`
      : 'Waiting for a signed-in WisWork PC.',
    rejected: 'The connection was rejected in WisWork PC.',
    expired: 'The connection request expired. Try again.',
    incompatible: 'Upgrade Office Relay, then connect again.',
    pc_incompatible: 'Upgrade WisWork PC, then connect again.',
    connected: '',
  }[status]
  const busy = status === 'connecting' || status === 'reconnecting' || status === 'pending'
  return Object.freeze({
    title:
      status === 'reconnecting'
        ? 'Reconnecting to WisWork PC…'
        : status === 'waiting_for_pc' && !verificationCode
          ? 'Waiting for WisWork PC'
          : 'Connect to WisWork PC',
    detail,
    busy,
    actionDisabled: busy,
  })
}

export function shouldResetOfficeSession(status: OfficeRelayStatus | 'signed_out'): boolean {
  return status === 'rejected' || status === 'expired' || status === 'signed_out'
}

export function shouldShowRelayStatusScreen(
  status: OfficeRelayStatus | 'signed_out',
  hasWorkspace: boolean,
): boolean {
  if (status === 'connected') return false
  if (!hasWorkspace) return true
  return !['connecting', 'reconnecting', 'waiting_for_pc'].includes(status)
}

export function relayPersistenceNotice(snapshot: OfficeRelaySnapshot): string | undefined {
  return snapshot.status === 'connected' && snapshot.remembered === false
    ? 'Connected, but this Office installation was not remembered. Pair again after reconnecting.'
    : undefined
}

type DisplayProposal = OfficeProposal | StructuredProposal

function isLegacyProposal(proposal: DisplayProposal): proposal is OfficeProposal {
  return 'value' in proposal
}

function humanLabel(value: string): string {
  const words = value
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replaceAll('_', ' ')
    .replaceAll('-', ' ')
    .toLowerCase()
    .replace(/\bid\b/g, 'ID')
  return words.replace(/^./, (character) => character.toUpperCase())
}

function proposalHostLabel(host: string): string {
  const normalized = host
    .trim()
    .toLowerCase()
    .replace(/^microsoft\s+/, '')
  if (normalized === 'word') return 'Word'
  if (normalized === 'excel') return 'Excel'
  if (normalized === 'powerpoint' || normalized === 'power point') return 'PowerPoint'
  return 'Office'
}

const INTERNAL_PREVIEW_KEY = /(fingerprint|hash|code|xml|operation|program|payload|request)/i

function previewSummary(preview: Readonly<Record<string, unknown>>): string {
  const lines: string[] = []
  for (const [key, value] of Object.entries(preview)) {
    if (lines.length >= 12 || INTERNAL_PREVIEW_KEY.test(key)) continue
    if (!['string', 'number', 'boolean'].includes(typeof value) && value !== null) continue
    const rendered = value === null ? 'None' : String(value).replaceAll('\n', ' · ')
    const readable =
      /id$/i.test(key) && typeof value === 'string' ? proposalTarget(rendered) : rendered
    lines.push(`${humanLabel(key)}: ${readable}`)
  }
  return lines.join('\n').slice(0, 2_000)
}

function proposalTarget(target: string, host?: string): string {
  if (host && proposalHostLabel(host) === 'PowerPoint') {
    const indexedShape = /^(\d+)\/(\d+)$/.exec(target)
    if (indexedShape) return `Slide ${Number(indexedShape[1]) + 1} · Shape ${indexedShape[2]}`
    const packageSlide = /^ppt\/slides\/slide(\d+)\.xml$/i.exec(target)
    if (packageSlide) return `Slide ${packageSlide[1]} package`
  }
  if (target === 'document:end') return 'End of document'
  if (target === 'document:start') return 'Start of document'
  if (target === 'document') return 'Document'
  if (target === 'selection') return 'Current selection'
  const slide = /^slide[-:](\d+)$/i.exec(target)
  if (slide) return `Slide ${slide[1]}`
  return target
    .split('/')
    .filter(Boolean)
    .map((part) => (/^[A-Z]+\d+(?::[A-Z]+\d+)?$/i.test(part) ? part : humanLabel(part)))
    .join(' · ')
}

function isDisplaySafeProposalTarget(target: string): boolean {
  // Office object-path identifiers are implementation details and are not stable or meaningful
  // enough for a user confirmation surface.
  return !/^\d+(?:#\d+)+$/.test(target)
}

export function proposalPresentation(proposal: DisplayProposal) {
  const legacy = isLegacyProposal(proposal)
  const hasComparison =
    legacy || (typeof proposal.before === 'string' && typeof proposal.after === 'string')
  const before = hasComparison && typeof proposal.before === 'string' ? proposal.before : ''
  const after = hasComparison
    ? String(
        legacy
          ? proposal.operation === 'replace'
            ? proposal.value
            : `${proposal.before}${proposal.value}`
          : proposal.after,
      )
    : ''
  return {
    title: legacy
      ? proposal.operation === 'replace'
        ? 'Replace selection'
        : 'Append to selection'
      : proposal.title,
    host: legacy ? undefined : proposalHostLabel(proposal.impact.host),
    count: legacy ? undefined : proposal.impact.count,
    targets: legacy
      ? []
      : proposal.impact.targets
          .filter(isDisplaySafeProposalTarget)
          .map((target) => proposalTarget(target, proposal.impact.host)),
    before,
    after,
    preview:
      !legacy &&
      hasComparison &&
      (proposal.preview.beforeTruncated === true || proposal.preview.afterTruncated === true)
        ? '部分文本预览已截断，仅展示开头片段；实际操作会作用于完整文本。请核对完整内容后再确认。'
        : hasComparison || legacy
          ? ''
          : previewSummary(proposal.preview),
    // Declarative code is an internal safety protocol, not user-facing review content.
    code: undefined,
    ...(!legacy && proposal.lockReview ? { lockReview: proposal.lockReview } : {}),
  }
}

const MEBIBYTE = 1024 * 1024

function displayMebibytes(bytes: number): string {
  const value = bytes / MEBIBYTE
  return Number.isInteger(value) ? String(value) : value.toFixed(1)
}

export function safeUploadError(error: unknown, file?: Pick<SessionFile, 'size'>): string {
  const code = error instanceof Error ? error.message : ''
  if (code === 'presentation_animated_image_unsupported')
    return file
      ? '动画原件已保留在 PC。请在附件列表中选择“生成静态首帧”，或改用静态 PNG/JPEG。'
      : '动画网址未保存。请先下载并上传原件，再在附件列表中选择“生成静态首帧”。'
  const attachmentErrors: Record<string, string> = {
    presentation_attachment_too_large: '制作资料每个文件最多 50 MiB。',
    presentation_text_too_long: '粘贴资料最多 100 万字符。请拆分后分别保存。',
    presentation_image_too_large: '图片每个文件最多 10 MiB。',
    presentation_remote_image_unavailable:
      '图片网址无法安全下载或图片格式不受支持，请检查网址后重试。',
    presentation_image_candidates_exhausted:
      '这些候选图片都无法安全下载或解码，请换一组网址后重试。',
    invalid_tool_input: '请输入 1 到 4 个不同的 HTTP(S) 图片网址，每行一个。',
    presentation_remote_image_source_conflict:
      '相同图片内容已从另一来源加入当前文档。请使用已有素材，或手动上传本地文件。',
    presentation_remote_webpage_unavailable:
      '网页无法安全抓取，或未返回 HTML。请检查网址后重试，也可上传保存的 HTML 文件。',
    presentation_remote_webpage_source_conflict:
      '相同网页内容已作为另一份资料加入当前文档。请使用已有附件。',
    presentation_webpages_unavailable: '请更新并连接支持网页资料的 PC 端后重试。',
    presentation_aborted: '下载超时或已取消，请重试。',
    presentation_parse_failed: '图片无法解码为受支持的 PNG、JPEG、静态 GIF 或 WebP。',
    presentation_animated_image_staged:
      '动画网址原件已保存在 PC，但不会直接用于页面。请在附件列表中选择“生成静态首帧”，或使用其他静态图片。',
    presentation_assets_unavailable: '请更新并连接支持图片素材的 PC 端后重试。',
    presentation_attachment_failed: '资料解析未完成，请检查文件或重新上传。',
    presentation_not_found: '这份 PC 资料已不存在，请刷新附件列表。',
    presentation_attachment_in_use: '这份资料正被图片使用权声明引用，请先撤回声明再删除。',
    presentation_invalid_state: 'PC 资料状态异常，请重连后重试。',
    presentation_quota_exceeded:
      '当前文档在 PC 的资料容量已满（资料最多 32 个，附件预留容量总计 100 MiB）。请删除不再需要的资料或图片后重试。',
    presentation_document_changed: '文档已改变，本次上传已停止。请在目标文档重新上传。',
    presentation_service_unavailable: 'PC 连接不可用，请重连后重新选择同一文件续传。',
    presentation_unavailable: 'PC 连接不可用，请重连后重新选择同一文件续传。',
    presentation_response_invalid: '上传响应无效，请重连后重新选择同一文件续传。',
  }
  if (attachmentErrors[code]) return attachmentErrors[code]
  if (code === 'vfs_limit') {
    if (file && file.size > MAX_VFS_FILE_BYTES) {
      return `File is ${displayMebibytes(file.size)} MiB. Attachments must be ${displayMebibytes(MAX_VFS_FILE_BYTES)} MiB or smaller.`
    }
    return `Attachment limit reached. Files are limited to ${displayMebibytes(MAX_VFS_FILE_BYTES)} MiB each and ${displayMebibytes(MAX_VFS_TOTAL_BYTES)} MiB per session.`
  }
  return [
    'upload_cancelled',
    'vfs_path_denied',
    'invalid_skill_package',
    'skill_already_installed',
    'skill_not_installed',
    'skill_package_limit',
    'skill_package_timeout',
    'office_capability_disabled',
  ].includes(code)
    ? code
    : 'upload_failed'
}

interface SessionFile {
  name: string
  size: number
  arrayBuffer(): Promise<ArrayBuffer>
  text(): Promise<string>
}

export interface OfficeWorkspaceUi {
  readonly team?: PresentationTeamController
  readonly teamConnection?: OfficeTeamConnection
  readonly governance?: ReturnType<typeof createPresentationProjectGovernanceController>
  readonly research?: PresentationResearchController
  readonly project?: PresentationProjectController
  readonly importProgress?: PresentationImportProgressController
  readonly qa?: PresentationQaController
  readonly changes?: PresentationChangesController
  readonly interruptedChange?: { agentRunId: string; toolCallId: string }
  readonly durableAttachmentsAvailable?: () => boolean
  readonly durableImagesAvailable?: () => boolean
  readonly remoteImagesAvailable?: () => boolean
  readonly webpagesAvailable?: () => boolean
  readonly rightsAvailable?: () => boolean
  readonly animationFrameAvailable?: () => boolean
  readonly readPresentationAcquisitionHistory?: () => Promise<
    PresentationAcquisitionHistory | undefined
  >
  readonly listDurableAttachments?: () => Promise<PresentationAttachmentMetadata[]>
  readonly deleteDurableAttachment?: (attachmentId: string) => Promise<void>
  readonly importPresentationImageUrl?: (url: string | string[]) => Promise<void>
  readonly importPresentationWebpageUrl?: (url: string) => Promise<void>
  readonly attestPresentationImageLicense?: (
    imageId: string,
    license: 'owned' | 'licensed' | 'public_domain',
    evidenceId: string,
  ) => Promise<void>
  readonly revokePresentationImageLicense?: (imageId: string) => Promise<void>
  readonly extractPresentationImageFirstFrame?: (imageId: string) => Promise<void>
  readonly attachments: () => readonly string[]
  readonly downloadFile?: (path: string) => void
  readonly skills: () => readonly string[]
  readonly skillPackagesEnabled: boolean
  readonly upload: (file: SessionFile) => Promise<void>
  readonly copyDiagnostics?: () => Promise<void>
  readonly copyDiagnosticsWithContext?: () => Promise<void>
  readonly uninstallSkill?: (name: string) => void
  readonly clear: () => void
}

export function DiagnosticCopyButton(props: {
  copyDiagnostics: () => Promise<void>
  copyDiagnosticsWithContext?: () => Promise<void>
}): React.ReactElement {
  const [status, setStatus] = useState('')
  return (
    <>
      <button
        type="button"
        onClick={() => {
          setStatus('')
          void props
            .copyDiagnostics()
            .then(() => setStatus('诊断信息已复制；已移除文档、会话、项目和页面 ID'))
            .catch(() => setStatus('复制诊断信息失败'))
        }}
      >
        复制诊断信息
      </button>
      {props.copyDiagnosticsWithContext && (
        <button
          type="button"
          onClick={() => {
            setStatus('')
            void props
              .copyDiagnosticsWithContext?.()
              .then(() => setStatus('已复制含定位 ID 的本机诊断；请检查后分享'))
              .catch(() => setStatus('复制诊断信息失败'))
          }}
        >
          复制含定位 ID 的诊断
        </button>
      )}
      {status && (
        <p className="diagnostic-status" role="status">
          {status}
        </p>
      )}
    </>
  )
}

export function createOfficeWorkspaceUi(
  runtime: OfficeHostRuntime,
  diagnostics?: Pick<OfficeDiagnostics, 'exportJson'>,
  clipboard: { writeText(value: string): Promise<void> } | undefined = globalThis.navigator
    ?.clipboard,
  interruptedChange?: { agentRunId: string; toolCallId: string },
  teamConnection?: OfficeTeamConnection,
): OfficeWorkspaceUi {
  return Object.freeze({
    team: runtime.team,
    teamConnection,
    research: runtime.research,
    governance: runtime.governance,
    project: runtime.presentation,
    importProgress: runtime.importProgress,
    qa: runtime.qa,
    changes: runtime.changes,
    interruptedChange,
    durableAttachmentsAvailable: runtime.durableAttachmentsAvailable,
    durableImagesAvailable: runtime.durableImagesAvailable,
    remoteImagesAvailable: runtime.remoteImagesAvailable,
    webpagesAvailable: runtime.webpagesAvailable,
    rightsAvailable: runtime.rightsAvailable,
    animationFrameAvailable: runtime.animationFrameAvailable,
    readPresentationAcquisitionHistory: runtime.readPresentationAcquisitionHistory,
    listDurableAttachments: runtime.listDurableAttachments,
    deleteDurableAttachment: runtime.deleteDurableAttachment,
    importPresentationImageUrl: runtime.importPresentationImageUrl,
    importPresentationWebpageUrl: runtime.importPresentationWebpageUrl,
    attestPresentationImageLicense: runtime.attestPresentationImageLicense,
    revokePresentationImageLicense: runtime.revokePresentationImageLicense,
    extractPresentationImageFirstFrame: runtime.extractPresentationImageFirstFrame,
    attachments: () => Object.freeze([...runtime.vfs.list('/home/user')]),
    downloadFile: (path: string) => downloadSessionFile(runtime.vfs, path),
    skills: () => Object.freeze(runtime.skills.list().map((skill) => skill.name)),
    skillPackagesEnabled: runtime.skillPackagesEnabled,
    upload: (file: SessionFile) => uploadSessionFile(runtime, file),
    ...(diagnostics
      ? {
          copyDiagnostics: async () => {
            if (!clipboard || typeof clipboard.writeText !== 'function')
              throw new Error('diagnostic_copy_failed')
            await clipboard.writeText(diagnostics.exportJson())
          },
          copyDiagnosticsWithContext: async () => {
            if (!clipboard || typeof clipboard.writeText !== 'function')
              throw new Error('diagnostic_copy_failed')
            await clipboard.writeText(
              diagnostics.exportJson({
                includeLocalContext: true,
                ...(runtime.qa?.attempts
                  ? { screenshotAttempts: () => runtime.qa!.attempts!() }
                  : {}),
              }),
            )
          },
        }
      : {}),
    uninstallSkill: (name: string) => runtime.uninstallSkill(name),
    clear: () => runtime.clearSession(),
  })
}

export function uploadSessionFile(runtime: OfficeHostRuntime, file: SessionFile): Promise<void> {
  if (file.name.toLowerCase().endsWith('.zip')) {
    return runtime.installSkillPackage(file.arrayBuffer())
  }
  if (file.name === 'SKILL.md') {
    if (file.size > MAX_SKILL_BYTES) return Promise.reject(new Error('invalid_skill_package'))
    return runtime.installSkill(file.text())
  }
  if (
    runtime.durableAttachmentsAvailable?.() &&
    supportsPresentationAttachment(file.name, runtime.durableImagesAvailable?.())
  ) {
    const image = isPresentationImage(file.name)
    if (file.size > (image ? MAX_PRESENTATION_IMAGE_BYTES : MAX_PRESENTATION_ATTACHMENT_BYTES))
      return Promise.reject(
        new Error(image ? 'presentation_image_too_large' : 'presentation_attachment_too_large'),
      )
  } else if (file.size > MAX_VFS_FILE_BYTES) return Promise.reject(new Error('vfs_limit'))
  return runtime.uploadFile(file.name, file.arrayBuffer())
}

export function createPastedSourceFile(text: string, now = Date.now()): File {
  if (!text.trim() || text.length > 1_000_000) throw new Error('presentation_text_too_long')
  const name = `粘贴资料-${new Date(now).toISOString().replace(/[:.]/g, '-')}.txt`
  return new File([text], name, { type: 'text/plain;charset=utf-8' })
}

export type WorkspacePanelName = 'attachments' | 'skills'

export function composerKeyAction(event: {
  key: string
  shiftKey: boolean
  isComposing: boolean
}): 'send' | 'newline' | 'none' {
  if (event.key !== 'Enter') return 'none'
  return event.shiftKey || event.isComposing ? 'newline' : 'send'
}

interface FocusTarget {
  focus(): void
}

export function focusWorkspacePanel(heading: FocusTarget, opener: FocusTarget): () => void {
  heading.focus()
  return () => opener.focus()
}

export function isTimelineNearBottom(viewport: {
  scrollHeight: number
  scrollTop: number
  clientHeight: number
}): boolean {
  return viewport.scrollHeight - viewport.scrollTop - viewport.clientHeight <= 48
}

const starterPrompts: Record<OfficeHost, string[]> = {
  word: ['帮我写一份项目周报', '写一篇产品发布公告', '列一个活动策划提纲'],
  excel: ['分析这份表格的数据', '整理一份项目进度表', '找出数据中的异常'],
  powerpoint: ['起草一份项目汇报', '优化这份演示文稿', '列一个路演演示提纲'],
  unknown: ['帮我起草一份文档', '总结当前内容', '优化这份材料'],
}

function ProposalReview(props: {
  event: ProposalPresentationEvent
  activeProposalId?: string
  busy: boolean
  applying: boolean
  confirm: (id: string) => void
  reject: () => void
}) {
  const { event } = props
  const presentation = proposalPresentation(event.proposal)
  const hasComparison =
    isLegacyProposal(event.proposal) ||
    (event.proposal.before !== undefined && event.proposal.after !== undefined)
  const canReview = event.state === 'pending' && props.activeProposalId === event.proposal.id
  if (event.state === 'applied') {
    return (
      <article className="proposal-result" aria-label="Document change completed">
        <span className="proposal-result-icon" aria-hidden="true">
          ✓
        </span>
        <div>
          <h2>{presentation.title}</h2>
          {!isLegacyProposal(event.proposal) && <p>已更新 {presentation.count} 项</p>}
        </div>
      </article>
    )
  }
  return (
    <article
      className={`proposal-card proposal-${event.state}`}
      aria-label="Proposed document change"
    >
      <div className="proposal-heading">
        <span className="eyebrow">
          {event.state === 'pending'
            ? 'Approval required'
            : event.state === 'applying'
              ? 'Applying approved change'
              : event.state === 'uncertain'
                ? 'Write status uncertain'
                : event.state === 'rejected'
                  ? 'Change rejected'
                  : 'Change failed'}
        </span>
        <h2>{presentation.title}</h2>
      </div>
      {!isLegacyProposal(event.proposal) && (
        <div className="proposal-impact">
          <span>{presentation.host}</span>
          <span>{presentation.count} item(s)</span>
          <span>{presentation.targets.join(', ') || 'No named targets'}</span>
        </div>
      )}
      {(hasComparison || presentation.preview) && (
        <details className="proposal-preview" open>
          <summary>查看修改内容</summary>
          {hasComparison && (
            <div className="proposal-diff">
              <div className="preview-block">
                <strong>修改前</strong>
                <p className="proposal-copy">{presentation.before || '空白内容'}</p>
              </div>
              <div className="preview-block after">
                <strong>修改后</strong>
                <p className="proposal-copy">{presentation.after || '空白内容'}</p>
              </div>
            </div>
          )}
          {presentation.preview && <p className="proposal-copy">{presentation.preview}</p>}
        </details>
      )}
      {event.error && <p className="error-text">{event.error}</p>}
      {presentation.lockReview && (
        <div className="proposal-lock-review" role="status">
          {presentation.lockReview.state === 'checking' ? (
            <p>正在核对当前锁页，完成后可确认。</p>
          ) : presentation.lockReview.state === 'unavailable' ? (
            <p>锁页状态无法可靠核对，请拒绝此提案并在连接或页面身份恢复后重新生成。</p>
          ) : (
            presentation.lockReview.pages.length > 0 && (
              <>
                <p>本次修改可能影响以下锁定页面。确认即允许本次覆盖，页面保持锁定。</p>
                <ul>
                  {presentation.lockReview.pages.map((page) => (
                    <li key={`${page.projectId}/${page.pageId}`}>
                      {page.title}（{page.slideIds.length} 个宿主副本）
                    </li>
                  ))}
                </ul>
              </>
            )
          )}
        </div>
      )}
      {canReview && (
        <div className="actions">
          <button
            type="button"
            className="secondary"
            disabled={props.applying}
            onClick={props.reject}
          >
            Reject
          </button>
          <button
            type="button"
            disabled={
              props.applying ||
              (presentation.lockReview !== undefined && presentation.lockReview.state !== 'ready')
            }
            onClick={() => props.confirm(event.proposal.id)}
          >
            {props.applying
              ? 'Applying…'
              : presentation.lockReview?.state === 'ready' &&
                  presentation.lockReview.pages.length > 0
                ? '确认本次覆盖锁页'
                : 'Confirm change'}
          </button>
        </div>
      )}
      {event.state === 'applying' && <p className="proposal-state">Applying…</p>}
    </article>
  )
}

function TimelineEvent(props: {
  event: OfficePresentationEvent
  activeProposalId?: string
  busy: boolean
  applying: boolean
  confirm: (id: string) => void
  reject: () => void
  presentationParity?: boolean
}) {
  const { event } = props
  if (event.kind === 'proposal') return <ProposalReview {...props} event={event} />
  if (event.kind === 'tool') {
    return (
      <article className={`tool-event tool-${event.state}`} aria-label="Agent activity">
        <span className="tool-indicator" aria-hidden="true" />
        <p>{event.summary}</p>
      </article>
    )
  }
  if (props.presentationParity) {
    if (event.kind === 'system') {
      return <div className="ai-msg ai-msg-system">{event.text}</div>
    }
    const role = event.kind === 'user' ? 'user' : event.kind === 'error' ? 'error' : 'assistant'
    return (
      <PresentationMessage role={role} streaming={event.kind === 'assistant' && event.streaming}>
        {event.kind === 'assistant' ? <Markdown text={event.text} /> : event.text}
      </PresentationMessage>
    )
  }
  return (
    <article
      className={`timeline-message message-${event.kind}`}
      {...(event.kind === 'error' ? { role: 'alert' } : {})}
    >
      <span className="message-role">{event.kind === 'user' ? 'You' : 'WisWork'}</span>
      {event.kind === 'assistant' ? <Markdown text={event.text} /> : <p>{event.text}</p>}
      {event.kind === 'assistant' && event.streaming && (
        <span className="streaming-cursor" aria-label="Response streaming" />
      )}
    </article>
  )
}

export function presentationDesignLifecycle(output: string | undefined) {
  if (!output) return undefined
  try {
    const value = JSON.parse(output) as { status?: unknown; revision?: unknown }
    const status = typeof value.status === 'string' ? value.status : 'draft'
    const revision = typeof value.revision === 'number' ? value.revision : 1
    return {
      editable: status === 'draft' || status === 'ready',
      label:
        status === 'verified'
          ? 'DESIGN.md · 已验证'
          : status === 'ready' || status === 'producing'
            ? 'DESIGN.md · 已锁定'
            : revision > 1
              ? 'DESIGN.md · 已修订'
              : 'DESIGN.md · 已创建',
    }
  } catch {
    return { editable: true, label: 'DESIGN.md · 已创建' }
  }
}

function PowerPointTimeline(props: {
  timeline: OfficePresentationTimeline
  groupStages?: boolean
  latestDesignToolId?: string
  activeProposalId?: string
  busy: boolean
  applying: boolean
  activity: string
  confirm: (id: string) => void
  reject: () => void
  onOpenDesign: (designMd: string, editable: boolean) => void
}) {
  const toolDetail = (
    tool: Extract<OfficePresentationEvent, { kind: 'tool' }>,
  ): React.ReactNode | undefined => {
    if (tool.display?.kind === 'images' && tool.display.items?.length) {
      return (
        <div className="ai-tool-display-images">
          {tool.display.items.map((item) => (
            <a key={item.url} href={item.url} target="_blank" rel="noreferrer">
              {item.thumb?.startsWith('data:image/') ? (
                <img src={item.thumb} alt={item.title || ''} />
              ) : (
                item.title || item.url
              )}
            </a>
          ))}
        </div>
      )
    }
    if (tool.display?.kind === 'links' && tool.display.items?.length) {
      return (
        <ul className="ai-tool-display-links">
          {tool.display.items.map((item) => (
            <li key={item.url}>
              <a href={item.url} target="_blank" rel="noreferrer">
                {item.title || item.url}
              </a>
            </li>
          ))}
        </ul>
      )
    }
    const text = tool.display?.kind === 'text' ? tool.display.text : tool.output
    return text ? <pre className="ai-tool-display-text">{text}</pre> : undefined
  }
  const nodes: React.ReactNode[] = []
  const latestDesignToolId =
    props.latestDesignToolId ??
    [...props.timeline]
      .reverse()
      .find(
        (event) =>
          event.kind === 'tool' && Boolean(extractPresentationDesignDocument(event.output ?? '')),
      )?.id
  const timelineItems =
    props.groupStages === false ? props.timeline : presentationStageTimeline(props.timeline)
  const latestRequestIndex = timelineItems.reduce(
    (last, event, index) => (event.kind === 'user' ? index : last),
    -1,
  )
  let index = 0
  while (index < timelineItems.length) {
    const event = timelineItems[index]!
    if (event.kind === 'stage') {
      nodes.push(
        <PresentationStageCard
          key={event.id}
          group={event}
          runActive={(props.busy || props.applying) && index > latestRequestIndex}
        >
          <PowerPointTimeline
            {...props}
            timeline={event.events}
            groupStages={false}
            latestDesignToolId={latestDesignToolId}
          />
        </PresentationStageCard>,
      )
      index += 1
      continue
    }
    if (event.kind === 'phase') {
      nodes.push(
        <div className="ai-phase-row" key={event.id}>
          {event.text}…
        </div>,
      )
      index += 1
      continue
    }
    if (event.kind !== 'tool') {
      nodes.push(
        <TimelineEvent
          key={event.id}
          event={event}
          activeProposalId={props.activeProposalId}
          busy={props.busy}
          applying={props.applying}
          confirm={props.confirm}
          reject={props.reject}
          presentationParity
        />,
      )
      index += 1
      continue
    }
    const tools = []
    while (index < timelineItems.length && timelineItems[index]?.kind === 'tool') {
      const tool = timelineItems[index] as Extract<OfficePresentationEvent, { kind: 'tool' }>
      const designMd = extractPresentationDesignDocument(tool.output ?? '')
      const lifecycle = designMd ? presentationDesignLifecycle(tool.output) : undefined
      tools.push({
        id: tool.callId,
        label: lifecycle?.label ?? tool.summary,
        status:
          tool.state === 'running'
            ? ('running' as const)
            : tool.state === 'error'
              ? ('error' as const)
              : ('done' as const),
        ...(designMd
          ? {
              onActivate: () =>
                props.onOpenDesign(
                  designMd,
                  tool.id === latestDesignToolId && lifecycle?.editable === true,
                ),
            }
          : { detail: toolDetail(tool) }),
      })
      index += 1
    }
    nodes.push(
      <PresentationActivityGroup
        key={`tools:${tools[0]?.id}`}
        items={tools}
        workingLabel="处理中…"
        workedLabel={(count) => `已完成 · ${count} 个步骤`}
      />,
    )
  }
  const lastEvent = props.timeline.at(-1)
  const waitingForFirstAgentEvent = !lastEvent || lastEvent.kind === 'user'
  const streamingAssistant = lastEvent?.kind === 'assistant' && lastEvent.streaming
  if (props.busy && (waitingForFirstAgentEvent || streamingAssistant)) {
    nodes.push(
      <div className="ai-typing-row" key="active-agent-work">
        <AiTypingIndicator label={presentationProgressLabel(props.timeline)} />
      </div>,
    )
  }
  return <>{nodes}</>
}

function PowerPointQuestionnaire(props: {
  questions: readonly OfficeClarificationQuestion[]
  onSubmit: (answers: string) => void
  onSkip: () => void
}) {
  const [answers, setAnswers] = useState<Record<string, string>>({})
  const [activeIndex, setActiveIndex] = useState(0)
  const question = props.questions[activeIndex]
  const finish = (nextAnswers: Record<string, string>) =>
    props.onSubmit(
      props.questions
        .map((item) => `${item.label}: ${nextAnswers[item.id] || '（帮我决定）'}`)
        .join('\n'),
    )
  const advance = (delegated = false) => {
    if (!question) return
    const nextAnswers = delegated ? { ...answers, [question.id]: '' } : answers
    if (activeIndex + 1 < props.questions.length) {
      setAnswers(nextAnswers)
      setActiveIndex((current) => current + 1)
    } else finish(nextAnswers)
  }
  if (!question) return null
  return (
    <section className="ppt-questionnaire" aria-label="演示文稿制作问卷">
      <p className="ppt-questionnaire-progress">
        {activeIndex + 1} / {props.questions.length}
      </p>
      <label key={question.id}>
        <strong>{question.label}</strong>
        {question.description && <span>{question.description}</span>}
        <select
          value={answers[question.id] ?? ''}
          onChange={(event) =>
            setAnswers((current) => ({ ...current, [question.id]: event.target.value }))
          }
        >
          <option value="">帮我决定</option>
          {question.options.map((option) => (
            <option key={option} value={option}>
              {option}
            </option>
          ))}
        </select>
      </label>
      <div className="ppt-questionnaire-actions">
        <button type="button" className="secondary" onClick={() => advance(true)}>
          帮我决定
        </button>
        <button type="button" onClick={() => advance()}>
          {activeIndex + 1 < props.questions.length ? '下一题' : '继续制作'}
        </button>
      </div>
    </section>
  )
}

export function AgentWorkspace(props: {
  session: OfficeAgentSession
  ui: OfficeWorkspaceUi
  disconnect: () => void
  host: OfficeHost
  initialPanel?: WorkspacePanelName
  legacy?: boolean
  connectionNotice?: string
  runtimeMode?: 'standard' | 'enhanced'
  connectionAvailable?: boolean
  designRequest?: OfficeDesignRequest
  repairDesignConnection?: () => void | Promise<void>
}) {
  const { session, ui, disconnect, host } = props
  const state = useOfficeAgent(session)
  const projectPhase = useSyncExternalStore(
    (listener) => ui.project?.subscribe(listener) ?? (() => undefined),
    () => ui.project?.snapshot().phase ?? 'idle',
    () => ui.project?.snapshot().phase ?? 'idle',
  )
  const [instruction, setInstruction] = useState('')
  const [files, setFiles] = useState<readonly string[]>(ui.attachments())
  const [acquisitionRefresh, setAcquisitionRefresh] = useState(0)
  const [durableFiles, setDurableFiles] = useState<PresentationAttachmentMetadata[]>([])
  const [skills, setSkills] = useState<readonly string[]>(ui.skills())
  const [uploadError, setUploadError] = useState('')
  const [copyStatus, setCopyStatus] = useState('')
  const [copyPending, setCopyPending] = useState(false)
  const [uploadPending, setUploadPending] = useState(false)
  const [uploadStatus, setUploadStatus] = useState('')
  const [imageUrl, setImageUrl] = useState('')
  const [webpageUrl, setWebpageUrl] = useState('')
  const [pastedSource, setPastedSource] = useState('')
  const uploadEpoch = useRef(0)
  const [diagnosticStatus, setDiagnosticStatus] = useState('')
  const [designEditor, setDesignEditor] = useState<{
    markdown: string
    editable: boolean
  }>()
  const [designConversation, setDesignConversation] = useState<{
    session: OfficeAgentSession
    generation: number
    current?: { markdown: string; sourceId: string }
  }>({ session, generation: 0 })
  const [panel, setPanel] = useState<WorkspacePanelName | undefined>(props.initialPanel)
  const mounted = useRef(true)
  const panelHeading = useRef<HTMLHeadingElement>(null)
  const composerFileInput = useRef<HTMLInputElement>(null)
  const panelOpener = useRef<HTMLElement | undefined>(undefined)
  const timeline = useRef<HTMLElement>(null)
  const followLatest = useRef(true)
  useEffect(
    () => () => {
      mounted.current = false
    },
    [],
  )
  useEffect(() => {
    if (!ui.durableAttachmentsAvailable?.()) return
    void ui
      .listDurableAttachments?.()
      .then((items) => {
        if (mounted.current) setDurableFiles(items)
      })
      .catch(() => undefined)
  }, [ui])
  useEffect(() => {
    const heading = panelHeading.current
    const opener = panelOpener.current
    if (!panel || !heading || !opener) return
    return focusWorkspacePanel(heading, opener)
  }, [panel])
  useEffect(() => {
    const viewport = timeline.current
    if (!viewport || !followLatest.current || typeof viewport.scrollTo !== 'function') return
    viewport.scrollTo({ top: viewport.scrollHeight, behavior: state.busy ? 'auto' : 'smooth' })
  }, [state.busy, state.timeline])

  useEffect(() => {
    setFiles(ui.attachments())
  }, [ui, state.timeline])

  useEffect(() => {
    if (!state.busy) void ui.project?.refresh()
  }, [ui.project, state.busy])

  useEffect(() => {
    if (!state.busy) void ui.research?.refresh()
  }, [ui.research, state.busy])

  useEffect(() => ui.project?.subscribe(() => setFiles(ui.attachments())), [ui])

  function send() {
    if (
      props.connectionAvailable === false ||
      !instruction.trim() ||
      state.busy ||
      uploadPending ||
      state.applying ||
      state.proposal ||
      (ui.project && ui.project.snapshot().phase !== 'idle')
    )
      return
    session.send(instruction)
    setInstruction('')
  }

  async function uploadFiles(selected: FileList | readonly File[]) {
    if (
      uploadPending ||
      state.busy ||
      state.applying ||
      projectPhase !== 'idle' ||
      props.connectionAvailable === false
    )
      return
    const captured = ++uploadEpoch.current
    const current = () => mounted.current && captured === uploadEpoch.current
    setUploadError('')
    setUploadPending(true)
    setUploadStatus('正在上传…')
    try {
      for (const file of Array.from(selected)) {
        if (!current()) break
        try {
          await ui.upload(file)
          if (current()) setUploadStatus(`${file.name} 已上传，可让 Agent 读取。`)
        } catch (error) {
          if (current()) {
            setUploadStatus('')
            setUploadError(safeUploadError(error, file))
          }
          break
        }
      }
      if (current()) setFiles(ui.attachments())
      const items = await ui.listDurableAttachments?.().catch(() => undefined)
      if (current() && items) setDurableFiles(items)
    } finally {
      if (current()) setUploadPending(false)
    }
  }

  const proposal = state.proposal
  const latestDesign = [...state.timeline]
    .reverse()
    .find(
      (event) =>
        event.kind === 'tool' && Boolean(extractPresentationDesignDocument(event.output ?? '')),
    )
  const latestDesignDocument =
    latestDesign?.kind === 'tool'
      ? {
          markdown: extractPresentationDesignDocument(latestDesign.output ?? '')!,
          sourceId: latestDesign.id,
        }
      : undefined
  // The timeline is a rolling activity window, not the lifetime of DESIGN.md.
  // Reset synchronously on session replacement so a previous PC draft cannot
  // appear in the new session even when its timeline reuses the same event IDs.
  if (designConversation.session !== session) {
    setDesignConversation({
      session,
      generation: designConversation.generation + 1,
      current: latestDesignDocument,
    })
    setDesignEditor(undefined)
  } else if (
    latestDesignDocument &&
    (latestDesignDocument.sourceId !== designConversation.current?.sourceId ||
      latestDesignDocument.markdown !== designConversation.current?.markdown)
  ) {
    setDesignConversation({ ...designConversation, current: latestDesignDocument })
  }
  const currentDesign =
    latestDesignDocument ??
    (designConversation.session === session ? designConversation.current : undefined)
  const hasTimeline = state.timeline.length > 0
  const showConversationChrome =
    hasTimeline || state.busy || state.applying || Boolean(state.error) || Boolean(proposal)
  const showHeader = showConversationChrome || host === 'powerpoint'
  const showStatus =
    host !== 'powerpoint' &&
    (state.busy || state.applying || Boolean(state.activity) || state.status === 'cancelled')

  return (
    <main
      className={`agent-workspace ${host === 'powerpoint' ? 'presentation-agent ' : ''}${props.legacy ? 'legacy-workspace ' : ''}${panel ? 'has-management ' : ''}${showConversationChrome ? 'has-conversation' : `is-empty ${showHeader ? 'has-empty-header' : ''}`}`}
      aria-busy={state.busy || state.applying}
    >
      {showHeader && (
        <header className="app-header">
          <div className="editor-identity">
            <span className="connection-dot" aria-hidden="true" />
            <div>
              <h1>{agentProductLabels[host]}</h1>
              <p className="runtime-mode" aria-label="由 WisWork PC 管理的 Agent 模式">
                {props.runtimeMode === 'enhanced' ? '增强模式' : '标准模式'}
                <span>由 WisWork PC 管理</span>
              </p>
            </div>
            <span className="visually-hidden">Connected to WisWork PC</span>
          </div>
          <div className="header-actions">
            <button
              type="button"
              className="quiet"
              onClick={() => {
                uploadEpoch.current += 1
                setUploadPending(false)
                setUploadStatus('')
                session.newTask()
                ui.clear()
                setFiles([])
                setSkills([])
                setPanel(undefined)
                setUploadError('')
                setDesignEditor(undefined)
                setDesignConversation((current) => ({
                  session,
                  generation: current.generation + 1,
                }))
              }}
            >
              新对话
            </button>
            <details className="session-menu">
              <summary aria-label="Session menu">•••</summary>
              <button
                type="button"
                disabled={state.applying}
                onClick={(event) => {
                  panelOpener.current =
                    event.currentTarget.closest('details')?.querySelector('summary') ??
                    event.currentTarget
                  setPanel('skills')
                  event.currentTarget.closest('details')?.removeAttribute('open')
                }}
              >
                管理技能
              </button>
              {ui.copyDiagnostics && (
                <button
                  type="button"
                  onClick={(event) => {
                    setDiagnosticStatus('')
                    void ui
                      .copyDiagnostics?.()
                      .then(
                        () =>
                          mounted.current &&
                          setDiagnosticStatus('诊断信息已复制；已移除文档、会话、项目和页面 ID'),
                      )
                      .catch(() => mounted.current && setDiagnosticStatus('复制诊断信息失败'))
                    event.currentTarget.closest('details')?.removeAttribute('open')
                  }}
                >
                  复制诊断信息
                </button>
              )}
              {ui.copyDiagnosticsWithContext && (
                <button
                  type="button"
                  onClick={(event) => {
                    setDiagnosticStatus('')
                    void ui
                      .copyDiagnosticsWithContext?.()
                      .then(
                        () =>
                          mounted.current &&
                          setDiagnosticStatus('已复制含定位 ID 的本机诊断；请检查后分享'),
                      )
                      .catch(() => mounted.current && setDiagnosticStatus('复制诊断信息失败'))
                    event.currentTarget.closest('details')?.removeAttribute('open')
                  }}
                >
                  复制含定位 ID 的诊断
                </button>
              )}
              <button type="button" onClick={disconnect}>
                退出登录
              </button>
            </details>
          </div>
        </header>
      )}

      {diagnosticStatus && (
        <p className="diagnostic-status" role="status">
          {diagnosticStatus}
        </p>
      )}

      {props.connectionNotice && (
        <p className="diagnostic-status" role="status">
          {props.connectionNotice}
        </p>
      )}

      {showStatus && (
        <section className="agent-status" aria-live="polite">
          <span className={`status-dot ${state.busy || state.applying ? 'busy' : ''}`} />
          <strong>
            {state.applying
              ? 'Applying approved change'
              : state.busy
                ? 'Agent is working'
                : 'Agent is ready'}
          </strong>
          <span>{state.activity || (state.status === 'cancelled' ? 'Run stopped' : '')}</span>
        </section>
      )}

      {host === 'powerpoint' && (
        <OfficeDesignPanel
          key={designConversation.generation}
          current={currentDesign}
          selection={designEditor}
          busy={state.busy || state.applying || props.connectionAvailable === false}
          request={props.designRequest}
          onRepairConnection={props.repairDesignConnection}
          onApply={(markdown) => session.reviseDesignContract?.(markdown)}
        />
      )}

      <section
        ref={timeline}
        className="agent-timeline"
        aria-label="Agent conversation"
        aria-live="polite"
        onScroll={(event) => {
          followLatest.current = isTimelineNearBottom(event.currentTarget)
        }}
      >
        {!hasTimeline && host === 'powerpoint' && (
          <PresentationEmptyState
            title="让 AI 为你生成演示文稿"
            body="描述主题、场合和大致页数，AI 直接生成整份幻灯片。"
            prompts={starterPrompts.powerpoint}
            onChoose={setInstruction}
          />
        )}
        {!hasTimeline && host !== 'powerpoint' && (
          <div className="empty-state">
            <h2>让 AI 帮你从零起草</h2>
            <p>
              描述主题、要点或粘贴参考素材，
              <br />
              AI 直接为你写出初稿。
            </p>
            <div className="starter-prompts">
              {starterPrompts[host].map((prompt) => (
                <button
                  key={prompt}
                  type="button"
                  className="prompt-chip"
                  onClick={() => setInstruction(prompt)}
                >
                  {prompt}
                </button>
              ))}
            </div>
          </div>
        )}
        {host === 'powerpoint' ? (
          <PowerPointTimeline
            timeline={state.timeline}
            activeProposalId={proposal?.id}
            busy={state.busy}
            applying={state.applying}
            activity={state.activity}
            confirm={(id) => void session.confirm(id)}
            reject={() => session.reject()}
            onOpenDesign={(designMd, editable) => {
              setDesignEditor({ markdown: designMd, editable })
            }}
          />
        ) : (
          state.timeline.map((event) => (
            <TimelineEvent
              key={event.id}
              event={event}
              activeProposalId={proposal?.id}
              busy={state.busy}
              applying={state.applying}
              confirm={(id) => void session.confirm(id)}
              reject={() => session.reject()}
            />
          ))
        )}
        {state.questionnaire && (
          <PowerPointQuestionnaire
            questions={state.questionnaire}
            onSubmit={(answers) => session.answerQuestionnaire?.(answers)}
            onSkip={() => session.skipQuestionnaire?.()}
          />
        )}
        {state.recoveryAvailable && (
          <button
            type="button"
            className="secondary"
            disabled={state.busy || state.applying || projectPhase !== 'idle'}
            onClick={() => void session.resumeInterrupted?.()}
          >
            恢复上次任务
          </button>
        )}
        {proposal &&
          !state.timeline.some(
            (event) => event.kind === 'proposal' && event.proposal.id === proposal.id,
          ) && (
            <ProposalReview
              event={{
                id: `proposal-${proposal.id}`,
                kind: 'proposal',
                proposal,
                state: 'pending',
              }}
              activeProposalId={proposal.id}
              busy={state.busy}
              applying={state.applying}
              confirm={(id) => void session.confirm(id)}
              reject={() => session.reject()}
            />
          )}
        {state.error && state.errorMessage && (
          <div className="error-banner" role="alert">
            <div>
              <p>{state.errorMessage}</p>
            </div>
            {state.retryable && (
              <button
                type="button"
                className="secondary"
                disabled={
                  uploadPending ||
                  state.applying ||
                  state.busy ||
                  Boolean(state.proposal) ||
                  projectPhase !== 'idle'
                }
                onClick={() => {
                  if (!uploadPending && (!ui.project || ui.project.snapshot().phase === 'idle'))
                    session.retry()
                }}
              >
                {state.error === 'proposal_stale' ? '重新生成' : 'Retry'}
              </button>
            )}
          </div>
        )}
      </section>

      {panel && (
        <section
          className="management-panel"
          role="dialog"
          aria-modal="false"
          aria-labelledby="panel-title"
          onKeyDown={(event) => {
            if (event.key === 'Escape') setPanel(undefined)
          }}
        >
          <div className="panel-heading">
            <h2 id="panel-title" ref={panelHeading} tabIndex={-1}>
              {panel === 'attachments' ? 'Session attachments' : 'Installed skills'}
            </h2>
            <button
              type="button"
              className="icon-button"
              aria-label="Close panel"
              onClick={() => setPanel(undefined)}
            >
              ×
            </button>
          </div>
          {panel === 'attachments' ? (
            <>
              <label
                className="upload-button"
                htmlFor="session-upload"
                aria-disabled={
                  state.applying || state.busy || uploadPending || projectPhase !== 'idle'
                }
              >
                Add attachment
              </label>
              <input
                id="session-upload"
                className="visually-hidden"
                type="file"
                disabled={state.applying || state.busy || uploadPending || projectPhase !== 'idle'}
                onChange={(event) => {
                  const file = event.currentTarget.files?.[0]
                  event.currentTarget.value = ''
                  if (
                    !file ||
                    uploadPending ||
                    state.busy ||
                    state.applying ||
                    projectPhase !== 'idle'
                  )
                    return
                  const captured = ++uploadEpoch.current
                  const durable =
                    ui.durableAttachmentsAvailable?.() &&
                    supportsPresentationAttachment(file.name, ui.durableImagesAvailable?.()) &&
                    file.name !== 'SKILL.md'
                  const current = () => mounted.current && captured === uploadEpoch.current
                  setUploadError('')
                  setUploadStatus(durable ? '正在上传到 PC 并解析…' : '正在上传…')
                  setUploadPending(true)
                  void ui
                    .upload(file)
                    .then(() => {
                      if (!current()) return
                      setFiles(ui.attachments())
                      void ui
                        .listDurableAttachments?.()
                        .then((items) => {
                          if (current()) setDurableFiles(items)
                        })
                        .catch(() => undefined)
                      setUploadStatus(
                        durable
                          ? `${file.name} 已保存到 PC 并解析，可让 Agent 读取。`
                          : `${file.name} 已加入本次会话。`,
                      )
                    })
                    .catch((error: unknown) => {
                      if (!current()) return
                      setUploadStatus('')
                      setUploadError(safeUploadError(error, file))
                      void ui
                        .listDurableAttachments?.()
                        .then((items) => {
                          if (current()) setDurableFiles(items)
                        })
                        .catch(() => undefined)
                    })
                    .finally(() => {
                      if (current()) setUploadPending(false)
                    })
                }}
              />
              <p>
                {ui.durableAttachmentsAvailable?.()
                  ? 'PDF、Word（DOCX）、HTML、TXT、MD、CSV、JSON 资料每个最多 50 MiB，保存于 PC 并绑定当前文档；退出登录不会删除。重连后可让 Agent 列出和读取，重新选择同一文件可续传。'
                  : `Files are limited to ${displayMebibytes(MAX_VFS_FILE_BYTES)} MiB each and ${displayMebibytes(MAX_VFS_TOTAL_BYTES)} MiB per session, then cleared on logout.`}
              </p>
              {ui.durableAttachmentsAvailable?.() && (
                <form
                  onSubmit={(event) => {
                    event.preventDefault()
                    if (uploadPending || state.busy || state.applying || projectPhase !== 'idle')
                      return
                    let file: File
                    try {
                      file = createPastedSourceFile(pastedSource)
                    } catch (error) {
                      setUploadError(safeUploadError(error))
                      return
                    }
                    const captured = ++uploadEpoch.current
                    const current = () => mounted.current && captured === uploadEpoch.current
                    setUploadPending(true)
                    setUploadError('')
                    setUploadStatus('正在保存粘贴资料到 PC…')
                    void ui
                      .upload(file)
                      .then(() => {
                        if (!current()) return
                        setPastedSource('')
                        setUploadStatus('粘贴资料已保存到当前文档，可让 Agent 读取。')
                        void ui
                          .listDurableAttachments?.()
                          .then((items) => {
                            if (current()) setDurableFiles(items)
                          })
                          .catch(() => undefined)
                      })
                      .catch((error: unknown) => {
                        if (current()) {
                          setUploadStatus('')
                          setUploadError(safeUploadError(error, file))
                        }
                      })
                      .finally(() => {
                        if (current()) setUploadPending(false)
                      })
                  }}
                >
                  <label htmlFor="presentation-pasted-source">粘贴资料原文</label>
                  <textarea
                    id="presentation-pasted-source"
                    value={pastedSource}
                    onChange={(event) => setPastedSource(event.currentTarget.value)}
                    rows={5}
                    maxLength={1_000_000}
                    required
                  />
                  <button
                    type="submit"
                    disabled={
                      uploadPending ||
                      state.busy ||
                      state.applying ||
                      projectPhase !== 'idle' ||
                      !pastedSource.trim()
                    }
                  >
                    保存粘贴资料
                  </button>
                </form>
              )}
              {ui.webpagesAvailable?.() && (
                <form
                  onSubmit={(event) => {
                    event.preventDefault()
                    if (uploadPending || state.busy || !webpageUrl.trim()) return
                    const captured = ++uploadEpoch.current
                    const current = () => mounted.current && captured === uploadEpoch.current
                    setUploadPending(true)
                    setUploadError('')
                    setUploadStatus('正在由 PC 抓取网页并保存原文…')
                    void ui
                      .importPresentationWebpageUrl?.(webpageUrl.trim())
                      .then(() => {
                        if (!current()) return
                        setUploadStatus('网页原文及正文已保存到当前文档的 PC 资料。')
                        setWebpageUrl('')
                        void ui
                          .listDurableAttachments?.()
                          .then((items) => {
                            if (current()) setDurableFiles(items)
                          })
                          .catch(() => undefined)
                      })
                      .catch((error: unknown) => {
                        if (current()) {
                          setUploadStatus('')
                          setUploadError(safeUploadError(error))
                        }
                      })
                      .finally(() => {
                        if (current()) {
                          setUploadPending(false)
                          setAcquisitionRefresh((value) => value + 1)
                        }
                      })
                  }}
                >
                  <label htmlFor="presentation-webpage-url">网页网址</label>
                  <input
                    id="presentation-webpage-url"
                    type="url"
                    value={webpageUrl}
                    onChange={(event) => setWebpageUrl(event.currentTarget.value)}
                    placeholder="https://example.com/article"
                    required
                  />
                  <button
                    type="submit"
                    disabled={uploadPending || state.busy || !webpageUrl.trim()}
                  >
                    由 PC 获取网页
                  </button>
                </form>
              )}
              {ui.durableImagesAvailable?.() && (
                <>
                  <p>
                    PNG、JPEG、静态 GIF、WebP 图片每个最多 10 MiB，上传后在 PC
                    校验并缓存；可直接用于制作，无需将图片编码发给 Agent。
                  </p>
                  {ui.remoteImagesAvailable?.() && (
                    <form
                      onSubmit={(event) => {
                        event.preventDefault()
                        if (uploadPending || state.busy || !imageUrl.trim()) return
                        const urls = imageUrl
                          .split(/\r?\n/)
                          .map((url) => url.trim())
                          .filter(Boolean)
                        if (urls.length > 4) {
                          setUploadError('候选图片网址最多 4 个。')
                          return
                        }
                        const captured = ++uploadEpoch.current
                        const current = () => mounted.current && captured === uploadEpoch.current
                        setUploadPending(true)
                        setUploadError('')
                        setUploadStatus('正在由 PC 下载并校验图片…')
                        void ui
                          .importPresentationImageUrl?.(urls.length === 1 ? urls[0]! : urls)
                          .then(async () => {
                            if (!current()) return
                            const items = await ui.listDurableAttachments?.()
                            if (!current()) return
                            if (items) setDurableFiles(items)
                            setUploadStatus(
                              '图片已保存到当前文档的 PC 素材缓存；许可状态仍需核验。',
                            )
                            setImageUrl('')
                          })
                          .catch((error: unknown) => {
                            if (current()) {
                              setUploadStatus('')
                              setUploadError(safeUploadError(error))
                              void ui
                                .listDurableAttachments?.()
                                .then((items) => {
                                  if (current()) setDurableFiles(items)
                                })
                                .catch(() => undefined)
                            }
                          })
                          .finally(() => {
                            if (current()) {
                              setUploadPending(false)
                              setAcquisitionRefresh((value) => value + 1)
                            }
                          })
                      }}
                    >
                      <label htmlFor="presentation-image-url">
                        图片网址（每行一个，最多 4 个，按顺序尝试）
                      </label>
                      <textarea
                        id="presentation-image-url"
                        value={imageUrl}
                        onChange={(event) => setImageUrl(event.currentTarget.value)}
                        placeholder={
                          'https://example.com/image.png\nhttps://backup.example.com/image.webp'
                        }
                        rows={3}
                        required
                      />
                      <button
                        type="submit"
                        disabled={uploadPending || state.busy || !imageUrl.trim()}
                      >
                        由 PC 获取图片
                      </button>
                    </form>
                  )}
                </>
              )}
              {ui.durableAttachmentsAvailable?.() && (
                <p>
                  下方仅显示本次会话可下载的副本；超过 20 MiB 或会话容量的资料仍可由 Agent 在 PC
                  读取。其他文件及技能仅保留在会话中。
                </p>
              )}
              {uploadStatus && <p role="status">{uploadStatus}</p>}
              {uploadError && (
                <p className="error-text" role="alert">
                  {uploadError}
                </p>
              )}
              {ui.durableAttachmentsAvailable?.() && (
                <PresentationAcquisitionHistoryCard
                  read={ui.readPresentationAcquisitionHistory}
                  refreshKey={`${acquisitionRefresh}:${state.busy}`}
                />
              )}
              {ui.durableAttachmentsAvailable?.() && durableFiles.length > 0 && (
                <ul>
                  {durableFiles.map((file) => (
                    <li key={file.attachmentId}>
                      {file.name} · {file.status}
                      {file.sectionCount !== undefined && (
                        <span>
                          {' · '}
                          {file.name.toLowerCase().endsWith('.pdf') ? 'PDF' : '资料'}{' '}
                          {file.sectionCount}{' '}
                          {file.name.toLowerCase().endsWith('.pdf') ? '页' : '段'}
                        </span>
                      )}
                      {file.pagesWithoutExtractedText?.length ? (
                        <p role="status">
                          第 {file.pagesWithoutExtractedText.slice(0, 20).join('、')}
                          {file.pagesWithoutExtractedText.length > 20
                            ? ` 等 ${file.pagesWithoutExtractedText.length} `
                            : ' '}
                          页未提取到文字；如这些页面包含内容，请补充可读取文本或 OCR 结果。
                        </p>
                      ) : null}
                      {file.pagesWithSparseExtractedText?.length ? (
                        <p role="status">
                          第 {file.pagesWithSparseExtractedText.slice(0, 20).join('、')}
                          {file.pagesWithSparseExtractedText.length > 20
                            ? ` 等 ${file.pagesWithSparseExtractedText.length} `
                            : ' '}
                          页提取到的文字较少；如关键内容在这些页面，请核对原 PDF
                          或补充可读取文本。此提示不判断 OCR 准确性。
                        </p>
                      ) : null}
                      {file.pagesWithFullPageImage?.length ? (
                        <p role="status">
                          第 {file.pagesWithFullPageImage.slice(0, 20).join('、')}
                          {file.pagesWithFullPageImage.length > 20
                            ? ` 等 ${file.pagesWithFullPageImage.length} `
                            : ' '}
                          页含覆盖整页的图像；请核对提取文字与原 PDF。
                        </p>
                      ) : null}
                      {file.pagesWithInvisibleTextLayer?.length ? (
                        <p role="status">
                          第 {file.pagesWithInvisibleTextLayer.slice(0, 20).join('、')}
                          {file.pagesWithInvisibleTextLayer.length > 20
                            ? ` 等 ${file.pagesWithInvisibleTextLayer.length} `
                            : ' '}
                          页含不可见文字层；该文字未经核对，不能仅凭匹配结果支持主张，请补充经核对文本。
                        </p>
                      ) : null}
                      {file.status === 'failed' && file.error === 'parse_failed' && (
                        <p role="alert">
                          资料解析失败。
                          {file.name.toLowerCase().endsWith('.pdf')
                            ? '如 PDF 是扫描件，请补充可读取文本或 OCR 结果后重新上传。'
                            : '请检查文件格式或重新上传。'}
                        </p>
                      )}
                      {file.animationHandling === 'first_frame' && (
                        <p>已按你的选择从动画原件生成静态首帧；原件仍保存在 PC。</p>
                      )}
                      {file.status === 'failed' &&
                        file.error === 'animated_image_unsupported' &&
                        ui.animationFrameAvailable?.() && (
                          <button
                            type="button"
                            disabled={uploadPending || state.busy}
                            onClick={() => {
                              setUploadPending(true)
                              setUploadError('')
                              void ui
                                .extractPresentationImageFirstFrame?.(file.attachmentId)
                                .then(async () => {
                                  if (!mounted.current) return
                                  setDurableFiles((await ui.listDurableAttachments?.()) ?? [])
                                  setUploadStatus(
                                    `${file.name} 的静态首帧已生成，可作为图片素材使用。`,
                                  )
                                })
                                .catch((error: unknown) => {
                                  if (mounted.current) setUploadError(safeUploadError(error))
                                })
                                .finally(() => {
                                  if (mounted.current) setUploadPending(false)
                                })
                            }}
                          >
                            生成静态首帧
                          </button>
                        )}
                      {file.kind === 'image' && file.source && (
                        <details>
                          <summary>
                            来源 {(file.sources ?? [file.source]).length} 条 · 许可未核验
                          </summary>
                          <ul>
                            {(file.sources ?? [file.source]).map((source, index) => (
                              <li key={`${index}:${source}`}>{source}</li>
                            ))}
                          </ul>
                        </details>
                      )}
                      {file.kind === 'text' && file.source && (
                        <p>
                          网页来源：{file.source}；抓取时间：
                          {file.retrievedAt
                            ? new Date(file.retrievedAt).toLocaleString()
                            : '未记录'}
                        </p>
                      )}
                      {file.kind === 'image' && file.licenseDeclaration && (
                        <p>
                          使用权：用户声明 {file.licenseDeclaration.kind}；依据附件{' '}
                          {durableFiles.find(
                            (item) =>
                              item.attachmentId === file.licenseDeclaration?.evidenceAttachmentId,
                          )?.name ??
                            `${file.licenseDeclaration.evidenceAttachmentId.slice(0, 12)}…`}
                          ；尚未核验。
                        </p>
                      )}
                      {file.kind === 'image' && ui.rightsAvailable?.() && (
                        <details>
                          <summary>管理使用权声明</summary>
                          <form
                            onSubmit={(event) => {
                              event.preventDefault()
                              const data = new FormData(event.currentTarget)
                              const license = data.get('license')
                              const evidenceId = data.get('evidence')
                              if (
                                !['owned', 'licensed', 'public_domain'].includes(String(license)) ||
                                typeof evidenceId !== 'string' ||
                                !evidenceId
                              )
                                return
                              setUploadPending(true)
                              setUploadError('')
                              void ui
                                .attestPresentationImageLicense?.(
                                  file.attachmentId,
                                  license as 'owned' | 'licensed' | 'public_domain',
                                  evidenceId,
                                )
                                .then(async () => {
                                  if (mounted.current) {
                                    setDurableFiles((await ui.listDurableAttachments?.()) ?? [])
                                    setUploadStatus('图片使用权声明已保存，仍需人工核验依据。')
                                  }
                                })
                                .catch((error: unknown) => {
                                  if (mounted.current) setUploadError(safeUploadError(error))
                                })
                                .finally(() => {
                                  if (mounted.current) setUploadPending(false)
                                })
                            }}
                          >
                            <label>
                              使用权声明
                              <select name="license" defaultValue="licensed">
                                <option value="owned">自有</option>
                                <option value="licensed">已获许可</option>
                                <option value="public_domain">公有领域</option>
                              </select>
                            </label>
                            <label>
                              依据附件
                              <select name="evidence" required defaultValue="">
                                <option value="" disabled>
                                  选择已上传的资料
                                </option>
                                {durableFiles
                                  .filter((item) => item.kind === 'text' && item.status === 'ready')
                                  .map((item) => (
                                    <option key={item.attachmentId} value={item.attachmentId}>
                                      {item.name}
                                    </option>
                                  ))}
                              </select>
                            </label>
                            <button
                              type="submit"
                              disabled={
                                uploadPending ||
                                state.busy ||
                                !durableFiles.some(
                                  (item) => item.kind === 'text' && item.status === 'ready',
                                )
                              }
                            >
                              保存声明
                            </button>
                          </form>
                        </details>
                      )}
                      {file.kind === 'image' &&
                        file.licenseDeclaration &&
                        ui.rightsAvailable?.() && (
                          <button
                            type="button"
                            disabled={uploadPending || state.busy}
                            onClick={() => {
                              setUploadPending(true)
                              void ui
                                .revokePresentationImageLicense?.(file.attachmentId)
                                .then(async () => {
                                  if (mounted.current)
                                    setDurableFiles((await ui.listDurableAttachments?.()) ?? [])
                                })
                                .catch((error: unknown) => {
                                  if (mounted.current) setUploadError(safeUploadError(error))
                                })
                                .finally(() => {
                                  if (mounted.current) setUploadPending(false)
                                })
                            }}
                          >
                            撤回声明
                          </button>
                        )}
                      <button
                        type="button"
                        disabled={uploadPending || state.busy}
                        onClick={() => {
                          if (
                            !window.confirm(
                              `删除 PC 中的“${file.name}”？此操作会移除原文件及解析结果，之后使用需重新上传。`,
                            )
                          )
                            return
                          setUploadPending(true)
                          void ui
                            .deleteDurableAttachment?.(file.attachmentId)
                            .then(async () => {
                              if (!mounted.current) return
                              setDurableFiles((await ui.listDurableAttachments?.()) ?? [])
                              setUploadStatus(`${file.name} 已从 PC 删除。`)
                            })
                            .catch((error: unknown) => {
                              if (mounted.current) setUploadError(safeUploadError(error))
                            })
                            .finally(() => {
                              if (mounted.current) setUploadPending(false)
                            })
                        }}
                      >
                        删除 PC 副本
                      </button>
                    </li>
                  ))}
                </ul>
              )}
              <ul>
                {files.map((file) => (
                  <li key={file}>
                    {ui.downloadFile ? (
                      <button type="button" onClick={() => ui.downloadFile?.(file)}>
                        {sessionAttachmentLabel(file)} · 下载
                      </button>
                    ) : (
                      sessionAttachmentLabel(file)
                    )}
                  </li>
                ))}
              </ul>
              {!files.length && <p>No session attachments.</p>}
            </>
          ) : (
            <>
              {ui.skillPackagesEnabled && (
                <>
                  <label
                    className="upload-button"
                    htmlFor="skill-package-upload"
                    aria-disabled={state.applying}
                  >
                    Install skill package
                  </label>
                  <input
                    id="skill-package-upload"
                    className="visually-hidden"
                    type="file"
                    accept=".zip,application/zip"
                    disabled={state.applying}
                    onChange={(event) => {
                      const file = event.currentTarget.files?.[0]
                      if (!file) return
                      setUploadError('')
                      void ui
                        .upload(file)
                        .then(() => mounted.current && setSkills(ui.skills()))
                        .catch(
                          (error: unknown) =>
                            mounted.current && setUploadError(safeUploadError(error)),
                        )
                    }}
                  />
                </>
              )}
              <p>Installed instructions can guide the Agent but cannot add Office authority.</p>
              {uploadError && (
                <p className="error-text" role="alert">
                  {uploadError}
                </p>
              )}
              <ul>
                {skills.map((skill) => (
                  <li key={skill}>
                    <span>{skill}</span>
                    <button
                      type="button"
                      className="quiet"
                      disabled={state.applying}
                      onClick={() => {
                        try {
                          ui.uninstallSkill?.(skill)
                          setSkills(ui.skills())
                        } catch (error) {
                          setUploadError(safeUploadError(error))
                        }
                      }}
                    >
                      Remove
                    </button>
                  </li>
                ))}
              </ul>
              {!skills.length && <p>No installed skills.</p>}
            </>
          )}
        </section>
      )}

      <section className="composer-shell" aria-label="Message WisWork Agent">
        {host === 'powerpoint' && ui.project && (
          <section aria-label="制作位置">
            <strong>制作位置</strong>
            <p>
              可以在当前文档制作。已有内容时建议先保存原稿、再创建副本。现有项目不会自动迁移到副本。
            </p>
            <button
              type="button"
              disabled={
                copyPending ||
                uploadPending ||
                state.busy ||
                state.applying ||
                projectPhase !== 'idle' ||
                !supportsPowerPointPresentationCopy()
              }
              onClick={() => {
                setCopyPending(true)
                setCopyStatus('正在创建副本…')
                void openPowerPointPresentationCopy()
                  .then(() =>
                    setCopyStatus(
                      '副本已打开。请先另存为新文件，再在副本中打开 WisWork 并开始制作。',
                    ),
                  )
                  .catch((error: unknown) =>
                    setCopyStatus(
                      error instanceof Error && error.message === 'presentation_document_changed'
                        ? '导出期间文档已切换，副本未打开。请回到原文档重试。'
                        : error instanceof Error &&
                            error.message === 'presentation_copy_save_source_first'
                          ? '请先保存当前文档，再创建副本；未保存的两份文稿可能共用项目身份。'
                          : '创建副本失败；请检查文档身份、导出权限和 PowerPoint 版本。',
                    ),
                  )
                  .finally(() => setCopyPending(false))
              }}
            >
              创建副本后制作
            </button>
            {!supportsPowerPointPresentationCopy() && <p>当前 PowerPoint 不支持创建副本。</p>}
            {copyStatus && <p role="status">{copyStatus}</p>}
          </section>
        )}
        {host === 'powerpoint' && ui.team && (
          <PresentationTeamCard
            controller={ui.team}
            account={ui.teamConnection}
            disabled={uploadPending || state.busy || state.applying || Boolean(state.proposal)}
          />
        )}
        {host === 'powerpoint' &&
          import.meta.env.VITE_WISWORK_PPT_PROJECT_GOVERNANCE_ENABLED === '1' &&
          ui.governance && (
            <PresentationProjectGovernanceCard
              controller={ui.governance}
              disabled={uploadPending || state.busy || state.applying || Boolean(state.proposal)}
            />
          )}
        {ui.research && (
          <PresentationResearchCard
            controller={ui.research}
            disabled={uploadPending || state.busy || state.applying || Boolean(state.proposal)}
          />
        )}
        {ui.project && (
          <PresentationWorkflowCard
            project={ui.project}
            imported={ui.importProgress}
            qa={ui.qa}
            disabled={uploadPending || state.busy || state.applying || Boolean(state.proposal)}
          />
        )}
        {ui.project && (
          <PresentationProjectCard
            controller={ui.project}
            onEndFrontend={() => session.stop()}
            disabled={uploadPending || state.busy || state.applying || Boolean(state.proposal)}
          />
        )}
        {ui.importProgress && <PresentationImportProgressCard controller={ui.importProgress} />}
        {ui.changes && (
          <PresentationChangesCard
            controller={ui.changes}
            interruptedChange={ui.interruptedChange}
            disabled={
              uploadPending ||
              state.busy ||
              state.applying ||
              Boolean(state.proposal) ||
              projectPhase !== 'idle'
            }
          />
        )}
        {ui.qa && (
          <PresentationQaCard
            controller={ui.qa}
            disabled={
              uploadPending ||
              state.busy ||
              state.applying ||
              Boolean(state.proposal) ||
              projectPhase !== 'idle'
            }
            onRecheck={(record, pageIds) => {
              if (
                uploadPending ||
                state.busy ||
                state.applying ||
                state.proposal ||
                projectPhase !== 'idle' ||
                !validatePresentationQaRecord(record) ||
                pageIds.length === 0 ||
                new Set(pageIds).size !== pageIds.length ||
                pageIds.some((id) => !record.pages.some((page) => page.pageId === id))
              )
                return
              setInstruction(
                `请准备页面 QA 重审：projectId=${JSON.stringify(record.projectId)}，requestId=${JSON.stringify(record.requestId)}，pageIds=${JSON.stringify(pageIds)}。` +
                  '先确认当前文档、目标冻结任务及宿主页面映射一致；不一致时停止并说明。仅对指定页面逐页重新采集实际截图，观察截图后由 Agent 进行视觉复核；不得沿用历史通过结论或将采集成功视为通过。需要修改时先生成提案并等待用户确认。',
              )
            }}
          />
        )}
        {ui.downloadFile &&
          files.some(
            (file) =>
              file.startsWith('/home/user/generated/') &&
              (file.endsWith('.pptx') || file.endsWith('.pdf')),
          ) && (
            <section aria-label="生成的演示文稿" className="presentation-downloads">
              {files
                .filter(
                  (file) =>
                    file.startsWith('/home/user/generated/') &&
                    (file.endsWith('.pptx') ||
                      file.endsWith('.pdf') ||
                      file.endsWith('.report.json') ||
                      file.endsWith('/report.json')),
                )
                .slice(-4)
                .map((file) => (
                  <button type="button" key={file} onClick={() => ui.downloadFile?.(file)}>
                    {file.endsWith('.pptx')
                      ? '下载 PPTX'
                      : file.endsWith('.pdf')
                        ? '下载 PDF 预览'
                        : '下载验收报告'}{' '}
                    · {sessionAttachmentLabel(file)}
                  </button>
                ))}
            </section>
          )}
        <div className="composer-input-box">
          {files.length > 0 && (
            <div className="composer-attachments" aria-label="Attached files">
              {files.map((file) => (
                <span key={file}>{file.split('/').at(-1)}</span>
              ))}
            </div>
          )}
          <label className="visually-hidden" htmlFor="instruction">
            Message WisWork Agent
          </label>
          <textarea
            id="instruction"
            value={instruction}
            onChange={(event) => setInstruction(event.target.value)}
            onKeyDown={(event) => {
              if (
                composerKeyAction({
                  key: event.key,
                  shiftKey: event.shiftKey,
                  isComposing: event.nativeEvent.isComposing,
                }) === 'send'
              ) {
                event.preventDefault()
                send()
              }
            }}
            placeholder={
              host === 'powerpoint'
                ? '描述要生成的演示文稿，或直接提问'
                : '描述修改、写作要求，或直接提问'
            }
            rows={3}
            maxLength={12_000}
            disabled={props.connectionAvailable === false || state.busy || state.applying}
          />
          <div className="composer-toolbar">
            <div className="composer-tools">
              <input
                ref={composerFileInput}
                id="composer-attachment-upload"
                className="visually-hidden"
                type="file"
                multiple
                disabled={state.applying || state.busy || uploadPending || projectPhase !== 'idle'}
                onChange={(event) => {
                  const selected = event.currentTarget.files
                  if (selected?.length) void uploadFiles(selected)
                  event.currentTarget.value = ''
                }}
              />
              <button
                type="button"
                className="composer-attach-button"
                aria-label="Add attachments"
                disabled={state.applying || state.busy || uploadPending || projectPhase !== 'idle'}
                onClick={() => composerFileInput.current?.click()}
              >
                <IconPaperclip size={20} />
              </button>
              <button
                type="button"
                className="icon-button"
                aria-label="Attachments"
                aria-expanded={panel === 'attachments'}
                disabled={state.applying}
                onClick={(event) => {
                  panelOpener.current = event.currentTarget
                  setPanel(panel === 'attachments' ? undefined : 'attachments')
                }}
              >
                管理附件
              </button>
              <span className="confirmation-chip">
                <span aria-hidden="true" />
                {host === 'powerpoint' && !ui.project ? '自动应用常规更改' : '更改需确认'}
              </span>
            </div>
            {state.busy || state.applying ? (
              <button type="button" className="stop-button" onClick={() => session.stop()}>
                Stop
              </button>
            ) : (
              <button
                className="send-button"
                type="button"
                aria-label="Send message"
                disabled={
                  props.connectionAvailable === false ||
                  !instruction.trim() ||
                  state.applying ||
                  uploadPending ||
                  Boolean(state.proposal) ||
                  projectPhase !== 'idle'
                }
                onClick={send}
              >
                <IconEnter size={22} />
              </button>
            )}
          </div>
        </div>
        {uploadError && (
          <p className="composer-upload-error error-text" role="alert">
            {uploadError}
          </p>
        )}
      </section>
    </main>
  )
}

export function LegacyAgentWorkspace(props: {
  session: OfficeAgentSession
  ui: OfficeWorkspaceUi
  disconnect: () => void
  host: OfficeHost
  connectionNotice?: string
}) {
  return <AgentWorkspace {...props} legacy />
}

export function workspaceComponentForMode(mode: 'workspace' | 'legacy') {
  return mode === 'legacy' ? LegacyAgentWorkspace : AgentWorkspace
}

type ConnectionBridge = PcBridgeSession | OfficeRelaySession

export function ConfiguredApp(
  props: {
    documentClient?: OfficeDocumentClient
    connectionBridge?: ConnectionBridge
    workspaceFactory?: (dependencies: {
      host: Exclude<OfficeHost, 'unknown'>
      document: OfficeDocumentClient
      bridge: ConnectionBridge
    }) => { runtime: OfficeHostRuntime; session: OfficeAgentSession; ui: OfficeWorkspaceUi }
  } = {},
) {
  const workspaceFactory = props.workspaceFactory
  const document = useMemo(
    () => props.documentClient ?? createOfficeDocumentClient(createBrowserOfficeRuntime()),
    [props.documentClient],
  )
  const transportMode = useMemo(() => officeTransportMode(import.meta.env), [])
  const teamRuntime = useRef<OfficeHostRuntime | undefined>(undefined)
  const teamConnection = useMemo(() => {
    const config = officeTeamAuthConfig(import.meta.env, window.location.origin)
    if (!config || transportMode !== 'relay') return undefined
    return createOfficeTeamConnection({
      auth: createBrowserAuthClient({ config, store: createMemorySessionStore() }),
      loginDialog: createOfficeTeamLoginDialog({ ...config, addinOrigin: window.location.origin }),
      onUnavailable: () => teamRuntime.current?.team?.clear(),
    })
  }, [transportMode])
  useEffect(() => () => teamConnection?.dispose(), [teamConnection])
  const remoteDiagnosticsEnabled = useMemo(
    () => transportMode === 'relay' && officeRemoteDiagnosticsEnabled(import.meta.env),
    [transportMode],
  )
  const diagnosticSamplePercent = useMemo(() => officeDiagnosticSamplePercent(import.meta.env), [])
  const bridge = useMemo(
    () =>
      props.connectionBridge ??
      (transportMode === 'loopback'
        ? createPcBridgeSession()
        : createOfficeRelaySession({
            capabilities: [
              'agent.v1',
              'presentation.v1',
              'presentation-attachments.v1',
              'presentation-assets.v1',
              'presentation-remote-images.v1',
              'presentation-webpages.v1',
              'presentation-asset-rights.v1',
              'presentation-animation-frame.v1',
              'presentation-pdf.v1',
              'presentation-production-pdf.v1',
              'presentation-master-backups.v1',
              'presentation-package-backups.v1',
              ...(import.meta.env.VITE_WISWORK_PPT_PROJECT_GOVERNANCE_ENABLED === '1'
                ? ['presentation-governance.v1' as const]
                : []),
              'web-search.v1',
              'image-search.v1',
              'image-fetch.v1',
              'design-document.v1',
              'enhanced-lease.v1',
            ],
            persistentPairing: __WISWORK_OFFICE_PAIRING_RESUME__,
          })),
    [props.connectionBridge, transportMode],
  )
  const bridgeState = useSyncExternalStore(
    (listener) => bridge.subscribe(listener),
    () => bridge.snapshot(),
    () => bridge.snapshot(),
  )
  const designRequest = useMemo(
    () => ('capabilityFetch' in bridge ? createOfficeDesignRequest(bridge) : undefined),
    [bridge],
  )
  const [workspace, setWorkspace] = useState<
    { runtime: OfficeHostRuntime; session: OfficeAgentSession; ui: OfficeWorkspaceUi } | undefined
  >()
  const workspaceMode = useMemo(() => officeWorkspaceMode(import.meta.env), [])
  const capabilityFlags = useMemo(() => officeCapabilityFlags(import.meta.env), [])
  const presentationRolloutPercent = useMemo(
    () => officePresentationRolloutPercent(import.meta.env),
    [],
  )
  const presentationFlags = useMemo(() => officePresentationVerificationFlags(import.meta.env), [])
  const [host, setHost] = useState<OfficeHost>('unknown')
  const [hostSupported, setHostSupported] = useState(false)
  const [presentationRolloutExcluded, setPresentationRolloutExcluded] = useState(false)
  const [status, setStatus] = useState('Connecting to Office…')
  const [busy, setBusy] = useState(true)
  const [pairingForgetError, setPairingForgetError] = useState(false)
  const [pairingForgetBusy, setPairingForgetBusy] = useState(false)
  const rawDocumentId = useRef(`document_${crypto.randomUUID().replaceAll('-', '')}`)

  const forgetPairing = async () => {
    if (!('forget' in bridge)) {
      bridge.disconnect()
      setPairingForgetError(false)
      return
    }
    setPairingForgetBusy(true)
    try {
      await bridge.forget()
      setPairingForgetError(false)
    } catch {
      setPairingForgetError(true)
    } finally {
      setPairingForgetBusy(false)
    }
  }

  useEffect(() => {
    let active = true
    let created:
      { runtime: OfficeHostRuntime; session: OfficeAgentSession; ui: OfficeWorkspaceUi } | undefined
    void (async () => {
      try {
        const activeHost = await document.initialize()
        if (active) {
          setHost(activeHost)
          setHostSupported(activeHost !== 'unknown')
          if (activeHost !== 'unknown') {
            void bridge.connect(activeHost)
            if (workspaceFactory) created = workspaceFactory({ host: activeHost, document, bridge })
            else {
              if (
                activeHost === 'powerpoint' &&
                !presentationRolloutEnabled(
                  Office.context.document.url || undefined,
                  presentationRolloutPercent,
                )
              ) {
                setPresentationRolloutExcluded(true)
                setStatus('PPT Agent is not enabled for this presentation yet.')
                return
              }
              const presentationBinding =
                activeHost === 'powerpoint' ? createBrowserPresentationDocumentBinding() : undefined
              const boundPresentationDocumentId = presentationBinding
                ? await presentationBinding.documentId()
                : undefined
              const runRecovery =
                presentationBinding && boundPresentationDocumentId
                  ? await preparePresentationAgentRunRecovery(
                      presentationBinding,
                      boundPresentationDocumentId,
                    )
                  : undefined
              const runCheckpoint =
                presentationBinding && boundPresentationDocumentId
                  ? createPresentationAgentRunCheckpoint(
                      presentationBinding,
                      boundPresentationDocumentId,
                      (() => {
                        try {
                          return window.localStorage
                        } catch {
                          return undefined
                        }
                      })(),
                    )
                  : undefined
              const governancePersistence =
                import.meta.env.VITE_WISWORK_PPT_PROJECT_GOVERNANCE_ENABLED === '1'
                  ? (() => {
                      try {
                        return createPresentationGovernanceStorage(window.localStorage)
                      } catch {
                        return undefined
                      }
                    })()
                  : undefined
              const researchDeletePersistence = boundPresentationDocumentId
                ? createPresentationResearchDeletePersistence(
                    boundPresentationDocumentId,
                    (() => {
                      try {
                        return window.localStorage
                      } catch {
                        return undefined
                      }
                    })(),
                  )
                : undefined
              const researchAbandonPersistence = boundPresentationDocumentId
                ? createPresentationResearchAbandonPersistence(
                    boundPresentationDocumentId,
                    (() => {
                      try {
                        return window.localStorage
                      } catch {
                        return undefined
                      }
                    })(),
                  )
                : undefined
              const interruptedRun = runRecovery?.scrubFailed
                ? undefined
                : runCheckpoint?.recovery()
              const environment = officeDiagnosticEnvironment(activeHost)
              const diagnostics = createOfficeDiagnostics({
                host: activeHost,
                platform: environment.platform,
                build: __WISWORK_OFFICE_BUILD_ID__,
                localDocumentId: boundPresentationDocumentId,
                localSessionId: () =>
                  'diagnosticSessionId' in bridge ? bridge.diagnosticSessionId() : undefined,
                requirementSets: environment.requirementSets,
                remoteEnabled: remoteDiagnosticsEnabled,
                remoteSamplePercent: diagnosticSamplePercent,
                send: (event) => {
                  if (!('sendDiagnostic' in bridge)) throw new Error('diagnostic_upload_failed')
                  return bridge.sendDiagnostic(event)
                },
              })
              const runtime = createOfficeHostRuntime(activeHost, {
                enableHostSkills: import.meta.env.VITE_WISWORK_OFFICE_HOST_SKILLS !== '0',
                presentationVerification: presentationFlags,
                presentationTelemetry: (event) =>
                  window.dispatchEvent(
                    new CustomEvent('wiswork:presentation-telemetry', { detail: event }),
                  ),
                enableConversions: capabilityFlags.conversions,
                enableSkillPackages: capabilityFlags.skillPackages,
                enableImportMedia: capabilityFlags.importMedia,
                ...('capabilityFetch' in bridge && activeHost === 'powerpoint'
                  ? {
                      fetchPowerPointImage: async (url: string, signal?: AbortSignal) => {
                        const response = await bridge.capabilityFetch(
                          'image-fetch.v1',
                          { url },
                          signal,
                        )
                        if (!response.ok) throw new Error('image_fetch_unavailable')
                        const payload = (await response.json()) as { data_base64?: unknown }
                        if (typeof payload.data_base64 !== 'string')
                          throw new Error('image_fetch_unavailable')
                        const binary = atob(payload.data_base64)
                        return Uint8Array.from(binary, (character) => character.charCodeAt(0))
                      },
                      powerPointImageFetchAvailable: () =>
                        bridge.snapshot().capabilities?.includes('image-fetch.v1') === true,
                    }
                  : {}),
                document,
                diagnostics,
                ...('capabilityFetch' in bridge && activeHost === 'powerpoint'
                  ? {
                      additionalSkills: [
                        createOfficeWebSkill(bridge, {
                          advertisedCapabilities: ['web-search.v1', 'image-search.v1'],
                        }),
                      ],
                    }
                  : {}),

                ...(activeHost === 'powerpoint' && 'capabilityFetch' in bridge
                  ? {
                      presentation: {
                        ...presentationBinding!,
                        projectGovernanceEnabled:
                          import.meta.env.VITE_WISWORK_PPT_PROJECT_GOVERNANCE_ENABLED === '1',
                        governanceAvailable: () => {
                          const current = bridge.snapshot()
                          return (
                            current.status === 'connected' &&
                            current.capabilities?.includes('presentation-governance.v1') === true
                          )
                        },
                        governanceRequest: (body: unknown, signal?: AbortSignal) =>
                          bridge.capabilityFetch('presentation-governance.v1', body, signal),
                        governanceSessionId: () =>
                          'diagnosticSessionId' in bridge
                            ? bridge.diagnosticSessionId()
                            : undefined,
                        readGovernanceAttempt: governancePersistence?.read,
                        writeGovernanceAttempt: governancePersistence?.write,
                        teamAvailable: () => teamConnection?.available() === true,
                        teamRequest: teamConnection
                          ? (body: unknown, signal?: AbortSignal) =>
                              teamConnection.request(body, signal)
                          : undefined,
                        readResearchAbandonAttempt: researchAbandonPersistence?.read,
                        writeResearchAbandonAttempt: researchAbandonPersistence?.write,
                        readResearchDeleteAttempt: researchDeletePersistence?.read,
                        writeResearchDeleteAttempt: researchDeletePersistence?.write,
                        available: () => {
                          const snapshot = bridge.snapshot()
                          return (
                            snapshot.status === 'connected' &&
                            snapshot.capabilities?.includes('presentation.v1') === true
                          )
                        },
                        request: (body: unknown, signal?: AbortSignal) =>
                          bridge.capabilityFetch('presentation.v1', body, signal),
                        packageBackupAvailable: () => {
                          const snapshot = bridge.snapshot()
                          return (
                            snapshot.status === 'connected' &&
                            snapshot.capabilities?.includes('presentation-package-backups.v1') ===
                              true
                          )
                        },
                        packageBackupRequest: (body: unknown, signal?: AbortSignal) =>
                          bridge.capabilityFetch('presentation-package-backups.v1', body, signal),
                        masterBackupAvailable: () => {
                          const snapshot = bridge.snapshot()
                          return (
                            snapshot.status === 'connected' &&
                            snapshot.capabilities?.includes('presentation-master-backups.v1') ===
                              true
                          )
                        },
                        masterBackupRequest: (body: unknown, signal?: AbortSignal) =>
                          bridge.capabilityFetch('presentation-master-backups.v1', body, signal),
                        pdfAvailable: () => {
                          const snapshot = bridge.snapshot()
                          return (
                            snapshot.status === 'connected' &&
                            snapshot.capabilities?.includes('presentation-pdf.v1') === true
                          )
                        },
                        productionPdfAvailable: () => {
                          const snapshot = bridge.snapshot()
                          return (
                            snapshot.status === 'connected' &&
                            snapshot.capabilities?.includes('presentation-production-pdf.v1') ===
                              true
                          )
                        },
                        pdfRequest: (body: unknown, signal?: AbortSignal) =>
                          bridge.capabilityFetch(
                            (body as { source?: string })?.source === 'production'
                              ? 'presentation-production-pdf.v1'
                              : 'presentation-pdf.v1',
                            body,
                            signal,
                          ),
                        assetsAvailable: () => {
                          const snapshot = bridge.snapshot()
                          return (
                            snapshot.status === 'connected' &&
                            snapshot.capabilities?.includes('presentation-assets.v1') === true
                          )
                        },
                        remoteImagesAvailable: () => {
                          const snapshot = bridge.snapshot()
                          return (
                            snapshot.status === 'connected' &&
                            snapshot.capabilities?.includes('presentation-remote-images.v1') ===
                              true
                          )
                        },
                        webpagesAvailable: () => {
                          const snapshot = bridge.snapshot()
                          return (
                            snapshot.status === 'connected' &&
                            snapshot.capabilities?.includes('presentation-webpages.v1') === true
                          )
                        },
                        rightsAvailable: () => {
                          const snapshot = bridge.snapshot()
                          return (
                            snapshot.status === 'connected' &&
                            snapshot.capabilities?.includes('presentation-asset-rights.v1') === true
                          )
                        },
                        animationFrameAvailable: () => {
                          const snapshot = bridge.snapshot()
                          return (
                            snapshot.status === 'connected' &&
                            snapshot.capabilities?.includes('presentation-animation-frame.v1') ===
                              true
                          )
                        },
                        attachmentsAvailable: () => {
                          const snapshot = bridge.snapshot()
                          return (
                            snapshot.status === 'connected' &&
                            snapshot.capabilities?.includes('presentation-attachments.v1') === true
                          )
                        },
                        attachmentsRequest: (body: unknown, signal?: AbortSignal) =>
                          bridge.capabilityFetch(
                            body &&
                              typeof body === 'object' &&
                              'operation' in body &&
                              body.operation === 'attachment_import_url'
                              ? 'presentation-remote-images.v1'
                              : body &&
                                  typeof body === 'object' &&
                                  'operation' in body &&
                                  body.operation === 'attachment_import_webpage'
                                ? 'presentation-webpages.v1'
                                : body &&
                                    typeof body === 'object' &&
                                    'operation' in body &&
                                    [
                                      'attachment_attest_license',
                                      'attachment_revoke_license',
                                    ].includes(body.operation as string)
                                  ? 'presentation-asset-rights.v1'
                                  : body &&
                                      typeof body === 'object' &&
                                      'operation' in body &&
                                      body.operation === 'attachment_extract_first_frame'
                                    ? 'presentation-animation-frame.v1'
                                    : 'presentation-attachments.v1',
                            body,
                            signal,
                          ),
                      },
                    }
                  : {}),
              })
              const session = createOfficeAgentSession({
                host: activeHost,
                transport: createPcBridgeAgentTransport(bridge),
                skill: runtime.skill,
                proposals: runtime.proposals,
                automaticPowerPointMutations: activeHost === 'powerpoint',
                ...('setToolHandler' in bridge ? { remoteTools: bridge } : {}),
                presentationText: (key) =>
                  officePresentationText(globalThis.Office?.context?.displayLanguage, key),
                diagnostics,
                ...(presentationBinding && boundPresentationDocumentId
                  ? {
                      runCheckpoint: {
                        interrupted: runRecovery!.interrupted,
                        scrubFailed: runRecovery!.scrubFailed,
                        recovery: interruptedRun,
                        readRecovery: () =>
                          runRecovery!.scrubFailed ? undefined : runCheckpoint!.recovery(),
                        validateDocument: async () =>
                          (await presentationBinding.documentId()) === boundPresentationDocumentId,
                        begin: runCheckpoint!.begin,
                        tool: runCheckpoint!.tool,
                        conversation: runCheckpoint!.conversation,
                        adopt: runCheckpoint!.adopt,
                        finish: runCheckpoint!.finish,
                      },
                    }
                  : {}),
              })
              teamRuntime.current = runtime
              created = {
                runtime,
                session,
                ui: createOfficeWorkspaceUi(
                  runtime,
                  diagnostics,
                  undefined,
                  interruptedRun?.changeReceipt && interruptedRun.toolCallId
                    ? { agentRunId: interruptedRun.runId, toolCallId: interruptedRun.toolCallId }
                    : undefined,
                  teamConnection,
                ),
              }
            }
            setWorkspace(created)
          }
          setStatus(
            activeHost === 'unknown'
              ? 'office_host_unsupported'
              : `${hostLabels[activeHost]} is ready`,
          )
        }
      } catch {
        if (active) setStatus('office_unavailable')
      } finally {
        if (active) setBusy(false)
      }
    })()
    return () => {
      active = false
      if (teamRuntime.current === created?.runtime) teamRuntime.current = undefined
      created?.session.dispose()
      created?.runtime.dispose()
      bridge.disconnect()
    }
  }, [
    bridge,
    capabilityFlags,
    diagnosticSamplePercent,
    document,
    presentationRolloutPercent,
    remoteDiagnosticsEnabled,
    teamConnection,
    presentationFlags,
    workspaceFactory,
  ])

  useEffect(() => {
    if (shouldResetOfficeSession(bridgeState.status) && workspace) {
      workspace.session.authenticationLost()
      workspace.runtime.clearSession()
      workspace.runtime.presentation?.prepareReconnect()
    }
  }, [bridgeState.status, workspace])

  const governanceSessionId =
    'diagnosticSessionId' in bridge ? bridge.diagnosticSessionId() : undefined
  useEffect(() => {
    workspace?.runtime.governance?.clear()
  }, [workspace, governanceSessionId])
  useEffect(() => {
    workspace?.runtime.setPowerPointImageFetchAvailable?.(
      bridgeState.status === 'connected' &&
        'capabilities' in bridgeState &&
        bridgeState.capabilities?.includes('image-fetch.v1') === true,
    )
  }, [bridgeState, workspace])

  useEffect(() => {
    const enhanced = bridgeState.enhanced
    if (!workspace || host === 'unknown') return
    if (bridgeState.status !== 'connected' || !enhanced?.raw_office) {
      workspace.runtime.disableElevatedOffice?.()
      return
    }
    workspace.runtime.enableElevatedOffice?.(
      () => {
        const snapshot = bridge.snapshot()
        const current = snapshot.enhanced
        const raw = current ? rawOfficeCapabilities(current) : { rawJs: false, rawOoxml: false }
        const valid =
          snapshot.status === 'connected' &&
          current?.raw_office === true &&
          current.host === `office-${host}` &&
          current.expires_at > Date.now()
        return {
          activeMode: valid ? ('enhanced' as const) : ('standard' as const),
          signedIn: valid,
          paired: valid,
          hostEnabled: valid,
          rawOfficeEnabled: valid,
          rawOfficeJsEnabled: valid && raw.rawJs,
          rawOfficeOoxmlEnabled: valid && raw.rawOoxml,
          documentId: rawDocumentId.current,
          sessionId: current?.runtime_instance ?? 'revoked_session_0000',
          generation: current?.session_generation ?? -1,
          revision: `revision_${String(current?.session_generation ?? 0).padStart(8, '0')}`,
        }
      },
      translateRawOfficeConfirmation(normalizeLang(globalThis.Office?.context?.displayLanguage)),
    )
  }, [bridge, bridgeState.enhanced, bridgeState.status, host, workspace])

  if (busy) return <StatusScreen title="Starting WisWork Agent" detail={status} busy />
  if (presentationRolloutExcluded)
    return <StatusScreen title="PPT Agent unavailable" detail={status} />
  if (!hostSupported) {
    return (
      <StatusScreen title="Unsupported Office host" detail="This host cannot use document tools." />
    )
  }
  if (pairingForgetError || shouldShowRelayStatusScreen(bridgeState.status, Boolean(workspace))) {
    if (pairingForgetError) {
      return (
        <StatusScreen
          title="Couldn’t forget this Office pairing"
          detail="This taskpane is disconnected, but its saved pairing could not be removed. Try again before reconnecting."
          busy={pairingForgetBusy}
        >
          <button type="button" disabled={pairingForgetBusy} onClick={() => void forgetPairing()}>
            {pairingForgetBusy ? 'Forgetting pairing…' : 'Try forgetting again'}
          </button>
        </StatusScreen>
      )
    }
    const presentation = relayConnectionPresentation(
      bridgeState.status,
      bridgeState.verificationCode,
    )
    return (
      <StatusScreen
        title={presentation.title}
        detail={presentation.detail}
        busy={presentation.busy}
      >
        <button
          type="button"
          disabled={presentation.actionDisabled}
          onClick={() => {
            void bridge.connect(host)
          }}
        >
          {bridgeState.status === 'offline'
            ? 'Connect to WisWork PC'
            : bridgeState.status === 'connecting'
              ? 'Looking for WisWork PC…'
              : 'Try again'}
        </button>
        {workspace?.ui.copyDiagnostics && (
          <DiagnosticCopyButton
            copyDiagnostics={workspace.ui.copyDiagnostics}
            copyDiagnosticsWithContext={workspace.ui.copyDiagnosticsWithContext}
          />
        )}
      </StatusScreen>
    )
  }
  if (!workspace)
    return <StatusScreen title="Starting WisWork Agent" detail="Loading tools…" busy />
  const disconnect = () => {
    void teamConnection?.signOut()
    workspace.session.logout()
    workspace.runtime.disableElevatedOffice()
    workspace.runtime.clearSession()
    void forgetPairing()
  }
  const WorkspaceComponent = workspaceComponentForMode(workspaceMode)
  const connectionNotice =
    bridgeState.status !== 'connected'
      ? relayConnectionPresentation(bridgeState.status, bridgeState.verificationCode).detail
      : 'remembered' in bridgeState
        ? relayPersistenceNotice(bridgeState as OfficeRelaySnapshot)
        : undefined
  return (
    <WorkspaceComponent
      session={workspace.session}
      ui={workspace.ui}
      disconnect={disconnect}
      host={host}
      connectionNotice={connectionNotice}
      runtimeMode={officeRuntimeModeForTaskpane(host, bridgeState)}
      connectionAvailable={bridgeState.status === 'connected'}
      designRequest={
        'capabilities' in bridgeState && bridgeState.capabilities?.includes('design-document.v1')
          ? designRequest
          : undefined
      }
      repairDesignConnection={() => forgetPairing()}
    />
  )
}

function StatusScreen(props: {
  title: string
  detail: string
  busy?: boolean
  children?: React.ReactNode
}) {
  return (
    <main className="taskpane centered">
      <section className="welcome-card">
        <span className="eyebrow">WisWork Office</span>
        <h1>{props.title}</h1>
        <p aria-live="polite">{props.detail}</p>
        {props.busy && <span className="loading-line" aria-hidden="true" />}
        {props.children}
      </section>
    </main>
  )
}

export function App() {
  const [versionState, setVersionState] = useState<BuildVersionState>({ status: 'checking' })
  const [versionAttempt, setVersionAttempt] = useState(0)
  useEffect(() => {
    let active = true
    void (async () => {
      const deployed = await deployedBuildId()
      if (active)
        setVersionState(
          resolveBuildVersion(deployed, __WISWORK_OFFICE_BUILD_ID__, import.meta.env.PROD),
        )
    })()
    return () => {
      active = false
    }
  }, [versionAttempt])
  if (versionState.status === 'checking')
    return <StatusScreen title="Checking WisWork version" detail="Checking for updates…" busy />
  if (versionState.status === 'unavailable')
    return (
      <StatusScreen
        title="Cannot verify WisWork version"
        detail="Check the connection and retry before using document tools."
      >
        <button
          type="button"
          onClick={() => {
            setVersionState({ status: 'checking' })
            setVersionAttempt((attempt) => attempt + 1)
          }}
        >
          Retry version check
        </button>
      </StatusScreen>
    )
  if (versionState.status === 'stale')
    return (
      <StatusScreen
        title="WisWork update available"
        detail="Reload this pane to use the latest version."
      >
        <button
          type="button"
          onClick={() => {
            const url = new URL(window.location.href)
            url.searchParams.set('v', versionState.buildId)
            window.location.replace(url.href)
          }}
        >
          Reload WisWork
        </button>
      </StatusScreen>
    )
  return <ConfiguredApp />
}
