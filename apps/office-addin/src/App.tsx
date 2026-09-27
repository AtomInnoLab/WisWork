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
import { PresentationWorkflowCard } from './agent/presentation-workflow-card.js'
import type { PresentationProjectController } from './skills/powerpoint/presentation-project.js'
import {
  createBrowserPresentationDocumentBinding,
  createPresentationAgentRunCheckpoint,
} from './skills/powerpoint/presentation-document.js'
import { downloadSessionFile } from './agent/session-download.js'
import { useEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react'
import { deployedBuildId, resolveBuildVersion, type BuildVersionState } from './build-version.js'
import { Markdown } from '@wiswork/ui'
import { createOfficeHostRuntime, type OfficeHostRuntime } from './agent/host-runtime.js'
import type { PresentationAttachmentMetadata } from './skills/powerpoint/presentation-attachments.js'
import {
  officeCapabilityFlags,
  officeRemoteDiagnosticsEnabled,
  officeWorkspaceMode,
} from '../build-config.js'
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
import type {
  OfficePresentationEvent,
  ProposalPresentationEvent,
} from './agent/presentation-state.js'
import { createPcBridgeSession } from './pc-bridge/session.js'
import { createOfficeRelaySession, officeTransportMode } from './relay/session.js'
import {
  createBrowserOfficeRuntime,
  createOfficeDocumentClient,
  type OfficeHost,
} from './office-document.js'

const hostLabels: Record<OfficeHost, string> = {
  word: 'Microsoft Word',
  excel: 'Microsoft Excel',
  powerpoint: 'Microsoft PowerPoint',
  unknown: 'Office',
}

const agentProductLabels: Record<OfficeHost, string> = {
  word: 'AI Word',
  excel: 'AI Sheets',
  powerpoint: 'AI Slides',
  unknown: 'WisWork AI',
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
      : proposal.impact.targets.map((target) => proposalTarget(target, proposal.impact.host)),
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
  }
}

const MEBIBYTE = 1024 * 1024

function displayMegabytes(bytes: number): string {
  const value = bytes / MEBIBYTE
  return Number.isInteger(value) ? String(value) : value.toFixed(1)
}

export function safeUploadError(error: unknown, file?: Pick<SessionFile, 'size'>): string {
  const code = error instanceof Error ? error.message : ''
  const attachmentErrors: Record<string, string> = {
    presentation_attachment_too_large: '制作资料每个文件最多 50 MB。',
    presentation_image_too_large: '图片每个文件最多 10 MB。',
    presentation_remote_image_unavailable:
      '图片网址无法安全下载或图片格式不受支持，请检查网址后重试。',
    presentation_remote_image_source_conflict:
      '相同图片内容已从另一来源加入当前文档。请使用已有素材，或手动上传本地文件。',
    presentation_aborted: '图片下载超时或已取消，请重试。',
    presentation_parse_failed: '图片无法解码为受支持的 PNG 或 JPEG。',
    presentation_assets_unavailable: '请更新并连接支持图片素材的 PC 端后重试。',
    presentation_attachment_failed: '资料解析未完成，请检查文件或重新上传。',
    presentation_not_found: '这份 PC 资料已不存在，请刷新附件列表。',
    presentation_attachment_in_use: '这份资料正被图片使用权声明引用，请先撤回声明再删除。',
    presentation_invalid_state: 'PC 资料状态异常，请重连后重试。',
    presentation_quota_exceeded:
      '当前文档在 PC 的资料容量已满（资料最多 32 个，附件预留容量总计 100 MB）。请删除不再需要的资料或图片后重试。',
    presentation_document_changed: '文档已改变，本次上传已停止。请在目标文档重新上传。',
    presentation_service_unavailable: 'PC 连接不可用，请重连后重新选择同一文件续传。',
    presentation_unavailable: 'PC 连接不可用，请重连后重新选择同一文件续传。',
    presentation_response_invalid: '上传响应无效，请重连后重新选择同一文件续传。',
  }
  if (attachmentErrors[code]) return attachmentErrors[code]
  if (code === 'vfs_limit') {
    if (file && file.size > MAX_VFS_FILE_BYTES) {
      return `File is ${displayMegabytes(file.size)} MB. Attachments must be ${displayMegabytes(MAX_VFS_FILE_BYTES)} MB or smaller.`
    }
    return `Attachment limit reached. Files are limited to ${displayMegabytes(MAX_VFS_FILE_BYTES)} MB each and ${displayMegabytes(MAX_VFS_TOTAL_BYTES)} MB per session.`
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
  readonly project?: PresentationProjectController
  readonly importProgress?: PresentationImportProgressController
  readonly qa?: PresentationQaController
  readonly changes?: PresentationChangesController
  readonly durableAttachmentsAvailable?: () => boolean
  readonly durableImagesAvailable?: () => boolean
  readonly remoteImagesAvailable?: () => boolean
  readonly rightsAvailable?: () => boolean
  readonly listDurableAttachments?: () => Promise<PresentationAttachmentMetadata[]>
  readonly deleteDurableAttachment?: (attachmentId: string) => Promise<void>
  readonly importPresentationImageUrl?: (url: string) => Promise<void>
  readonly attestPresentationImageLicense?: (
    imageId: string,
    license: 'owned' | 'licensed' | 'public_domain',
    evidenceId: string,
  ) => Promise<void>
  readonly revokePresentationImageLicense?: (imageId: string) => Promise<void>
  readonly attachments: () => readonly string[]
  readonly downloadFile?: (path: string) => void
  readonly skills: () => readonly string[]
  readonly skillPackagesEnabled: boolean
  readonly upload: (file: SessionFile) => Promise<void>
  readonly copyDiagnostics?: () => Promise<void>
  readonly uninstallSkill?: (name: string) => void
  readonly clear: () => void
}

export function DiagnosticCopyButton(props: {
  copyDiagnostics: () => Promise<void>
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
            .then(() => setStatus('诊断信息已复制'))
            .catch(() => setStatus('复制诊断信息失败'))
        }}
      >
        复制诊断信息
      </button>
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
): OfficeWorkspaceUi {
  return Object.freeze({
    project: runtime.presentation,
    importProgress: runtime.importProgress,
    qa: runtime.qa,
    changes: runtime.changes,
    durableAttachmentsAvailable: runtime.durableAttachmentsAvailable,
    durableImagesAvailable: runtime.durableImagesAvailable,
    remoteImagesAvailable: runtime.remoteImagesAvailable,
    rightsAvailable: runtime.rightsAvailable,
    listDurableAttachments: runtime.listDurableAttachments,
    deleteDurableAttachment: runtime.deleteDurableAttachment,
    importPresentationImageUrl: runtime.importPresentationImageUrl,
    attestPresentationImageLicense: runtime.attestPresentationImageLicense,
    revokePresentationImageLicense: runtime.revokePresentationImageLicense,
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
              : event.state === 'applied'
                ? 'Change applied'
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
      <details className="proposal-preview" open>
        <summary>Review exact impact</summary>
        {hasComparison && (
          <div className="proposal-diff">
            <div className="preview-block">
              <strong>Before</strong>
              <p className="proposal-copy">{presentation.before || '(empty document)'}</p>
            </div>
            <div className="preview-block after">
              <strong>After</strong>
              <p className="proposal-copy">{presentation.after || '(empty document)'}</p>
            </div>
          </div>
        )}
        {presentation.preview && <p className="proposal-copy">{presentation.preview}</p>}
      </details>
      {event.error && <p className="error-text">{event.error}</p>}
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
            disabled={props.applying}
            onClick={() => props.confirm(event.proposal.id)}
          >
            {props.applying ? 'Applying…' : 'Confirm change'}
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

export function AgentWorkspace(props: {
  session: OfficeAgentSession
  ui: OfficeWorkspaceUi
  disconnect: () => void
  host: OfficeHost
  initialPanel?: WorkspacePanelName
  legacy?: boolean
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
  const [durableFiles, setDurableFiles] = useState<PresentationAttachmentMetadata[]>([])
  const [skills, setSkills] = useState<readonly string[]>(ui.skills())
  const [uploadError, setUploadError] = useState('')
  const [uploadPending, setUploadPending] = useState(false)
  const [uploadStatus, setUploadStatus] = useState('')
  const [imageUrl, setImageUrl] = useState('')
  const uploadEpoch = useRef(0)
  const [diagnosticStatus, setDiagnosticStatus] = useState('')
  const [panel, setPanel] = useState<WorkspacePanelName | undefined>(props.initialPanel)
  const mounted = useRef(true)
  const panelHeading = useRef<HTMLHeadingElement>(null)
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

  useEffect(() => ui.project?.subscribe(() => setFiles(ui.attachments())), [ui])

  function send() {
    if (
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

  const proposal = state.proposal
  const hasTimeline = state.timeline.length > 0
  const showConversationChrome =
    hasTimeline || state.busy || state.applying || Boolean(state.error) || Boolean(proposal)
  const showStatus =
    state.busy || state.applying || Boolean(state.activity) || state.status === 'cancelled'

  return (
    <main
      className={`agent-workspace ${props.legacy ? 'legacy-workspace ' : ''}${panel ? 'has-management ' : ''}${showConversationChrome ? 'has-conversation' : 'is-empty'}`}
      aria-busy={state.busy || state.applying}
    >
      {showConversationChrome && (
        <header className="app-header">
          <div className="editor-identity">
            <span className="connection-dot" aria-hidden="true" />
            <h1>{agentProductLabels[host]}</h1>
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
                      .then(() => mounted.current && setDiagnosticStatus('诊断信息已复制'))
                      .catch(() => mounted.current && setDiagnosticStatus('复制诊断信息失败'))
                    event.currentTarget.closest('details')?.removeAttribute('open')
                  }}
                >
                  复制诊断信息
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

      <section
        ref={timeline}
        className="agent-timeline"
        aria-label="Agent conversation"
        aria-live="polite"
        onScroll={(event) => {
          followLatest.current = isTimelineNearBottom(event.currentTarget)
        }}
      >
        {!hasTimeline && (
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
        {state.timeline.map((event) => (
          <TimelineEvent
            key={event.id}
            event={event}
            activeProposalId={proposal?.id}
            busy={state.busy}
            applying={state.applying}
            confirm={(id) => void session.confirm(id)}
            reject={() => session.reject()}
          />
        ))}
        {state.recoveryAvailable && (
          <button
            type="button"
            className="secondary"
            disabled={state.busy || state.applying || projectPhase !== 'idle'}
            onClick={() => void session.resumeInterrupted?.()}
          >
            继续上次请求
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
                    })
                    .finally(() => {
                      if (current()) setUploadPending(false)
                    })
                }}
              />
              <p>
                {ui.durableAttachmentsAvailable?.()
                  ? 'PDF、Word（DOCX）、TXT、MD、CSV、JSON 资料每个最多 50 MB，保存于 PC 并绑定当前文档；退出登录不会删除。重连后可让 Agent 列出和读取，重新选择同一文件可续传。'
                  : `Files are limited to ${displayMegabytes(MAX_VFS_FILE_BYTES)} MB each and ${displayMegabytes(MAX_VFS_TOTAL_BYTES)} MB per session, then cleared on logout.`}
              </p>
              {ui.durableImagesAvailable?.() && (
                <>
                  <p>
                    PNG、JPEG 图片每个最多 10 MB，上传后在 PC
                    校验并缓存；可直接用于制作，无需将图片编码发给 Agent。
                  </p>
                  {ui.remoteImagesAvailable?.() && (
                    <form
                      onSubmit={(event) => {
                        event.preventDefault()
                        if (uploadPending || state.busy || !imageUrl.trim()) return
                        setUploadPending(true)
                        setUploadError('')
                        setUploadStatus('正在由 PC 下载并校验图片…')
                        void ui
                          .importPresentationImageUrl?.(imageUrl.trim())
                          .then(async () => {
                            if (!mounted.current) return
                            setDurableFiles((await ui.listDurableAttachments?.()) ?? [])
                            setUploadStatus(
                              '图片已保存到当前文档的 PC 素材缓存；许可状态仍需核验。',
                            )
                            setImageUrl('')
                          })
                          .catch((error: unknown) => {
                            if (mounted.current) {
                              setUploadStatus('')
                              setUploadError(safeUploadError(error))
                            }
                          })
                          .finally(() => {
                            if (mounted.current) setUploadPending(false)
                          })
                      }}
                    >
                      <label htmlFor="presentation-image-url">图片网址</label>
                      <input
                        id="presentation-image-url"
                        type="url"
                        value={imageUrl}
                        onChange={(event) => setImageUrl(event.currentTarget.value)}
                        placeholder="https://example.com/image.png"
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
                  下方仅显示本次会话可下载的副本；超过 20 MB 或会话容量的资料仍可由 Agent 在 PC
                  读取。其他文件及技能仅保留在会话中。
                </p>
              )}
              {uploadStatus && <p role="status">{uploadStatus}</p>}
              {uploadError && (
                <p className="error-text" role="alert">
                  {uploadError}
                </p>
              )}
              {ui.durableAttachmentsAvailable?.() && durableFiles.length > 0 && (
                <ul>
                  {durableFiles.map((file) => (
                    <li key={file.attachmentId}>
                      {file.name} · {file.status}
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
                        {file.split('/').at(-1)} · 下载
                      </button>
                    ) : (
                      file.split('/').at(-1)
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
            disabled={uploadPending || state.busy || state.applying || Boolean(state.proposal)}
          />
        )}
        {ui.importProgress && <PresentationImportProgressCard controller={ui.importProgress} />}
        {ui.changes && (
          <PresentationChangesCard
            controller={ui.changes}
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
            (file) => file.startsWith('/home/user/generated/') && file.endsWith('.pptx'),
          ) && (
            <section aria-label="生成的演示文稿" className="presentation-downloads">
              {files
                .filter(
                  (file) =>
                    file.startsWith('/home/user/generated/') &&
                    (file.endsWith('.pptx') || file.endsWith('.report.json')),
                )
                .slice(-4)
                .map((file) => (
                  <button type="button" key={file} onClick={() => ui.downloadFile?.(file)}>
                    {file.endsWith('.pptx') ? '下载 PPTX' : '下载验收报告'} ·{' '}
                    {file.split('/').at(-1)}
                  </button>
                ))}
            </section>
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
          placeholder="描述修改、写作要求，或直接提问"
          rows={3}
          maxLength={12_000}
          disabled={state.busy || state.applying}
        />
        <div className="composer-toolbar">
          <div className="composer-tools">
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
              📎
            </button>
            <span className="confirmation-chip">
              <span aria-hidden="true" />
              更改需确认
            </span>
          </div>
          {state.busy ? (
            <button type="button" className="stop-button" onClick={() => session.stop()}>
              Stop
            </button>
          ) : (
            <button
              className="send-button"
              type="button"
              aria-label="Send message"
              disabled={
                !instruction.trim() ||
                uploadPending ||
                state.applying ||
                Boolean(state.proposal) ||
                projectPhase !== 'idle'
              }
              onClick={send}
            >
              ↑
            </button>
          )}
        </div>
      </section>
    </main>
  )
}

export function LegacyAgentWorkspace(props: {
  session: OfficeAgentSession
  ui: OfficeWorkspaceUi
  disconnect: () => void
  host: OfficeHost
}) {
  return <AgentWorkspace {...props} legacy />
}

export function workspaceComponentForMode(mode: 'workspace' | 'legacy') {
  return mode === 'legacy' ? LegacyAgentWorkspace : AgentWorkspace
}

function ConfiguredApp() {
  const document = useMemo(() => createOfficeDocumentClient(createBrowserOfficeRuntime()), [])
  const transportMode = useMemo(() => officeTransportMode(import.meta.env), [])
  const remoteDiagnosticsEnabled = useMemo(
    () => transportMode === 'relay' && officeRemoteDiagnosticsEnabled(import.meta.env),
    [transportMode],
  )
  const bridge = useMemo(
    () =>
      transportMode === 'loopback'
        ? createPcBridgeSession()
        : createOfficeRelaySession({
            capabilities: [
              'agent.v1',
              'presentation.v1',
              'presentation-attachments.v1',
              'presentation-assets.v1',
              'presentation-remote-images.v1',
              'presentation-asset-rights.v1',
            ],
          }),
    [transportMode],
  )
  const bridgeState = useSyncExternalStore(
    (listener) => bridge.subscribe(listener),
    () => bridge.snapshot(),
    () => bridge.snapshot(),
  )
  const [workspace, setWorkspace] = useState<
    { runtime: OfficeHostRuntime; session: OfficeAgentSession; ui: OfficeWorkspaceUi } | undefined
  >()
  const workspaceMode = useMemo(() => officeWorkspaceMode(import.meta.env), [])
  const capabilityFlags = useMemo(() => officeCapabilityFlags(import.meta.env), [])
  const [host, setHost] = useState<OfficeHost>('unknown')
  const [hostSupported, setHostSupported] = useState(false)
  const [status, setStatus] = useState('Connecting to Office…')
  const [busy, setBusy] = useState(true)

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
            const presentationBinding =
              activeHost === 'powerpoint' ? createBrowserPresentationDocumentBinding() : undefined
            const boundPresentationDocumentId = presentationBinding
              ? await presentationBinding.documentId()
              : undefined
            const runCheckpoint = presentationBinding
              ? createPresentationAgentRunCheckpoint(presentationBinding)
              : undefined
            const environment = officeDiagnosticEnvironment(activeHost)
            const diagnostics = createOfficeDiagnostics({
              host: activeHost,
              platform: environment.platform,
              build: __WISWORK_OFFICE_BUILD_ID__,
              requirementSets: environment.requirementSets,
              remoteEnabled: remoteDiagnosticsEnabled,
              send: (event) => {
                if (!('sendDiagnostic' in bridge)) throw new Error('diagnostic_upload_failed')
                return bridge.sendDiagnostic(event)
              },
            })
            const runtime = createOfficeHostRuntime(activeHost, {
              enableHostSkills: import.meta.env.VITE_WISWORK_OFFICE_HOST_SKILLS !== '0',
              enableConversions: capabilityFlags.conversions,
              enableSkillPackages: capabilityFlags.skillPackages,
              enableImportMedia: capabilityFlags.importMedia,
              document,
              diagnostics,
              ...(activeHost === 'powerpoint' && 'capabilityFetch' in bridge
                ? {
                    presentation: {
                      ...presentationBinding!,
                      available: () => {
                        const snapshot = bridge.snapshot()
                        return (
                          snapshot.status === 'connected' &&
                          snapshot.capabilities?.includes('presentation.v1') === true
                        )
                      },
                      request: (body: unknown, signal?: AbortSignal) =>
                        bridge.capabilityFetch('presentation.v1', body, signal),
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
                          snapshot.capabilities?.includes('presentation-remote-images.v1') === true
                        )
                      },
                      rightsAvailable: () => {
                        const snapshot = bridge.snapshot()
                        return (
                          snapshot.status === 'connected' &&
                          snapshot.capabilities?.includes('presentation-asset-rights.v1') === true
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
                                ['attachment_attest_license', 'attachment_revoke_license'].includes(
                                  body.operation as string,
                                )
                              ? 'presentation-asset-rights.v1'
                              : 'presentation-attachments.v1',
                          body,
                          signal,
                        ),
                    },
                  }
                : {}),
            })
            const session = createOfficeAgentSession({
              transport: createPcBridgeAgentTransport(bridge),
              skill: runtime.skill,
              proposals: runtime.proposals,
              diagnostics,
              ...(presentationBinding && boundPresentationDocumentId
                ? {
                    runCheckpoint: {
                      interrupted: presentationBinding.interruptedAgentRun(
                        boundPresentationDocumentId,
                      ),
                      recovery: presentationBinding.agentRunRecovery(boundPresentationDocumentId),
                      validateDocument: async () =>
                        (await presentationBinding.documentId()) === boundPresentationDocumentId,
                      begin: runCheckpoint!.begin,
                      tool: runCheckpoint!.tool,
                      finish: runCheckpoint!.finish,
                    },
                  }
                : {}),
            })
            created = { runtime, session, ui: createOfficeWorkspaceUi(runtime, diagnostics) }
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
      created?.session.dispose()
      created?.runtime.dispose()
      bridge.disconnect()
    }
  }, [bridge, capabilityFlags, document, remoteDiagnosticsEnabled])

  useEffect(() => {
    if (bridgeState.status !== 'connected' && workspace) {
      workspace.session.authenticationLost()
      workspace.runtime.clearSession()
      workspace.runtime.presentation?.prepareReconnect()
    }
  }, [bridgeState.status, workspace])

  if (busy) return <StatusScreen title="Starting WisWork Agent" detail={status} busy />
  if (!hostSupported) {
    return (
      <StatusScreen title="Unsupported Office host" detail="This host cannot use document tools." />
    )
  }
  if (bridgeState.status !== 'connected') {
    const detail = {
      offline: 'Connect again to create a new secure pairing with WisWork PC.',
      connecting: 'Connecting securely to the WisWork Office Relay…',
      incompatible:
        'Upgrade Office Relay to a version that supports this Office add-in, then try again.',
      signed_out: 'Sign in to WisWork PC first.',
      pending: bridgeState.verificationCode
        ? `Enter code ${bridgeState.verificationCode} in WisWork PC, then approve the matching request.`
        : 'Enter the pairing code in WisWork PC.',
      waiting_for_pc: bridgeState.verificationCode
        ? `Enter code ${bridgeState.verificationCode} in WisWork PC to continue.`
        : 'Waiting for a signed-in WisWork PC.',
      rejected: 'The connection was rejected in WisWork PC.',
      expired: 'The connection request expired. Try again.',
    }[bridgeState.status]
    return (
      <StatusScreen
        title="Connect to WisWork PC"
        detail={detail}
        busy={bridgeState.status === 'connecting' || bridgeState.status === 'pending'}
      >
        <button
          type="button"
          disabled={bridgeState.status === 'connecting' || bridgeState.status === 'pending'}
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
          <DiagnosticCopyButton copyDiagnostics={workspace.ui.copyDiagnostics} />
        )}
      </StatusScreen>
    )
  }
  if (!workspace)
    return <StatusScreen title="Starting WisWork Agent" detail="Loading tools…" busy />
  const disconnect = () => {
    workspace.session.logout()
    workspace.runtime.dispose()
    bridge.disconnect()
  }
  const WorkspaceComponent = workspaceComponentForMode(workspaceMode)
  return (
    <WorkspaceComponent
      session={workspace.session}
      ui={workspace.ui}
      disconnect={disconnect}
      host={host}
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
  useEffect(() => {
    let active = true
    void (async () => {
      const deployed = await deployedBuildId()
      if (active) setVersionState(resolveBuildVersion(deployed, __WISWORK_OFFICE_BUILD_ID__))
    })()
    return () => {
      active = false
    }
  }, [])
  if (versionState.status === 'checking')
    return <StatusScreen title="Checking WisWork version" detail="Checking for updates…" busy />
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
