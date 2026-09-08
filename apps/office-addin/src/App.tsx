import { useEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react'
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
import {
  officeCapabilityFlags,
  officeRemoteDiagnosticsEnabled,
  officeWorkspaceMode,
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
    preview: hasComparison || legacy ? '' : previewSummary(proposal.preview),
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
  readonly attachments: () => readonly string[]
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
    attachments: () => Object.freeze([...runtime.vfs.list('/home/user')]),
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
  if (file.size > MAX_VFS_FILE_BYTES) return Promise.reject(new Error('vfs_limit'))
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

function PowerPointTimeline(props: {
  timeline: OfficePresentationTimeline
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
  const latestDesignToolId = [...props.timeline]
    .reverse()
    .find(
      (event) =>
        event.kind === 'tool' &&
        event.name === 'plan_deck' &&
        Boolean(extractPresentationDesignDocument(event.output ?? '')),
    )?.id
  let index = 0
  while (index < props.timeline.length) {
    const event = props.timeline[index]!
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
    while (index < props.timeline.length && props.timeline[index]?.kind === 'tool') {
      const tool = props.timeline[index] as Extract<OfficePresentationEvent, { kind: 'tool' }>
      const designMd =
        tool.name === 'plan_deck' ? extractPresentationDesignDocument(tool.output ?? '') : undefined
      tools.push({
        id: tool.callId,
        label: designMd ? 'DESIGN.md · 已创建' : tool.summary,
        status:
          tool.state === 'running'
            ? ('running' as const)
            : tool.state === 'error'
              ? ('error' as const)
              : ('done' as const),
        ...(designMd
          ? {
              onActivate: () => props.onOpenDesign(designMd, tool.id === latestDesignToolId),
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
}) {
  const { session, ui, disconnect, host } = props
  const state = useOfficeAgent(session)
  const [instruction, setInstruction] = useState('')
  const [files, setFiles] = useState<readonly string[]>(ui.attachments())
  const [skills, setSkills] = useState<readonly string[]>(ui.skills())
  const [uploadError, setUploadError] = useState('')
  const [diagnosticStatus, setDiagnosticStatus] = useState('')
  const [designEditor, setDesignEditor] = useState<{
    designMd: string
    editable: boolean
  }>()
  const [designDraft, setDesignDraft] = useState('')
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

  function send() {
    if (props.connectionAvailable === false || !instruction.trim()) return
    session.send(instruction)
    setInstruction('')
  }

  async function uploadFiles(selected: FileList | readonly File[]) {
    setUploadError('')
    for (const file of Array.from(selected)) {
      try {
        await ui.upload(file)
      } catch (error) {
        if (mounted.current) setUploadError(safeUploadError(error, file))
        break
      }
    }
    if (mounted.current) setFiles(ui.attachments())
  }

  const proposal = state.proposal
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
              setDesignDraft(designMd)
              setDesignEditor({ designMd, editable })
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
                disabled={state.applying}
                onClick={() => session.retry()}
              >
                {state.error === 'proposal_stale' ? '重新生成' : 'Retry'}
              </button>
            )}
          </div>
        )}
      </section>

      {designEditor && (
        <div className="design-dialog-backdrop" role="presentation">
          <section className="design-dialog" role="dialog" aria-modal="true" aria-label="DESIGN.md">
            <header>
              <strong>DESIGN.md</strong>
              <button type="button" className="quiet" onClick={() => setDesignEditor(undefined)}>
                ×
              </button>
            </header>
            <p>
              {designEditor.editable
                ? '修改后将生成新的设计合同节点，并用于后续制作。'
                : '这是历史设计合同快照，仅供查看。'}
            </p>
            <textarea
              value={designDraft}
              readOnly={!designEditor.editable}
              spellCheck={false}
              aria-label="DESIGN.md 内容"
              onChange={(event) => setDesignDraft(event.target.value)}
            />
            <footer>
              <button
                type="button"
                className="secondary"
                onClick={() => setDesignEditor(undefined)}
              >
                关闭
              </button>
              {designEditor.editable && (
                <button
                  type="button"
                  className="primary"
                  disabled={state.busy || state.applying || !session.reviseDesignContract}
                  onClick={() => {
                    const body = designDraft.replace(/^\s*#\s*DESIGN\.md\s*/i, '').trim()
                    if (!body) return
                    session.reviseDesignContract?.(`# DESIGN.md\n\n${body}`)
                    setDesignEditor(undefined)
                  }}
                >
                  保存并应用
                </button>
              )}
            </footer>
          </section>
        </div>
      )}

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
                aria-disabled={state.applying}
              >
                Add attachment
              </label>
              <input
                id="session-upload"
                className="visually-hidden"
                type="file"
                disabled={state.applying}
                onChange={(event) => {
                  const file = event.currentTarget.files?.[0]
                  if (!file) return
                  setUploadError('')
                  void ui
                    .upload(file)
                    .then(() => mounted.current && setFiles(ui.attachments()))
                    .catch(
                      (error: unknown) =>
                        mounted.current && setUploadError(safeUploadError(error, file)),
                    )
                }}
              />
              <p>
                Files are limited to {displayMegabytes(MAX_VFS_FILE_BYTES)} MB each and{' '}
                {displayMegabytes(MAX_VFS_TOTAL_BYTES)} MB per session, then cleared on logout.
              </p>
              {uploadError && (
                <p className="error-text" role="alert">
                  {uploadError}
                </p>
              )}
              <ul>
                {files.map((file) => (
                  <li key={file}>{file.split('/').at(-1)}</li>
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
                disabled={state.applying}
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
                disabled={state.applying}
                onClick={() => composerFileInput.current?.click()}
              >
                <IconPaperclip size={20} />
              </button>
              <span className="confirmation-chip">
                <span aria-hidden="true" />
                {host === 'powerpoint' ? '自动应用常规更改' : '更改需确认'}
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
                  props.connectionAvailable === false || !instruction.trim() || state.applying
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
  const remoteDiagnosticsEnabled = useMemo(
    () => transportMode === 'relay' && officeRemoteDiagnosticsEnabled(import.meta.env),
    [transportMode],
  )
  const bridge = useMemo(
    () =>
      props.connectionBridge ??
      (transportMode === 'loopback'
        ? createPcBridgeSession()
        : createOfficeRelaySession({
            capabilities: ['agent.v1', 'web-search.v1', 'image-search.v1', 'image-fetch.v1'],
            persistentPairing: __WISWORK_OFFICE_PAIRING_RESUME__,
          })),
    [props.connectionBridge, transportMode],
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
  const presentationFlags = useMemo(() => officePresentationVerificationFlags(import.meta.env), [])
  const [host, setHost] = useState<OfficeHost>('unknown')
  const [hostSupported, setHostSupported] = useState(false)
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
              })
              const session = createOfficeAgentSession({
                transport: createPcBridgeAgentTransport(bridge),
                skill: runtime.skill,
                proposals: runtime.proposals,
                automaticPowerPointMutations: activeHost === 'powerpoint',
                ...('setToolHandler' in bridge ? { remoteTools: bridge } : {}),
                diagnostics,
                presentationText: (key) =>
                  officePresentationText(globalThis.Office?.context?.displayLanguage, key),
              })
              created = { runtime, session, ui: createOfficeWorkspaceUi(runtime, diagnostics) }
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
      created?.session.dispose()
      created?.runtime.dispose()
      bridge.disconnect()
    }
  }, [
    bridge,
    capabilityFlags,
    document,
    presentationFlags,
    remoteDiagnosticsEnabled,
    workspaceFactory,
  ])

  useEffect(() => {
    if (shouldResetOfficeSession(bridgeState.status) && workspace) {
      workspace.session.authenticationLost()
      workspace.runtime.clearSession()
    }
  }, [bridgeState.status, workspace])

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
          <DiagnosticCopyButton copyDiagnostics={workspace.ui.copyDiagnostics} />
        )}
      </StatusScreen>
    )
  }
  if (!workspace)
    return <StatusScreen title="Starting WisWork Agent" detail="Loading tools…" busy />
  const disconnect = () => {
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
  return <ConfiguredApp />
}
