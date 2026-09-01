import { useCallback, useEffect, useRef, useState } from 'react'

const MAX_ID = 160
const MAX_TOOL = 96

const OPERATIONS = ['insert', 'replace', 'delete', 'format', 'restructure', 'compile'] as const
const TARGETS = [
  'document',
  'selection',
  'blocks',
  'cells',
  'sheet',
  'slides',
  'elements',
  'project-files',
] as const
const SCOPES = ['single', 'selection', 'bounded-set', 'whole-document'] as const

export type EnhancedMutationOperation = (typeof OPERATIONS)[number]
export type EnhancedMutationTarget = (typeof TARGETS)[number]
export type EnhancedMutationScope = (typeof SCOPES)[number]

export interface EnhancedMutationSummary {
  readonly operation: EnhancedMutationOperation
  readonly target: EnhancedMutationTarget
  readonly scope: EnhancedMutationScope
  readonly count?: number
}

export interface EnhancedMutationProposal {
  readonly proposalId: string
  readonly documentId: string
  readonly generation: number
  readonly toolName: string
  readonly summary: EnhancedMutationSummary
  readonly expiresAt: number
}

export interface EnhancedMutationProposalApi {
  onProposal(listener: (proposal: EnhancedMutationProposal) => void): () => void
  confirmProposal(documentId: string, generation: number, proposalId: string): Promise<void>
  cancelProposal(documentId: string, generation: number, proposalId: string): Promise<void>
}

function validText(value: unknown, maximum: number): value is string {
  return typeof value === 'string' && value.length > 0 && value.length <= maximum
}

function hasValidEnvelope(value: unknown): value is Omit<EnhancedMutationProposal, 'summary'> & {
  summary: unknown
} {
  if (!value || typeof value !== 'object') return false
  const proposal = value as Record<string, unknown>
  return (
    Object.keys(proposal).every((key) =>
      ['proposalId', 'documentId', 'generation', 'toolName', 'summary', 'expiresAt'].includes(key),
    ) &&
    validText(proposal.proposalId, MAX_ID) &&
    validText(proposal.documentId, MAX_ID) &&
    Number.isSafeInteger(proposal.generation) &&
    (proposal.generation as number) >= 0 &&
    validText(proposal.toolName, MAX_TOOL) &&
    Number.isSafeInteger(proposal.expiresAt) &&
    (proposal.expiresAt as number) > 0
  )
}

function isEnhancedMutationSummary(value: unknown): value is EnhancedMutationSummary {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const summary = value as Record<string, unknown>
  return (
    Object.keys(summary).every((key) => ['operation', 'target', 'scope', 'count'].includes(key)) &&
    OPERATIONS.includes(summary.operation as EnhancedMutationOperation) &&
    TARGETS.includes(summary.target as EnhancedMutationTarget) &&
    SCOPES.includes(summary.scope as EnhancedMutationScope) &&
    (summary.count === undefined ||
      (Number.isSafeInteger(summary.count) &&
        (summary.count as number) >= 1 &&
        (summary.count as number) <= 10_000))
  )
}

export function isEnhancedMutationProposal(value: unknown): value is EnhancedMutationProposal {
  return hasValidEnvelope(value) && isEnhancedMutationSummary(value.summary)
}

const EN = {
  title: 'Confirm document change',
  operation: 'Operation',
  target: 'Target',
  scope: 'Scope',
  count: 'Items affected',
  warning:
    'Review this request carefully. WisWork will apply it as one bounded transaction only after you confirm.',
  reject: 'Reject',
  confirm: 'Confirm change',
  operations: {
    insert: 'Insert',
    replace: 'Replace',
    delete: 'Delete',
    format: 'Format',
    restructure: 'Restructure',
    compile: 'Compile',
  },
  targets: {
    document: 'Document',
    selection: 'Selection',
    blocks: 'Blocks',
    cells: 'Cells',
    sheet: 'Sheet',
    slides: 'Slides',
    elements: 'Elements',
    'project-files': 'Project files',
  },
  scopes: {
    single: 'Single item',
    selection: 'Current selection',
    'bounded-set': 'Bounded set',
    'whole-document': 'Whole document',
  },
} as const

const ZH = {
  title: '确认文档更改',
  operation: '操作类型',
  target: '目标',
  scope: '影响范围',
  count: '影响数量',
  warning: '请仔细检查本次操作。只有在你确认后，WisWork 才会将其作为一个受限事务执行。',
  reject: '拒绝',
  confirm: '确认更改',
  operations: {
    insert: '插入',
    replace: '替换',
    delete: '删除',
    format: '格式调整',
    restructure: '结构调整',
    compile: '编译',
  },
  targets: {
    document: '文档',
    selection: '选区',
    blocks: '内容块',
    cells: '单元格',
    sheet: '工作表',
    slides: '幻灯片',
    elements: '页面元素',
    'project-files': '项目文件',
  },
  scopes: {
    single: '单个对象',
    selection: '当前选区',
    'bounded-set': '受限集合',
    'whole-document': '整个文档',
  },
} as const

/**
 * Renderer-only consent surface. It never receives tool arguments and never executes a writer;
 * confirmation merely asks the privileged owner to claim its exact pending proposal.
 */
export function EnhancedMutationConfirmation({ api }: { api?: EnhancedMutationProposalApi }) {
  const [pending, setPending] = useState<EnhancedMutationProposal | null>(null)
  const [submitting, setSubmitting] = useState(false)
  const pendingRef = useRef<EnhancedMutationProposal | null>(null)
  const consumedRef = useRef(new Set<string>())
  const mountedRef = useRef(true)

  const consume = useCallback(
    async (action: 'confirm' | 'cancel', expected?: EnhancedMutationProposal) => {
      const proposal = pendingRef.current
      if (!api || !proposal || (expected && proposal !== expected)) return
      const key = `${proposal.documentId}\0${proposal.generation}\0${proposal.proposalId}`
      if (consumedRef.current.has(key)) return
      consumedRef.current.add(key)
      if (mountedRef.current) setSubmitting(true)
      pendingRef.current = null
      if (mountedRef.current) setPending(null)
      try {
        const method = action === 'confirm' ? api.confirmProposal : api.cancelProposal
        await method.call(api, proposal.documentId, proposal.generation, proposal.proposalId)
      } catch {
        // The privileged owner remains fail-closed. The renderer must not retry a consumed
        // consent token or infer whether the pending transaction still exists.
      } finally {
        if (mountedRef.current) setSubmitting(false)
      }
    },
    [api],
  )

  useEffect(() => {
    mountedRef.current = true
    if (!api) return
    const unsubscribe = api.onProposal((candidate) => {
      if (!hasValidEnvelope(candidate)) return
      if (!isEnhancedMutationSummary(candidate.summary)) {
        void api.cancelProposal(candidate.documentId, candidate.generation, candidate.proposalId)
        return
      }
      if (candidate.expiresAt <= Date.now()) {
        void api.cancelProposal(candidate.documentId, candidate.generation, candidate.proposalId)
        return
      }
      const key = `${candidate.documentId}\0${candidate.generation}\0${candidate.proposalId}`
      if (consumedRef.current.has(key)) return
      const previous = pendingRef.current
      if (previous) {
        const previousKey = `${previous.documentId}\0${previous.generation}\0${previous.proposalId}`
        if (previousKey === key) return
        void consume('cancel', previous)
      }
      pendingRef.current = candidate
      setPending(candidate)
    })
    return () => {
      mountedRef.current = false
      unsubscribe()
      const current = pendingRef.current
      if (current) void consume('cancel', current)
    }
  }, [api, consume])

  useEffect(() => {
    if (!pending) return
    const remaining = pending.expiresAt - Date.now()
    if (remaining <= 0) {
      void consume('cancel', pending)
      return
    }
    const timer = window.setTimeout(() => void consume('cancel', pending), remaining)
    return () => window.clearTimeout(timer)
  }, [consume, pending])

  if (!pending) return null
  const text = document.documentElement.lang.toLowerCase().startsWith('zh') ? ZH : EN
  return (
    <div className="enhanced-confirm-backdrop" role="presentation">
      <section
        className="enhanced-confirm-dialog"
        role="alertdialog"
        aria-modal="true"
        aria-labelledby="enhanced-confirm-title"
        aria-describedby="enhanced-confirm-summary"
      >
        <h2 id="enhanced-confirm-title">{text.title}</h2>
        <dl id="enhanced-confirm-summary" className="enhanced-confirm-summary">
          <div>
            <dt>{text.operation}</dt>
            <dd>{text.operations[pending.summary.operation]}</dd>
          </div>
          <div>
            <dt>{text.target}</dt>
            <dd>{text.targets[pending.summary.target]}</dd>
          </div>
          <div>
            <dt>{text.scope}</dt>
            <dd>{text.scopes[pending.summary.scope]}</dd>
          </div>
          {pending.summary.count !== undefined && (
            <div>
              <dt>{text.count}</dt>
              <dd>{pending.summary.count}</dd>
            </div>
          )}
        </dl>
        <p className="enhanced-confirm-warning">{text.warning}</p>
        <div className="enhanced-confirm-actions">
          <button
            type="button"
            data-action="cancel"
            disabled={submitting}
            onClick={() => void consume('cancel')}
          >
            {text.reject}
          </button>
          <button
            type="button"
            data-action="confirm"
            disabled={submitting}
            onClick={() => void consume('confirm')}
          >
            {text.confirm}
          </button>
        </div>
      </section>
    </div>
  )
}
