import {
  extractPresentationDesignContract,
  extractPresentationDesignDocument,
} from '@wiswork/agent-core'

export function presentationDesignLifecycle(output: string) {
  const designMd = extractPresentationDesignDocument(output) ?? output
  const contract = extractPresentationDesignContract(designMd)
  if (!contract) return { label: 'DESIGN.md · created', editable: true }
  return {
    label:
      contract.status === 'verified'
        ? `DESIGN.md · verified · r${contract.revision}`
        : contract.status === 'producing'
          ? `DESIGN.md · locked · r${contract.revision}`
          : contract.revision > 1
            ? `DESIGN.md · revised · r${contract.revision}`
            : `DESIGN.md · created · r${contract.revision}`,
    editable: contract.status === 'draft' || contract.status === 'ready',
  }
}
