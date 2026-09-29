import type { StructuredProposal } from '../../agent/proposal-controller.js'
import { presentationQaMutationScope } from './presentation-qa.js'

/** Only internal operations with stable host identities may narrow the affected scope. */
export function presentationMutationScope(proposal: StructuredProposal): string[] | undefined {
  const target = proposal.impact.targets[0]
  let hostSlideIds =
    [
      'edit_existing_presentation_text',
      'edit_existing_presentation_geometry',
      'undo_existing_presentation_change',
      'resume_existing_presentation_change',
      'reapply_existing_presentation_change',
      'replace_existing_presentation_image',
      'resume_existing_presentation_image_change',
      'reapply_existing_presentation_image_change',
      'undo_existing_presentation_image_change',
      'stage_existing_presentation_page_change',
      'reapply_existing_presentation_page_change',
      'resume_existing_presentation_page_change',
      'commit_existing_presentation_page_change',
      'discard_existing_presentation_page_change',
      'undo_existing_presentation_page_change',
      'stage_presentation_page_replacement',
      'resume_presentation_page_replacement',
      'discard_presentation_page_replacement',
      'commit_presentation_page_replacement',
      'undo_presentation_page_replacement',
      'edit_presentation_page_text',
      'undo_presentation_text_change',
      'resume_presentation_text_change',
      'edit_presentation_page_geometry',
      'undo_presentation_geometry_change',
      'resume_presentation_geometry_change',
      'undo_presentation_image_replacement',
      'replace_presentation_page_image',
      'resume_presentation_image_replacement',
    ].includes(proposal.operation) &&
    proposal.operation === proposal.toolName &&
    proposal.impact.host === 'powerpoint' &&
    proposal.impact.count === 1 &&
    proposal.impact.targets.length === 1 &&
    typeof target === 'string' &&
    target.length > 0 &&
    target.length <= 256 &&
    !Array.from(target).some(
      (char) => char.charCodeAt(0) < 32 || (char.charCodeAt(0) >= 127 && char.charCodeAt(0) <= 159),
    )
      ? [target]
      : undefined
  if (
    [
      'edit_existing_presentation_batch',
      'resume_existing_presentation_batch',
      'reapply_existing_presentation_batch',
      'undo_existing_presentation_batch',
    ].includes(proposal.operation) &&
    proposal.operation === proposal.toolName &&
    proposal.impact.host === 'powerpoint' &&
    proposal.impact.count >= 2 &&
    proposal.impact.count <= 8 &&
    proposal.impact.targets.length >= 1 &&
    proposal.impact.targets.length <= 8
  ) {
    try {
      hostSlideIds = [...presentationQaMutationScope(proposal.impact.targets)!]
    } catch {
      hostSlideIds = undefined
    }
  }
  // Generic modifications derive these IDs from exact original-page savepoints.
  // Legacy index labels and caller-supplied programs cannot establish this proof.
  if (
    ['execute_office_js', 'resume_native_modify_batch'].includes(proposal.operation) &&
    proposal.operation === proposal.toolName &&
    proposal.impact.host === 'powerpoint' &&
    proposal.impact.count >= 1 &&
    proposal.impact.count <= 32 &&
    proposal.impact.targets.length >= 1 &&
    proposal.impact.targets.length <= 8
  ) {
    const scope = proposal.preview.qaScope
    if (
      scope &&
      typeof scope === 'object' &&
      !Array.isArray(scope) &&
      Object.keys(scope).sort().join(',') === 'basis,hostSlideIds' &&
      (scope as Record<string, unknown>).basis === 'native_modify_savepoints' &&
      Array.isArray((scope as Record<string, unknown>).hostSlideIds) &&
      JSON.stringify((scope as Record<string, unknown>).hostSlideIds) ===
        JSON.stringify(proposal.impact.targets)
    ) {
      try {
        hostSlideIds = [...presentationQaMutationScope(proposal.impact.targets)!]
      } catch {
        hostSlideIds = undefined
      }
    }
  }
  // Duplication saves one source page; removal also proves the owned copied page.
  if (
    ['duplicate_slide', 'execute_office_js', 'undo_slide_duplication'].includes(
      proposal.operation,
    ) &&
    proposal.operation === proposal.toolName &&
    proposal.impact.host === 'powerpoint' &&
    proposal.impact.count === 1 &&
    proposal.impact.targets.length >= 1 &&
    proposal.impact.targets.length <= (proposal.operation === 'undo_slide_duplication' ? 2 : 1)
  ) {
    const scope = proposal.preview.qaScope
    if (
      scope &&
      typeof scope === 'object' &&
      !Array.isArray(scope) &&
      Object.keys(scope).sort().join(',') === 'basis,hostSlideIds' &&
      (scope as Record<string, unknown>).basis === 'slide_duplication_savepoint' &&
      JSON.stringify((scope as Record<string, unknown>).hostSlideIds) ===
        JSON.stringify(proposal.impact.targets)
    ) {
      try {
        hostSlideIds = [...presentationQaMutationScope(proposal.impact.targets)!]
      } catch {
        hostSlideIds = undefined
      }
    }
  }
  // Only the native master tool derives this scope from a complete, revalidated
  // host dependency snapshot. XML/package edits and generic labels remain unknown.
  if (
    ['edit_slide_master', 'resume_slide_master_change', 'undo_slide_master_change'].includes(
      proposal.operation,
    ) &&
    proposal.toolName === proposal.operation &&
    proposal.impact.host === 'powerpoint'
  ) {
    const scope = proposal.preview.qaScope
    if (
      scope &&
      typeof scope === 'object' &&
      !Array.isArray(scope) &&
      Object.keys(scope).length === 2 &&
      Object.keys(scope).every((key) => ['basis', 'hostSlideIds'].includes(key)) &&
      (scope as Record<string, unknown>).basis === 'native_master_layout' &&
      Array.isArray((scope as Record<string, unknown>).hostSlideIds)
    ) {
      try {
        hostSlideIds = [
          ...presentationQaMutationScope((scope as { hostSlideIds: string[] }).hostSlideIds)!,
        ]
      } catch {
        hostSlideIds = undefined
      }
    }
  }
  return hostSlideIds
}
