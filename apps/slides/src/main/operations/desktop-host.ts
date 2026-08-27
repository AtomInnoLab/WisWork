import type { PresentationOperation, PresentationTarget } from '@wiswork/presentation-ops'
import {
  addElement,
  collectDeckCreationIds,
  deleteElement,
  fingerprintPresentation,
  ensureElementCreationId,
  editGroupChildTransform,
  getSlideNotes,
  mintUniqueCreationIds,
  resolvePresentationContainer,
  resolvePresentationTarget,
  setSlideNotes,
  type Fill,
  type GroupElement,
  type OpenedPptx,
  type PackageArchive,
  type Slide,
  type SlideElement,
  type Stroke,
  type TextElement,
} from '@wiswork/pptx-engine'
import { commitHistorySnapshot, type HistorySnapshot, type Session } from '../session-state'
import { applyEditParagraphs } from '../edit-text'
import type { AtomicPresentationHost, PlannedPresentationOperation } from './executor'
import type { PresentationPlan } from './planner'

const EMU_PER_POINT = 12_700

interface DesktopSnapshot {
  history: HistorySnapshot
  metaDirty: boolean
  mutationGeneration: number
}

function cloneArchive(source: PackageArchive, entries: Map<string, Uint8Array>): PackageArchive {
  const clone = Object.create(Object.getPrototypeOf(source)) as PackageArchive
  Object.defineProperties(clone, Object.getOwnPropertyDescriptors(source))
  Object.defineProperty(clone, 'entries', {
    configurable: true,
    enumerable: true,
    writable: false,
    value: new Map(entries),
  })
  return clone
}

function openedFromSnapshot(session: Session, snapshot: DesktopSnapshot): OpenedPptx {
  return {
    deck: {
      ...session.opened.deck,
      slides: structuredClone(snapshot.history.slides),
      size: { ...snapshot.history.size },
    },
    archive: cloneArchive(session.opened.archive, snapshot.history.entries),
  }
}

function findSlide(opened: OpenedPptx, slideId: string): Slide | undefined {
  const matches = opened.deck.slides.filter((slide) => slide.durableId === slideId)
  return matches.length === 1 ? matches[0] : undefined
}

function findTopLevelElement(slide: Slide, elementId: string): SlideElement | undefined {
  const matches: SlideElement[] = []
  const visit = (elements: readonly SlideElement[]) => {
    for (const element of elements) {
      if (element.creationId === elementId) matches.push(element)
      if (element.type === 'group') visit(element.children)
    }
  }
  visit(slide.elements)
  return matches.length === 1 ? matches[0] : undefined
}

function findElementLocation(
  slide: Slide,
  elementId: string,
): { element: SlideElement; directParent?: GroupElement } | undefined {
  const matches: Array<{ element: SlideElement; directParent?: GroupElement }> = []
  const visit = (elements: readonly SlideElement[], parent?: GroupElement, depth = 0) => {
    for (const element of elements) {
      if (element.creationId === elementId)
        matches.push({ element, ...(parent && depth === 1 ? { directParent: parent } : {}) })
      if (element.type === 'group') visit(element.children, element, depth + 1)
    }
  }
  visit(slide.elements)
  return matches.length === 1 ? matches[0] : undefined
}

export interface PresentationTargetEnrollment {
  slideId: string
  sourceId: string
  elementId: string
}

function findLegacyElement(slide: Slide, sourceId: string): SlideElement | undefined {
  const matches: SlideElement[] = []
  const visit = (elements: readonly SlideElement[]) => {
    for (const element of elements) {
      if (element.id === sourceId) matches.push(element)
      if (element.type === 'group') visit(element.children)
    }
  }
  visit(slide.elements)
  return matches.length === 1 ? matches[0] : undefined
}

function replaceText(
  element: TextElement,
  operation: Extract<PresentationOperation, { kind: 'set_text' }>,
): boolean {
  const value = 'text' in operation ? operation.text : undefined
  const firstRun = element.text?.paragraphs[0]?.runs[0]
  const firstParagraph = element.text?.paragraphs[0]
  element.text ??= { paragraphs: [] }
  const next =
    value !== undefined
      ? value.split('\n').map((text) => ({
          ...(firstParagraph
            ? {
                ...firstParagraph,
                runs: [{ ...(firstRun ?? {}), text }],
              }
            : { runs: [{ text }] }),
        }))
      : applyEditParagraphs(
          element.text.paragraphs,
          operation.paragraphs!.map((paragraph) => ({
            ...paragraph,
            runs: paragraph.runs.map((run) => ({ ...run })),
          })),
        )
  if (same(element.text.paragraphs, next)) return false
  element.text.paragraphs = next
  element.dirty = true
  return true
}

function fillValue(operation: Extract<PresentationOperation, { kind: 'set_fill' }>): Fill {
  if (operation.fill.kind === 'none') return { type: 'none' }
  const alpha = Math.round(255 * (1 - (operation.fill.transparency ?? 0)))
    .toString(16)
    .padStart(2, '0')
    .toUpperCase()
  return {
    type: 'solid',
    color: alpha === 'FF' ? operation.fill.color : `${operation.fill.color}${alpha}`,
  }
}

function strokeValue(operation: Extract<PresentationOperation, { kind: 'set_stroke' }>): Stroke {
  return {
    fill: { type: 'solid', color: operation.stroke.color },
    width: Math.round(operation.stroke.width * EMU_PER_POINT),
    ...(operation.stroke.dash === undefined
      ? {}
      : { dash: operation.stroke.dash === 'dash_dot' ? 'dashDot' : operation.stroke.dash }),
  }
}

function same(valueA: unknown, valueB: unknown): boolean {
  return JSON.stringify(valueA) === JSON.stringify(valueB)
}

function targetId(operation: PresentationOperation): string | undefined {
  return 'target' in operation ? operation.target.elementId : undefined
}

function applyOperation(
  opened: OpenedPptx,
  planned: PlannedPresentationOperation,
): { changed: boolean; metaDirty: boolean } {
  const operation = planned.operation
  const slideId = operation.kind === 'add_text_box' ? operation.slideId : operation.target.slideId
  const slide = findSlide(opened, slideId)
  if (!slide) throw new Error('Planned slide disappeared')
  if (operation.kind === 'add_text_box') {
    const id = planned.createdId
    if (!id) throw new Error('Planned insertion has no durable id')
    addElement(
      slide,
      {
        kind: 'textbox',
        offset: {
          x: Math.round(operation.geometry.x * EMU_PER_POINT),
          y: Math.round(operation.geometry.y * EMU_PER_POINT),
          cx: Math.round(operation.geometry.width * EMU_PER_POINT),
          cy: Math.round(operation.geometry.height * EMU_PER_POINT),
        },
        paragraphs: operation.text.split('\n').map((text) => ({ runs: [{ text }] })),
      },
      { creationIdFactory: () => id },
    )
    return { changed: true, metaDirty: false }
  }
  if (operation.kind === 'set_speaker_notes') {
    const index = opened.deck.slides.indexOf(slide)
    if (getSlideNotes(opened.archive, slide.path) === operation.notes)
      return { changed: false, metaDirty: false }
    if (!setSlideNotes(opened, index, operation.notes))
      throw new Error('Speaker notes write failed')
    return { changed: true, metaDirty: true }
  }
  const elementId = operation.target.elementId
  if (!elementId) throw new Error('Planned element target disappeared')
  const location = findElementLocation(slide, elementId)
  const element = location?.element
  if (!element) throw new Error('Planned element target disappeared')
  switch (operation.kind) {
    case 'set_text':
      if (element.type !== 'text' && element.type !== 'shape')
        throw new Error('Text operation target changed type')
      return { changed: replaceText(element, operation), metaDirty: false }
    case 'set_geometry': {
      const next = {
        x: Math.round(operation.geometry.x * EMU_PER_POINT),
        y: Math.round(operation.geometry.y * EMU_PER_POINT),
        cx: Math.round(operation.geometry.width * EMU_PER_POINT),
        cy: Math.round(operation.geometry.height * EMU_PER_POINT),
      }
      const rotation = Math.round(
        (operation.geometry.rotation ?? element.transform.rot / 60_000) * 60_000,
      )
      if (same(element.transform.offset, next) && element.transform.rot === rotation)
        return { changed: false, metaDirty: false }
      if (location.directParent) {
        if (
          !editGroupChildTransform(
            slide,
            location.directParent.id,
            element.id,
            next,
            rotation / 60_000,
          )
        )
          throw new Error('Group child transform failed')
        return { changed: true, metaDirty: false }
      }
      element.transform.offset = next
      element.transform.rot = rotation
      element.dirtyTransform = true
      return { changed: true, metaDirty: false }
    }
    case 'set_fill': {
      if (element.type !== 'text' && element.type !== 'shape')
        throw new Error('Fill operation target changed type')
      const next = fillValue(operation)
      if (same(element.fill, next)) return { changed: false, metaDirty: false }
      element.fill = next
      element.dirtyFill = true
      return { changed: true, metaDirty: false }
    }
    case 'set_stroke': {
      if (element.type !== 'text' && element.type !== 'shape' && element.type !== 'picture')
        throw new Error('Stroke operation target changed type')
      const next = strokeValue(operation)
      if (same(element.stroke, next)) return { changed: false, metaDirty: false }
      element.stroke = next
      element.dirtyStroke = true
      return { changed: true, metaDirty: false }
    }
    case 'delete_element':
      if (!deleteElement(slide, element.id)) throw new Error('Element deletion failed')
      return { changed: true, metaDirty: false }
  }
}

function conflictTarget(operation: PresentationOperation): PresentationTarget {
  return operation.kind === 'add_text_box' ? { slideId: operation.slideId } : operation.target
}

export class DesktopPresentationHost implements AtomicPresentationHost<DesktopSnapshot> {
  private readonly reservedIds: Set<string>
  private expectedRevisions: string[] = []
  private finalMetaDirty = false
  private transactionGeneration = 0

  constructor(
    private readonly session: Session,
    private readonly enrollmentForTarget?: (
      elementId: string,
    ) => PresentationTargetEnrollment | undefined,
  ) {
    this.reservedIds = collectDeckCreationIds(session.opened.deck)
  }

  private enrollTarget(opened: OpenedPptx, target: PresentationTarget): boolean {
    if (!target.elementId) return false
    const slide = findSlide(opened, target.slideId)
    if (!slide || findTopLevelElement(slide, target.elementId)) return false
    const enrollment = this.enrollmentForTarget?.(target.elementId)
    if (!enrollment || enrollment.slideId !== target.slideId) return false
    const element = findLegacyElement(slide, enrollment.sourceId)
    if (!element || (element.creationId && element.creationId !== enrollment.elementId))
      return false
    ensureElementCreationId(slide, element, () => enrollment.elementId)
    return true
  }

  readRevision(): Promise<string> {
    return fingerprintPresentation(this.session.opened)
  }

  async captureSnapshot(): Promise<DesktopSnapshot> {
    return {
      history: {
        slides: structuredClone(this.session.opened.deck.slides),
        entries: new Map(this.session.opened.archive.entries),
        size: { ...this.session.opened.deck.size },
      },
      metaDirty: this.session.metaDirty === true,
      mutationGeneration: this.session.mutationGeneration ?? 0,
    }
  }

  async plan(
    snapshot: DesktopSnapshot,
    operations: readonly PresentationOperation[],
    allocateId: (clientId: string) => string,
  ): Promise<PresentationPlan> {
    this.transactionGeneration = snapshot.mutationGeneration
    // Every external precondition is resolved against the same authoritative snapshot.
    const authoritative = openedFromSnapshot(this.session, snapshot)
    for (const [index, operation] of operations.entries()) {
      if (operation.kind !== 'add_text_box') this.enrollTarget(authoritative, operation.target)
      const resolution =
        operation.kind === 'add_text_box'
          ? await resolvePresentationContainer(authoritative.deck, conflictTarget(operation))
          : await resolvePresentationTarget(authoritative, conflictTarget(operation))
      if (resolution.status === 'conflict') {
        return {
          status: 'conflict',
          code: resolution.code,
          operationIndex: index,
          ...(targetId(operation) ? { targetId: targetId(operation) } : {}),
        }
      }
      if (operation.kind === 'set_fill' && resolution.element?.type === 'picture') {
        return {
          status: 'conflict',
          code: 'target_stale',
          operationIndex: index,
          ...(targetId(operation) ? { targetId: targetId(operation) } : {}),
        }
      }
    }

    const simulated = openedFromSnapshot(this.session, snapshot)
    for (const operation of operations) {
      if (operation.kind !== 'add_text_box') this.enrollTarget(simulated, operation.target)
    }
    const initialSimulatedRevision = await fingerprintPresentation(simulated)
    const planned = operations.map((operation, index) => ({
      index,
      operation,
      ...(operation.kind === 'add_text_box' ? { createdId: allocateId(operation.clientId) } : {}),
    }))
    const revisions: string[] = []
    let changed = false
    let metaDirty = false
    try {
      for (const item of planned) {
        const result = applyOperation(simulated, item)
        changed ||= result.changed
        metaDirty ||= result.metaDirty
        revisions.push(await fingerprintPresentation(simulated))
      }
    } catch {
      return { status: 'conflict', code: 'target_stale' }
    }
    this.expectedRevisions = revisions
    this.finalMetaDirty = metaDirty
    return {
      status: 'planned',
      operations: planned,
      noOp: !changed || revisions.at(-1) === initialSimulatedRevision,
    }
  }

  allocateElementId(): string {
    for (const id of collectDeckCreationIds(this.session.opened.deck)) this.reservedIds.add(id)
    const id = mintUniqueCreationIds(1, this.reservedIds)[0]!
    this.reservedIds.add(id)
    return id
  }

  async apply(planned: PlannedPresentationOperation): Promise<{ revision: string }> {
    const generation = this.session.mutationGeneration ?? 0
    if (generation !== this.transactionGeneration)
      throw new Error('Presentation transaction lease changed')
    if (planned.operation.kind !== 'add_text_box')
      this.enrollTarget(this.session.opened, planned.operation.target)
    applyOperation(this.session.opened, planned)
    const revision = await fingerprintPresentation(this.session.opened)
    if ((this.session.mutationGeneration ?? 0) !== generation)
      throw new Error('Concurrent presentation mutation')
    if (revision !== this.expectedRevisions[planned.index])
      throw new Error('Transaction state diverged')
    return { revision }
  }

  async verify(): Promise<{ status: 'matched'; revision: string } | { status: 'mismatch' }> {
    const generation = this.session.mutationGeneration ?? 0
    if (generation !== this.transactionGeneration) return { status: 'mismatch' }
    const revision = await fingerprintPresentation(this.session.opened)
    if ((this.session.mutationGeneration ?? 0) !== generation) return { status: 'mismatch' }
    return revision === this.expectedRevisions.at(-1)
      ? { status: 'matched', revision }
      : { status: 'mismatch' }
  }

  async isAttributableRevision(revision: string): Promise<boolean> {
    return this.expectedRevisions.includes(revision)
  }

  async restoreIfCurrent(
    _expectedCurrentRevision: string,
    _snapshot: DesktopSnapshot,
  ): Promise<'unsupported'> {
    // Ordinary editor IPC does not yet share a mutation generation/CAS lock.
    // An async fingerprint followed by restore could overwrite a third state,
    // so desktop recovery must fail closed until that host-wide primitive exists.
    return 'unsupported'
  }

  async publishHistory(snapshot: DesktopSnapshot): Promise<boolean> {
    // No await before compare+commit: this is the host-wide synchronous CAS.
    if ((this.session.mutationGeneration ?? 0) !== snapshot.mutationGeneration) return false
    commitHistorySnapshot(this.session, snapshot.history)
    if (this.finalMetaDirty) this.session.metaDirty = true
    return true
  }
}
