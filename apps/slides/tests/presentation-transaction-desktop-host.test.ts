import { describe, expect, it, vi } from 'vitest'

vi.mock('../src/main/session-state', () => ({
  commitHistorySnapshot: (session: { undoStack: unknown[] }, snapshot: unknown) => {
    session.undoStack.push(snapshot)
  },
  restoreSnapshot: (
    session: {
      opened: {
        deck: { slides: unknown[]; size: unknown }
        archive: { entries: Map<string, Uint8Array> }
      }
    },
    snapshot: { slides: unknown[]; size: unknown; entries: Map<string, Uint8Array> },
  ) => {
    session.opened.deck.slides = structuredClone(snapshot.slides)
    session.opened.deck.size = structuredClone(snapshot.size)
    session.opened.archive.entries.clear()
    for (const [key, value] of snapshot.entries) session.opened.archive.entries.set(key, value)
  },
}))
import {
  addElement,
  addPicture,
  createBlankPptx,
  fingerprintPresentation,
  fingerprintSlideElement,
  openPptx,
  savePptx,
} from '@wiswork/pptx-engine'
import type { PresentationTransaction } from '@wiswork/presentation-ops'
import { DesktopPresentationHost } from '../src/main/operations/desktop-host'
import { PresentationTransactionExecutor } from '../src/main/operations/executor'
import type { Session } from '../src/main/session-state'

async function fixture() {
  const opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  const element = addElement(slide, {
    kind: 'textbox',
    offset: { x: 12_700, y: 25_400, cx: 127_000, cy: 50_800 },
    paragraphs: [{ runs: [{ text: 'before', bold: true }] }],
  })
  const session: Session = {
    path: '',
    opened,
    fitWidthPx: 960,
    undoStack: [],
    redoStack: [],
  }
  return { session, slide, element }
}

describe('DesktopPresentationHost', () => {
  it('preflights repeated targets against one snapshot and commits one undo entry', async () => {
    const { session, slide, element } = await fixture()
    const elementFingerprint = await fingerprintSlideElement(session.opened, slide, element)
    const tx: PresentationTransaction = {
      transactionId: 'desktop-dependent',
      expectedDeckRevision: await fingerprintPresentation(session.opened),
      mode: 'atomic',
      operations: [
        {
          kind: 'set_text',
          clientId: 'text',
          target: {
            slideId: slide.durableId,
            elementId: element.creationId!,
            expectedType: 'text',
            expectedFingerprint: elementFingerprint,
          },
          text: 'after',
        },
        {
          kind: 'set_geometry',
          clientId: 'geometry',
          target: {
            slideId: slide.durableId,
            elementId: element.creationId!,
            expectedType: 'text',
            expectedFingerprint: elementFingerprint,
          },
          geometry: { x: 10, y: 20, width: 100, height: 40, rotation: 15 },
        },
      ],
    }
    const receipt = await new PresentationTransactionExecutor(
      new DesktopPresentationHost(session),
      { verifyDelayMs: 0 },
    ).execute(tx)
    expect(receipt).toMatchObject({ status: 'applied', operationCount: 2 })
    expect(session.undoStack).toHaveLength(1)
    const current = session.opened.deck.slides[0]!.elements[0]!
    expect(current.type === 'text' && current.text?.paragraphs[0]?.runs[0]?.text).toBe('after')
    expect(current.transform).toMatchObject({
      offset: { x: 127_000, y: 254_000, cx: 1_270_000, cy: 508_000 },
      rot: 900_000,
    })
  })

  it('allocates a durable creation id before inserting and reports it', async () => {
    const { session, slide } = await fixture()
    const tx: PresentationTransaction = {
      transactionId: 'desktop-insert',
      expectedDeckRevision: await fingerprintPresentation(session.opened),
      mode: 'atomic',
      operations: [
        {
          kind: 'add_text_box',
          clientId: 'insert',
          slideId: slide.durableId,
          text: 'new',
          geometry: { x: 1, y: 2, width: 3, height: 4 },
        },
      ],
    }
    const receipt = await new PresentationTransactionExecutor(
      new DesktopPresentationHost(session),
      { verifyDelayMs: 0 },
    ).execute(tx)
    expect(receipt.status).toBe('applied')
    if (receipt.status !== 'applied') return
    expect(receipt.createdIds).toHaveLength(1)
    expect(session.opened.deck.slides[0]!.elements.at(-1)?.creationId).toBe(receipt.createdIds?.[0])
  })

  it('enrolls a generated durable id inside the same atomic text transaction and survives reopen', async () => {
    const { session, slide, element } = await fixture()
    delete element.creationId
    const proposedId = '{11111111-2222-3333-4444-555555555555}'
    const transaction: PresentationTransaction = {
      transactionId: 'desktop-enroll-text',
      expectedDeckRevision: await fingerprintPresentation(session.opened),
      mode: 'atomic',
      operations: [
        {
          kind: 'set_text',
          clientId: 'text',
          target: {
            slideId: slide.durableId,
            elementId: proposedId,
            expectedType: 'text',
            expectedFingerprint: await fingerprintSlideElement(session.opened, slide, element),
          },
          paragraphs: [{ runs: [{ text: 'durable', bold: true }] }],
        },
      ],
    }
    const executor = new PresentationTransactionExecutor(
      new DesktopPresentationHost(session, (id) =>
        id === proposedId
          ? { slideId: slide.durableId, sourceId: element.id, elementId: proposedId }
          : undefined,
      ),
      { verifyDelayMs: 0 },
    )
    const first = await executor.execute(transaction)
    const retry = await executor.execute(transaction)
    expect(first).toMatchObject({ status: 'applied', operationCount: 1 })
    expect(retry).toEqual(first)
    expect(session.undoStack).toHaveLength(1)
    expect(session.opened.deck.slides[0]!.elements[0]!.creationId).toBe(proposedId)

    const reopened = await openPptx(await savePptx(session.opened))
    expect(reopened.deck.slides[0]!.elements[0]!.creationId).toBe(proposedId)
    expect(
      reopened.deck.slides[0]!.elements[0]!.type === 'text' &&
        reopened.deck.slides[0]!.elements[0]!.text?.paragraphs[0]?.runs[0]?.text,
    ).toBe('durable')
  })

  it('does not enroll a proposed id when the text target fingerprint is stale', async () => {
    const { session, slide, element } = await fixture()
    delete element.creationId
    const proposedId = '{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}'
    const receipt = await new PresentationTransactionExecutor(
      new DesktopPresentationHost(session, () => ({
        slideId: slide.durableId,
        sourceId: element.id,
        elementId: proposedId,
      })),
      { verifyDelayMs: 0 },
    ).execute({
      transactionId: 'desktop-stale-enroll',
      expectedDeckRevision: await fingerprintPresentation(session.opened),
      mode: 'atomic',
      operations: [
        {
          kind: 'set_text',
          clientId: 'text',
          target: {
            slideId: slide.durableId,
            elementId: proposedId,
            expectedType: 'text',
            expectedFingerprint: `sha256:${'f'.repeat(64)}`,
          },
          text: 'must not land',
        },
      ],
    })
    expect(receipt).toMatchObject({ status: 'conflict', code: 'target_stale' })
    expect(session.opened.deck.slides[0]!.elements[0]!.creationId).toBeUndefined()
    expect(session.undoStack).toHaveLength(0)
  })

  it('classifies a canceling operation sequence as a net no-op', async () => {
    const { session, slide, element } = await fixture()
    const elementFingerprint = await fingerprintSlideElement(session.opened, slide, element)
    const target = {
      slideId: slide.durableId,
      elementId: element.creationId!,
      expectedType: 'text' as const,
      expectedFingerprint: elementFingerprint,
    }
    const tx: PresentationTransaction = {
      transactionId: 'desktop-net-noop',
      expectedDeckRevision: await fingerprintPresentation(session.opened),
      mode: 'atomic',
      operations: [
        { kind: 'set_text', clientId: 'temporary', target, text: 'temporary' },
        { kind: 'set_text', clientId: 'restore', target, text: 'before' },
      ],
    }
    const receipt = await new PresentationTransactionExecutor(
      new DesktopPresentationHost(session),
    ).execute(tx)
    expect(receipt).toMatchObject({ status: 'unchanged', code: 'operation_noop' })
    expect(session.undoStack).toHaveLength(0)
    expect(
      session.opened.deck.slides[0]!.elements[0]!.type === 'text' &&
        session.opened.deck.slides[0]!.elements[0]!.text?.paragraphs[0]?.runs[0]?.text,
    ).toBe('before')
  })

  it('reports an exact rich-text replacement as a no-op without history', async () => {
    const { session, slide, element } = await fixture()
    const receipt = await new PresentationTransactionExecutor(
      new DesktopPresentationHost(session),
      { verifyDelayMs: 0 },
    ).execute({
      transactionId: 'desktop-rich-noop',
      expectedDeckRevision: await fingerprintPresentation(session.opened),
      mode: 'atomic',
      operations: [
        {
          kind: 'set_text',
          clientId: 'text',
          target: {
            slideId: slide.durableId,
            elementId: element.creationId!,
            expectedType: 'text',
            expectedFingerprint: await fingerprintSlideElement(session.opened, slide, element),
          },
          paragraphs: [{ runs: [{ text: 'before', bold: true }] }],
        },
      ],
    })
    expect(receipt).toMatchObject({ status: 'unchanged', code: 'operation_noop' })
    expect(session.undoStack).toHaveLength(0)
  })

  it('rejects picture fill because the pptx save path cannot persist it', async () => {
    const { session, slide } = await fixture()
    const picture = addPicture(session.opened, slide, {
      bytes: new Uint8Array([1, 2, 3]),
      ext: 'png',
      offset: { x: 0, y: 0, cx: 12_700, cy: 12_700 },
    })!
    const tx: PresentationTransaction = {
      transactionId: 'desktop-picture-fill',
      expectedDeckRevision: await fingerprintPresentation(session.opened),
      mode: 'atomic',
      operations: [
        {
          kind: 'set_fill',
          clientId: 'fill',
          target: {
            slideId: slide.durableId,
            elementId: picture.creationId!,
            expectedType: 'image',
            expectedFingerprint: await fingerprintSlideElement(session.opened, slide, picture),
          },
          fill: { kind: 'solid', color: '#000000' },
        },
      ],
    }
    const receipt = await new PresentationTransactionExecutor(
      new DesktopPresentationHost(session),
    ).execute(tx)
    expect(receipt).toMatchObject({ status: 'conflict', code: 'target_stale' })
    expect(session.undoStack).toHaveLength(0)
  })

  it('refuses automatic restore without a host-wide atomic CAS', async () => {
    const { session } = await fixture()
    const host = new DesktopPresentationHost(session)
    const snapshot = await host.captureSnapshot()
    expect(await host.restoreIfCurrent(await host.readRevision(), snapshot)).toBe('unsupported')
  })

  it('refuses history publication after an ordinary mutation generation change', async () => {
    const { session } = await fixture()
    const host = new DesktopPresentationHost(session)
    const snapshot = await host.captureSnapshot()
    session.mutationGeneration = (session.mutationGeneration ?? 0) + 1
    expect(await host.publishHistory(snapshot)).toBe(false)
    expect(session.undoStack).toHaveLength(0)
  })

  it('does not apply after the authoritative snapshot generation changed', async () => {
    const { session, slide, element } = await fixture()
    const host = new DesktopPresentationHost(session)
    const snapshot = await host.captureSnapshot()
    const operation = {
      kind: 'set_text' as const,
      clientId: 'lease-race',
      target: {
        slideId: slide.durableId,
        elementId: element.creationId!,
        expectedType: 'text' as const,
        expectedFingerprint: await fingerprintSlideElement(session.opened, slide, element),
      },
      text: 'transaction',
    }
    const plan = await host.plan(snapshot, [operation], () => 'unused')
    expect(plan.status).toBe('planned')
    if (plan.status !== 'planned') return

    session.mutationGeneration = (session.mutationGeneration ?? 0) + 1
    const liveElement = session.opened.deck.slides[0]!.elements[0]!
    if (liveElement.type === 'text') {
      liveElement.text!.paragraphs = [{ runs: [{ text: 'ordinary edit' }] }]
    }
    await expect(host.apply(plan.operations[0]!)).rejects.toThrow(/lease changed/)
    expect(
      liveElement.type === 'text' ? liveElement.text?.paragraphs[0]?.runs[0]?.text : undefined,
    ).toBe('ordinary edit')
  })

  it('maps dash_dot to the valid DrawingML dashDot preset across save and reopen', async () => {
    const { session, slide, element } = await fixture()
    const tx: PresentationTransaction = {
      transactionId: 'desktop-dash-dot',
      expectedDeckRevision: await fingerprintPresentation(session.opened),
      mode: 'atomic',
      operations: [
        {
          kind: 'set_stroke',
          clientId: 'stroke',
          target: {
            slideId: slide.durableId,
            elementId: element.creationId!,
            expectedType: 'text',
            expectedFingerprint: await fingerprintSlideElement(session.opened, slide, element),
          },
          stroke: { color: '#112233', width: 1, dash: 'dash_dot' },
        },
      ],
    }
    expect(
      await new PresentationTransactionExecutor(new DesktopPresentationHost(session)).execute(tx),
    ).toMatchObject({ status: 'applied' })
    const reopened = await openPptx(await savePptx(session.opened))
    const reopenedElement = reopened.deck.slides[0]!.elements.find(
      (candidate) => candidate.creationId === element.creationId,
    )
    expect(
      reopenedElement &&
        (reopenedElement.type === 'text' ||
          reopenedElement.type === 'shape' ||
          reopenedElement.type === 'picture')
        ? reopenedElement.stroke?.dash
        : undefined,
    ).toBe('dashDot')
  })

  it('fails closed when a durable element target is missing', async () => {
    const { session, slide, element } = await fixture()
    const missingId = `{${element.creationId![1] === '0' ? '1' : '0'}${element.creationId!.slice(2)}`
    // A target absent from the authoritative top-level set cannot enter the write phase.
    const tx: PresentationTransaction = {
      transactionId: 'desktop-missing',
      expectedDeckRevision: await fingerprintPresentation(session.opened),
      mode: 'atomic',
      operations: [
        {
          kind: 'delete_element',
          clientId: 'delete',
          target: {
            slideId: slide.durableId,
            elementId: missingId,
            expectedType: 'text',
            expectedFingerprint: await fingerprintSlideElement(session.opened, slide, element),
          },
        },
      ],
    }
    const receipt = await new PresentationTransactionExecutor(
      new DesktopPresentationHost(session),
    ).execute(tx)
    expect(receipt).toMatchObject({ status: 'conflict', code: 'target_missing' })
    expect(session.undoStack).toHaveLength(0)
  })
})
