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
import type { GroupElement, SlideElement, TableElement } from '@wiswork/pptx-engine'
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
  it('patches direct group-child geometry through the group XML and survives save/reopen', async () => {
    const { session, slide, element } = await fixture()
    const childXml =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="10" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>before</a:t></a:r></a:p></p:txBody></p:sp>'
    element.nvId = '2'
    element.anchor.originalXml = childXml
    const groupXml =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="2000000" cy="3000000"/><a:chOff x="500000" y="600000"/><a:chExt cx="1000000" cy="1000000"/></a:xfrm></p:grpSpPr>' +
      childXml +
      '</p:grpSp>'
    const group: GroupElement = {
      id: 'group',
      type: 'group',
      creationId: '{AAAAAAAA-1111-2222-3333-BBBBBBBBBBBB}',
      anchor: { spIndex: 0, originalXml: groupXml, range: [0, groupXml.length] },
      transform: {
        offset: { x: 100_000, y: 200_000, cx: 2_000_000, cy: 3_000_000 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      childOffset: { x: 500_000, y: 600_000, cx: 1_000_000, cy: 1_000_000 },
      children: [element],
    }
    slide.elements = [group]
    slide.structureDirty = true
    const transaction: PresentationTransaction = {
      transactionId: 'desktop-group-geometry',
      expectedDeckRevision: await fingerprintPresentation(session.opened),
      mode: 'atomic',
      operations: [
        {
          kind: 'set_geometry',
          clientId: 'geometry',
          target: {
            slideId: slide.durableId,
            elementId: element.creationId!,
            expectedType: 'text',
            expectedFingerprint: await fingerprintSlideElement(session.opened, slide, element),
          },
          geometry: { x: 30, y: 50, width: 40, height: 60, rotation: 15 },
        },
      ],
    }
    const receipt = await new PresentationTransactionExecutor(
      new DesktopPresentationHost(session),
      { verifyDelayMs: 0 },
    ).execute(transaction)
    expect(receipt).toMatchObject({ status: 'applied' })
    expect(group.anchor.originalXml).toContain(
      '<a:off x="640500" y="745000"/><a:ext cx="254000" cy="254000"/>',
    )
    expect(group.anchor.originalXml).toContain('rot="900000"')

    const reopened = await openPptx(await savePptx(session.opened))
    const reopenedGroup = reopened.deck.slides[0]!.elements[0]
    expect(reopenedGroup?.type).toBe('group')
    expect(reopenedGroup?.type === 'group' && reopenedGroup.children[0]?.transform.offset).toEqual({
      x: 640_500,
      y: 745_000,
      cx: 254_000,
      cy: 254_000,
    })
  })

  it('resizes table grid geometry with the canonical frame', async () => {
    const { session, slide } = await fixture()
    const tableXml =
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="0" y="0"/><a:ext cx="200" cy="100"/></p:xfrm>' +
      '<a:graphic><a:graphicData><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="100"/><a:gridCol w="100"/></a:tblGrid>' +
      '<a:tr h="100"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>' +
      '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
    const table: TableElement = {
      id: 'table',
      type: 'table',
      creationId: '{CCCCCCCC-1111-2222-3333-DDDDDDDDDDDD}',
      nvId: '7',
      anchor: { spIndex: 0, originalXml: tableXml, range: [0, tableXml.length] },
      transform: {
        offset: { x: 0, y: 0, cx: 200, cy: 100 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      colWidths: [100, 100],
      rowHeights: [100],
      rows: [[{}, {}]],
    }
    slide.elements = [table]
    const receipt = await new PresentationTransactionExecutor(
      new DesktopPresentationHost(session),
      { verifyDelayMs: 0 },
    ).execute({
      transactionId: 'desktop-table-geometry',
      expectedDeckRevision: await fingerprintPresentation(session.opened),
      mode: 'atomic',
      operations: [
        {
          kind: 'set_geometry',
          clientId: 'table-size',
          target: {
            slideId: slide.durableId,
            elementId: table.creationId!,
            expectedType: 'table',
            expectedFingerprint: await fingerprintSlideElement(session.opened, slide, table),
          },
          geometry: { x: 1, y: 2, width: 30, height: 20 },
        },
      ],
    })
    expect(receipt).toMatchObject({ status: 'applied' })
    expect(table.colWidths.reduce((sum, value) => sum + value, 0)).toBe(381_000)
    expect(table.rowHeights.reduce((sum, value) => sum + value, 0)).toBe(254_000)
    expect(table.transform.offset).toEqual({ x: 12_700, y: 25_400, cx: 381_000, cy: 254_000 })
  })

  it('updates attached connectors once from the final multi-operation geometry', async () => {
    const { session, slide, element: first } = await fixture()
    first.nvId = '2'
    const second = addElement(slide, {
      kind: 'textbox',
      offset: { x: 1_270_000, y: 127_000, cx: 127_000, cy: 127_000 },
      paragraphs: [{ runs: [{ text: 'second' }] }],
    })
    second.nvId = '3'
    const connector = {
      id: 'connector',
      type: 'shape',
      nvId: '4',
      anchor: { spIndex: 2, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 1, cy: 1 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      connection: { start: { id: 2, idx: 3 }, end: { id: 3, idx: 1 } },
    } as SlideElement
    slide.elements.push(connector)
    const fingerprint = await fingerprintSlideElement(session.opened, slide, first)
    const target = {
      slideId: slide.durableId,
      elementId: first.creationId!,
      expectedType: 'text' as const,
      expectedFingerprint: fingerprint,
    }
    const receipt = await new PresentationTransactionExecutor(
      new DesktopPresentationHost(session),
      { verifyDelayMs: 0 },
    ).execute({
      transactionId: 'desktop-dependent-connectors',
      expectedDeckRevision: await fingerprintPresentation(session.opened),
      mode: 'atomic',
      operations: [
        {
          kind: 'set_geometry',
          clientId: 'move-1',
          target,
          geometry: { x: 10, y: 10, width: 10, height: 10 },
        },
        {
          kind: 'set_geometry',
          clientId: 'move-2',
          target,
          geometry: { x: 20, y: 10, width: 20, height: 10 },
        },
      ],
    })
    expect(receipt).toMatchObject({ status: 'applied', operationCount: 2 })
    expect(connector.transform.offset).toEqual({
      x: 508_000,
      y: 190_500,
      cx: 762_000,
      cy: 0,
    })
  })

  it('preserves connector endpoints across geometry-delete-geometry transaction boundaries', async () => {
    const { session, slide, element: first } = await fixture()
    first.nvId = '2'
    const second = addElement(slide, {
      kind: 'textbox',
      offset: { x: 1_270_000, y: 127_000, cx: 127_000, cy: 127_000 },
      paragraphs: [{ runs: [{ text: 'second' }] }],
    })
    second.nvId = '3'
    const connector = {
      id: 'connector-sequential',
      type: 'shape',
      nvId: '4',
      anchor: { spIndex: 2, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 1, cy: 1 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      connection: { start: { id: 2, idx: 3 }, end: { id: 3, idx: 1 } },
    } as SlideElement
    slide.elements.push(connector)
    const firstTarget = {
      slideId: slide.durableId,
      elementId: first.creationId!,
      expectedType: 'text' as const,
      expectedFingerprint: await fingerprintSlideElement(session.opened, slide, first),
    }
    const secondTarget = {
      slideId: slide.durableId,
      elementId: second.creationId!,
      expectedType: 'text' as const,
      expectedFingerprint: await fingerprintSlideElement(session.opened, slide, second),
    }
    const receipt = await new PresentationTransactionExecutor(
      new DesktopPresentationHost(session),
      { verifyDelayMs: 0 },
    ).execute({
      transactionId: 'desktop-connector-delete-boundary',
      expectedDeckRevision: await fingerprintPresentation(session.opened),
      mode: 'atomic',
      operations: [
        {
          kind: 'set_geometry',
          clientId: 'move-a',
          target: firstTarget,
          geometry: { x: 20, y: 10, width: 20, height: 10 },
        },
        { kind: 'delete_element', clientId: 'delete-a', target: firstTarget },
        {
          kind: 'set_geometry',
          clientId: 'move-b',
          target: secondTarget,
          geometry: { x: 120, y: 10, width: 10, height: 10 },
        },
      ],
    })
    expect(receipt).toMatchObject({ status: 'applied', operationCount: 3 })
    expect(slide.elements).not.toContain(first)
    expect(connector.transform.offset).toEqual({
      x: 508_000,
      y: 190_500,
      cx: 1_016_000,
      cy: 0,
    })
  })

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
