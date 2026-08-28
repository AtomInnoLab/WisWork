import { describe, expect, it } from 'vitest'

import {
  PRESENTATION_OPS_LIMITS,
  parsePresentationOperation,
  parsePresentationReceipt,
  parsePresentationTarget,
  parsePresentationTransaction,
} from '../src/index'

const fingerprint = `sha256:${'a'.repeat(64)}`
const target = {
  slideId: 'slide-1',
  elementId: 'shape-1',
  expectedType: 'text',
  expectedFingerprint: fingerprint,
}

const transaction = (operations: unknown[]) => ({
  transactionId: 'tx-1',
  expectedDeckRevision: fingerprint,
  mode: 'atomic',
  operations,
})

describe('presentation transaction parser', () => {
  it('accepts bounded OPC slide identities and DrawingML creation IDs', () => {
    const durableTarget = {
      ...target,
      slideId: 'ppt/slides/slide12.xml',
      elementId: '{01234567-89AB-CDEF-0123-456789ABCDEF}',
    }
    expect(parsePresentationTarget(durableTarget)).toEqual(durableTarget)
    expect(() =>
      parsePresentationTarget({ ...durableTarget, slideId: 'ppt/../secrets.xml' }),
    ).toThrow(/unsafe/i)
  })

  it('parses every closed operation family', () => {
    const value = transaction([
      { kind: 'set_text', clientId: 'op-1', target, text: 'Hello' },
      {
        kind: 'set_geometry',
        clientId: 'op-2',
        target,
        geometry: { x: 1, y: 2, width: 300, height: 100, rotation: 45 },
      },
      {
        kind: 'set_fill',
        clientId: 'op-3',
        target,
        fill: { kind: 'solid', color: '#12ABEF', transparency: 0.25 },
      },
      {
        kind: 'set_stroke',
        clientId: 'op-4',
        target,
        stroke: { color: '#000000', width: 1.5, dash: 'dash' },
      },
      { kind: 'set_stroke', clientId: 'op-4b', target, stroke: null },
      {
        kind: 'add_text_box',
        clientId: 'op-5',
        slideId: 'slide-1',
        text: 'New',
        geometry: { x: 10, y: 20, width: 100, height: 40 },
      },
      { kind: 'delete_element', clientId: 'op-6', target },
      {
        kind: 'set_speaker_notes',
        clientId: 'op-7',
        target: { slideId: 'slide-1', expectedFingerprint: fingerprint },
        notes: 'Private presenter notes',
      },
    ])

    expect(parsePresentationTransaction(value)).toEqual(value)
    expect(JSON.parse(JSON.stringify(parsePresentationTransaction(value)))).toEqual(value)
    expect(parsePresentationOperation(value.operations[0])).toEqual(value.operations[0])
    expect(parsePresentationTarget(target)).toEqual(target)
  })

  it('preserves bounded rich paragraphs for text replacement', () => {
    const rich = {
      kind: 'set_text',
      clientId: 'rich-text',
      target,
      paragraphs: [
        {
          runs: [
            { text: 'Hello', bold: true, fontSize: 24, color: '#12abef' },
            { text: ' world', italic: true },
          ],
          align: 'center',
        },
      ],
    }
    expect(parsePresentationOperation(rich)).toEqual({
      ...rich,
      paragraphs: [
        {
          runs: [
            { text: 'Hello', bold: true, fontSize: 24, color: '#12ABEF' },
            { text: ' world', italic: true },
          ],
          align: 'center',
        },
      ],
    })
    expect(() => parsePresentationOperation({ ...rich, text: 'ambiguous' })).toThrow(
      /field|exactly/i,
    )
  })

  it.each([
    ['unknown transaction field', { ...transaction([]), extra: true }],
    [
      'unknown operation field',
      transaction([{ kind: 'delete_element', clientId: 'op-1', target, payload: {} }]),
    ],
    ['unknown operation kind', transaction([{ kind: 'run_code', clientId: 'op-1', target }])],
    ['non-atomic mode', { ...transaction([]), mode: 'per_op' }],
    [
      'unsafe target id',
      transaction([
        { kind: 'delete_element', clientId: 'op-1', target: { ...target, slideId: '../slide' } },
      ]),
    ],
    ['invalid fingerprint', { ...transaction([]), expectedDeckRevision: 'weak:123' }],
    [
      'non-finite geometry',
      transaction([
        {
          kind: 'set_geometry',
          clientId: 'op-1',
          target,
          geometry: { x: Infinity, y: 0, width: 1, height: 1 },
        },
      ]),
    ],
    [
      'invalid dimensions',
      transaction([
        {
          kind: 'set_geometry',
          clientId: 'op-1',
          target,
          geometry: { x: 0, y: 0, width: 0, height: 1 },
        },
      ]),
    ],
    [
      'duplicate client ids',
      transaction([
        { kind: 'delete_element', clientId: 'same', target },
        { kind: 'delete_element', clientId: 'same', target },
      ]),
    ],
    [
      'notes targeting an element',
      transaction([{ kind: 'set_speaker_notes', clientId: 'op-1', target, notes: 'no' }]),
    ],
  ])('rejects %s', (_label, value) => {
    expect(() => parsePresentationTransaction(value)).toThrow()
  })

  it('rejects unsafe prototype keys even when they are own JSON properties', () => {
    const value = JSON.parse(
      `{"transactionId":"tx-1","expectedDeckRevision":"${fingerprint}","mode":"atomic","operations":[],"__proto__":{}}`,
    )
    expect(() => parsePresentationTransaction(value)).toThrow(/unsafe/i)
  })

  it('rejects symbol fields and accessor-backed inputs', () => {
    const withSymbol = transaction([]) as Record<PropertyKey, unknown>
    withSymbol[Symbol('hidden')] = true
    expect(() => parsePresentationTransaction(withSymbol)).toThrow(/unknown|plain/i)

    const withGetter = transaction([])
    Object.defineProperty(withGetter, 'hidden', { enumerable: true, get: () => true })
    expect(() => parsePresentationTransaction(withGetter)).toThrow(/accessor|plain/i)
  })

  it('enforces text and operation bounds', () => {
    expect(() => parsePresentationTransaction(transaction([]))).toThrow()
    expect(() =>
      parsePresentationTransaction(
        transaction([
          {
            kind: 'set_text',
            clientId: 'op-1',
            target,
            text: 'x'.repeat(PRESENTATION_OPS_LIMITS.maxTextLength + 1),
          },
        ]),
      ),
    ).toThrow()

    const operations = Array.from(
      { length: PRESENTATION_OPS_LIMITS.maxOperations + 1 },
      (_, index) => ({
        kind: 'delete_element',
        clientId: `op-${index}`,
        target,
      }),
    )
    expect(() => parsePresentationTransaction(transaction(operations))).toThrow()
  })

  it('strictly validates operation arrays without invoking accessors', () => {
    const validOperation = { kind: 'delete_element', clientId: 'op-1', target }
    const malformed: unknown[][] = [
      [validOperation],
      [validOperation],
      [validOperation],
      Array(1),
      [validOperation],
    ]
    Object.assign(malformed[0]!, { extra: true })
    Object.defineProperty(malformed[1]!, 'hidden', { value: true })
    malformed[2]![Symbol('hidden') as unknown as number] = true
    Object.setPrototypeOf(malformed[4]!, null)
    for (const operations of malformed) {
      expect(() => parsePresentationTransaction(transaction(operations))).toThrow(/array|field/i)
    }

    let accessed = false
    const accessor: unknown[] = []
    Object.defineProperty(accessor, '0', {
      enumerable: true,
      get: () => {
        accessed = true
        return validOperation
      },
    })
    expect(() => parsePresentationTransaction(transaction(accessor))).toThrow(/accessor/i)
    expect(accessed).toBe(false)
  })
})

describe('presentation receipt parser', () => {
  it.each([
    {
      status: 'applied',
      transactionId: 'tx-1',
      resultingDeckRevision: fingerprint,
      operationCount: 2,
      createdIds: ['shape-2'],
    },
    { status: 'unchanged', transactionId: 'tx-1', code: 'operation_noop', operationCount: 1 },
    {
      status: 'conflict',
      transactionId: 'tx-1',
      code: 'target_stale',
      operationIndex: 0,
      targetId: 'shape-1',
    },
    {
      status: 'uncertain',
      transactionId: 'tx-1',
      code: 'write_state_uncertain',
      operationIndex: 0,
    },
  ])('parses bounded $status receipts', (value) => {
    expect(parsePresentationReceipt(value)).toEqual(value)
  })

  it('rejects raw content and unknown receipt codes', () => {
    expect(() =>
      parsePresentationReceipt({
        status: 'uncertain',
        transactionId: 'tx-1',
        code: 'write_state_uncertain',
        rawContent: 'secret',
      }),
    ).toThrow()
    expect(() =>
      parsePresentationReceipt({ status: 'conflict', transactionId: 'tx-1', code: 'host_error' }),
    ).toThrow()
  })

  it('bounds receipt identifiers', () => {
    expect(() =>
      parsePresentationReceipt({
        status: 'applied',
        transactionId: 'tx-1',
        resultingDeckRevision: fingerprint,
        operationCount: 1,
        createdIds: Array.from(
          { length: PRESENTATION_OPS_LIMITS.maxReceiptIds + 1 },
          (_, index) => `shape-${index}`,
        ),
      }),
    ).toThrow()
  })

  it('strictly validates createdIds arrays without invoking accessors', () => {
    const receipt = (createdIds: unknown) => ({
      status: 'applied',
      transactionId: 'tx-1',
      resultingDeckRevision: fingerprint,
      operationCount: 1,
      createdIds,
    })
    const malformed: unknown[][] = [['shape-1'], ['shape-1'], ['shape-1'], Array(1), ['shape-1']]
    Object.assign(malformed[0]!, { extra: true })
    Object.defineProperty(malformed[1]!, 'hidden', { value: true })
    malformed[2]![Symbol('hidden') as unknown as number] = true
    Object.setPrototypeOf(malformed[4]!, null)
    for (const ids of malformed) expect(() => parsePresentationReceipt(receipt(ids))).toThrow()

    let accessed = false
    const accessor: unknown[] = []
    Object.defineProperty(accessor, '0', {
      enumerable: true,
      get: () => {
        accessed = true
        return 'shape-1'
      },
    })
    expect(() => parsePresentationReceipt(receipt(accessor))).toThrow(/accessor/i)
    expect(accessed).toBe(false)
  })
})
