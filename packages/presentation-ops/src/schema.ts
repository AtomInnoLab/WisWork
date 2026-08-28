import type {
  PresentationElementType,
  PresentationElementTarget,
  PresentationFill,
  PresentationGeometry,
  PresentationOperation,
  PresentationReceipt,
  PresentationQualityCode,
  PresentationQualityFinding,
  PresentationQualityReceipt,
  PresentationStroke,
  PresentationTarget,
  PresentationTextParagraph,
  PresentationTransaction,
} from './types'
import { readStrictArray } from './strict-array'

export const PRESENTATION_OPS_LIMITS = Object.freeze({
  maxOperations: 50,
  maxTextLength: 12_000,
  maxIdentifierLength: 128,
  maxReceiptIds: 50,
  maxCoordinateMagnitude: 1_000_000,
  maxStrokeWidth: 1_000,
  maxTextParagraphs: 1_000,
  maxTextRuns: 4_000,
  maxQualityFindings: 50,
  maxQualityEvidenceFields: 8,
})

const qualityCodes: readonly PresentationQualityCode[] = [
  'element_off_slide',
  'text_overflow_horizontal',
  'text_overflow_vertical',
  'element_collision',
  'tiny_text',
  'empty_placeholder',
  'low_contrast',
  'visual_quality',
]

const parseQualityFinding = (value: unknown, index: number): PresentationQualityFinding => {
  const name = `findings[${index}]`
  const record = object(value, name)
  exactKeys(
    record,
    ['code', 'severity', 'slideId', 'elementId', 'relatedElementId', 'evidence'],
    name,
  )
  if (!qualityCodes.includes(record.code as PresentationQualityCode))
    fail(`${name}.code is unknown`)
  if (!['warning', 'important', 'critical'].includes(record.severity as string))
    fail(`${name}.severity is unknown`)
  const evidenceRecord = object(record.evidence, `${name}.evidence`)
  if (Object.keys(evidenceRecord).length > PRESENTATION_OPS_LIMITS.maxQualityEvidenceFields)
    fail(`${name}.evidence has too many fields`)
  const evidence: Record<string, number> = {}
  for (const [key, evidenceValue] of Object.entries(evidenceRecord)) {
    if (!identifierPattern.test(key) || unsafeKeys.has(key)) fail(`${name}.evidence key is unsafe`)
    evidence[key] = finiteNumber(evidenceValue, `${name}.evidence.${key}`)
  }
  const elementId = optional(record, 'elementId', (item) =>
    elementIdentifier(item, `${name}.elementId`),
  )
  const relatedElementId = optional(record, 'relatedElementId', (item) =>
    elementIdentifier(item, `${name}.relatedElementId`),
  )
  return {
    code: record.code as PresentationQualityCode,
    severity: record.severity as PresentationQualityFinding['severity'],
    slideId: slideIdentifier(record.slideId, `${name}.slideId`),
    ...(elementId === undefined ? {} : { elementId }),
    ...(relatedElementId === undefined ? {} : { relatedElementId }),
    evidence,
  }
}

export const parsePresentationQualityReceipt = (value: unknown): PresentationQualityReceipt => {
  const record = object(value, 'quality receipt')
  if (record.status === 'available') {
    exactKeys(
      record,
      ['qualityRunId', 'source', 'status', 'findings', 'truncated'],
      'quality receipt',
    )
    if (!['deterministic', 'visual'].includes(record.source as string))
      fail('quality receipt.source is unknown')
    const values = readStrictArray(record.findings, 'findings', {
      minLength: 0,
      maxLength: PRESENTATION_OPS_LIMITS.maxQualityFindings,
    })
    if (values.length > PRESENTATION_OPS_LIMITS.maxQualityFindings)
      fail('quality receipt has too many findings')
    if (typeof record.truncated !== 'boolean') fail('quality receipt.truncated must be boolean')
    return {
      qualityRunId: identifier(record.qualityRunId, 'qualityRunId'),
      source: record.source as 'deterministic' | 'visual',
      status: 'available',
      findings: values.map(parseQualityFinding),
      truncated: record.truncated,
    }
  }
  exactKeys(record, ['qualityRunId', 'source', 'status', 'code'], 'quality receipt')
  if (record.source !== 'visual' || !['unavailable', 'cancelled'].includes(record.status as string))
    fail('quality receipt status is unknown')
  const codes = [
    'screenshot_unavailable',
    'transport_unavailable',
    'cancelled',
    'stale_session',
  ] as const
  if (!codes.includes(record.code as (typeof codes)[number]))
    fail('quality receipt.code is unknown')
  return {
    qualityRunId: identifier(record.qualityRunId, 'qualityRunId'),
    source: 'visual',
    status: record.status as 'unavailable' | 'cancelled',
    code: record.code as (typeof codes)[number],
  }
}

const unsafeKeys = new Set(['__proto__', 'prototype', 'constructor'])
const identifierPattern = /^[A-Za-z0-9][A-Za-z0-9._:-]*$/
const slidePartIdentifierPattern = /^ppt\/slides\/slide[1-9][0-9]*\.xml$/
const drawingCreationIdPattern = /^\{[0-9A-Fa-f]{8}(?:-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}\}$/
const fingerprintPattern = /^sha256:[0-9a-f]{64}$/
const colorPattern = /^#[0-9A-Fa-f]{6}$/

function fail(message: string): never {
  throw new TypeError(`Invalid presentation contract: ${message}`)
}

const object = (value: unknown, name: string): Record<string, unknown> => {
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    fail(`${name} must be an object`)
  const record = value as Record<string, unknown>
  const prototype = Object.getPrototypeOf(record)
  if (prototype !== Object.prototype && prototype !== null) fail(`${name} must be a plain object`)
  for (const key of Reflect.ownKeys(record)) {
    if (typeof key !== 'string') fail(`${name} contains an unknown symbol field`)
    if (unsafeKeys.has(key)) fail(`${name} contains unsafe key ${key}`)
    const descriptor = Object.getOwnPropertyDescriptor(record, key)
    if (descriptor === undefined || descriptor.get !== undefined || descriptor.set !== undefined) {
      fail(`${name} contains an accessor field`)
    }
    if (!descriptor.enumerable) fail(`${name} contains a hidden field`)
  }
  return record
}

const exactKeys = (
  record: Record<string, unknown>,
  allowed: readonly string[],
  name: string,
): void => {
  const allowedSet = new Set(allowed)
  for (const key of Object.keys(record)) {
    if (!allowedSet.has(key)) fail(`${name} contains unknown field ${key}`)
  }
}

const requiredString = (
  value: unknown,
  name: string,
  maxLength: number = PRESENTATION_OPS_LIMITS.maxTextLength,
): string => {
  if (typeof value !== 'string') fail(`${name} must be a bounded string`)
  if (value.length > maxLength) fail(`${name} must be a bounded string`)
  return value
}

const identifier = (value: unknown, name: string): string => {
  const result = requiredString(value, name, PRESENTATION_OPS_LIMITS.maxIdentifierLength)
  if (!identifierPattern.test(result) || unsafeKeys.has(result)) fail(`${name} is unsafe`)
  return result
}

const slideIdentifier = (value: unknown, name: string): string => {
  const result = requiredString(value, name, PRESENTATION_OPS_LIMITS.maxIdentifierLength)
  if (
    (!identifierPattern.test(result) && !slidePartIdentifierPattern.test(result)) ||
    unsafeKeys.has(result)
  )
    fail(`${name} is unsafe`)
  return result
}

const elementIdentifier = (value: unknown, name: string): string => {
  const result = requiredString(value, name, PRESENTATION_OPS_LIMITS.maxIdentifierLength)
  if (
    (!identifierPattern.test(result) && !drawingCreationIdPattern.test(result)) ||
    unsafeKeys.has(result)
  )
    fail(`${name} is unsafe`)
  return result
}

const fingerprint = (value: unknown, name: string): string => {
  if (typeof value !== 'string') fail(`${name} must be a SHA-256 fingerprint`)
  if (!fingerprintPattern.test(value)) fail(`${name} must be a SHA-256 fingerprint`)
  return value
}

const finiteNumber = (value: unknown, name: string): number => {
  if (typeof value !== 'number') fail(`${name} must be finite`)
  if (!Number.isFinite(value)) fail(`${name} must be finite`)
  return value
}

const boundedInteger = (value: unknown, name: string, maximum: number): number => {
  const result = finiteNumber(value, name)
  if (!Number.isInteger(result) || result < 0 || result > maximum) fail(`${name} is out of bounds`)
  return result
}

const optional = <T>(
  record: Record<string, unknown>,
  key: string,
  parse: (value: unknown) => T,
): T | undefined => (Object.hasOwn(record, key) ? parse(record[key]) : undefined)

const parseTarget = (value: unknown, requireElement: boolean): PresentationTarget => {
  const record = object(value, 'target')
  exactKeys(record, ['slideId', 'elementId', 'expectedType', 'expectedFingerprint'], 'target')
  const slideId = slideIdentifier(record.slideId, 'target.slideId')
  const elementId = optional(record, 'elementId', (item) =>
    elementIdentifier(item, 'target.elementId'),
  )
  const expectedType = optional(record, 'expectedType', (item) => {
    const types: readonly PresentationElementType[] = [
      'text',
      'shape',
      'image',
      'table',
      'chart',
      'group',
    ]
    if (!types.includes(item as PresentationElementType)) fail('target.expectedType is unknown')
    return item as PresentationElementType
  })
  const expectedFingerprint = optional(record, 'expectedFingerprint', (item) =>
    fingerprint(item, 'target.expectedFingerprint'),
  )
  if (requireElement && elementId === undefined) fail('target.elementId is required')
  if (!requireElement && (elementId !== undefined || expectedType !== undefined))
    fail('slide target cannot identify an element')
  return {
    slideId,
    ...(elementId === undefined ? {} : { elementId }),
    ...(expectedType === undefined ? {} : { expectedType }),
    ...(expectedFingerprint === undefined ? {} : { expectedFingerprint }),
  }
}

const parseElementTarget = (value: unknown): PresentationElementTarget => {
  const record = object(value, 'target')
  if (Object.hasOwn(record, 'createdByClientId')) {
    exactKeys(record, ['createdByClientId'], 'generated target')
    return { createdByClientId: identifier(record.createdByClientId, 'target.createdByClientId') }
  }
  return parseTarget(value, true)
}

export const parsePresentationTarget = (value: unknown): PresentationTarget => {
  const record = object(value, 'target')
  return parseTarget(value, Object.hasOwn(record, 'elementId'))
}

const parseGeometry = (value: unknown): PresentationGeometry => {
  const record = object(value, 'geometry')
  exactKeys(record, ['x', 'y', 'width', 'height', 'rotation'], 'geometry')
  const x = finiteNumber(record.x, 'geometry.x')
  const y = finiteNumber(record.y, 'geometry.y')
  const width = finiteNumber(record.width, 'geometry.width')
  const height = finiteNumber(record.height, 'geometry.height')
  const rotation = optional(record, 'rotation', (item) => finiteNumber(item, 'geometry.rotation'))
  const magnitude = PRESENTATION_OPS_LIMITS.maxCoordinateMagnitude
  if (
    Math.abs(x) > magnitude ||
    Math.abs(y) > magnitude ||
    width <= 0 ||
    width > magnitude ||
    height <= 0 ||
    height > magnitude
  ) {
    fail('geometry is out of bounds')
  }
  if (rotation !== undefined && Math.abs(rotation) > 360_000)
    fail('geometry.rotation is out of bounds')
  return { x, y, width, height, ...(rotation === undefined ? {} : { rotation }) }
}

const parseFill = (value: unknown): PresentationFill => {
  const record = object(value, 'fill')
  if (record.kind === 'none') {
    exactKeys(record, ['kind'], 'fill')
    return { kind: 'none' }
  }
  if (record.kind !== 'solid') fail('fill.kind is unknown')
  exactKeys(record, ['kind', 'color', 'transparency'], 'fill')
  if (typeof record.color !== 'string') fail('fill.color is invalid')
  if (!colorPattern.test(record.color)) fail('fill.color is invalid')
  const transparency = optional(record, 'transparency', (item) =>
    finiteNumber(item, 'fill.transparency'),
  )
  if (transparency !== undefined && (transparency < 0 || transparency > 1))
    fail('fill.transparency is out of bounds')
  return {
    kind: 'solid',
    color: record.color.toUpperCase(),
    ...(transparency === undefined ? {} : { transparency }),
  }
}

const parseStroke = (value: unknown): PresentationStroke => {
  const record = object(value, 'stroke')
  exactKeys(record, ['color', 'width', 'dash'], 'stroke')
  if (typeof record.color !== 'string') fail('stroke.color is invalid')
  if (!colorPattern.test(record.color)) fail('stroke.color is invalid')
  const width = finiteNumber(record.width, 'stroke.width')
  if (width < 0 || width > PRESENTATION_OPS_LIMITS.maxStrokeWidth)
    fail('stroke.width is out of bounds')
  const dash = optional(record, 'dash', (item) => {
    const values = ['solid', 'dash', 'dot', 'dash_dot'] as const
    if (!values.includes(item as (typeof values)[number])) fail('stroke.dash is unknown')
    return item as (typeof values)[number]
  })
  return { color: record.color.toUpperCase(), width, ...(dash === undefined ? {} : { dash }) }
}

const parseTextParagraphs = (value: unknown): PresentationTextParagraph[] => {
  const values = readStrictArray(value, 'paragraphs', {
    minLength: 1,
    maxLength: PRESENTATION_OPS_LIMITS.maxTextParagraphs,
  })
  let runCount = 0
  let textLength = 0
  return values.map((paragraphValue, paragraphIndex) => {
    const paragraph = object(paragraphValue, `paragraphs[${paragraphIndex}]`)
    exactKeys(paragraph, ['runs', 'align'], `paragraphs[${paragraphIndex}]`)
    const runs = readStrictArray(paragraph.runs, `paragraphs[${paragraphIndex}].runs`, {
      minLength: 1,
      maxLength: PRESENTATION_OPS_LIMITS.maxTextRuns,
    }).map((runValue, runIndex) => {
      runCount += 1
      if (runCount > PRESENTATION_OPS_LIMITS.maxTextRuns) fail('paragraphs contain too many runs')
      const run = object(runValue, `paragraphs[${paragraphIndex}].runs[${runIndex}]`)
      exactKeys(
        run,
        ['text', 'bold', 'italic', 'underline', 'fontSize', 'fontFamily', 'color'],
        `paragraphs[${paragraphIndex}].runs[${runIndex}]`,
      )
      const text = requiredString(run.text, 'run.text')
      textLength += text.length
      if (textLength > PRESENTATION_OPS_LIMITS.maxTextLength) fail('paragraph text is too long')
      const boolean = (key: 'bold' | 'italic' | 'underline') =>
        optional(run, key, (item) => {
          if (typeof item !== 'boolean') fail(`run.${key} must be boolean`)
          return item
        })
      const fontSize = optional(run, 'fontSize', (item) => finiteNumber(item, 'run.fontSize'))
      if (fontSize !== undefined && (fontSize <= 0 || fontSize > 1_000))
        fail('run.fontSize is out of bounds')
      const fontFamily = optional(run, 'fontFamily', (item) =>
        requiredString(item, 'run.fontFamily', 256),
      )
      const color = optional(run, 'color', (item) => {
        if (typeof item !== 'string' || !colorPattern.test(item)) fail('run.color is invalid')
        return item.toUpperCase()
      })
      const bold = boolean('bold')
      const italic = boolean('italic')
      const underline = boolean('underline')
      return {
        text,
        ...(bold === undefined ? {} : { bold }),
        ...(italic === undefined ? {} : { italic }),
        ...(underline === undefined ? {} : { underline }),
        ...(fontSize === undefined ? {} : { fontSize }),
        ...(fontFamily === undefined ? {} : { fontFamily }),
        ...(color === undefined ? {} : { color }),
      }
    })
    const align = optional(paragraph, 'align', (item) => {
      if (item !== 'left' && item !== 'center' && item !== 'right')
        fail('paragraph.align is unknown')
      return item
    })
    return { runs, ...(align === undefined ? {} : { align }) }
  })
}

export const parsePresentationOperation = (value: unknown): PresentationOperation => {
  const record = object(value, 'operation')
  const kind = record.kind
  const clientId = identifier(record.clientId, 'operation.clientId')
  switch (kind) {
    case 'set_text':
      exactKeys(record, ['kind', 'clientId', 'target', 'text', 'paragraphs'], 'set_text')
      if (Object.hasOwn(record, 'text') === Object.hasOwn(record, 'paragraphs'))
        fail('set_text requires exactly one of text or paragraphs')
      return Object.hasOwn(record, 'text')
        ? {
            kind,
            clientId,
            target: parseElementTarget(record.target),
            text: requiredString(record.text, 'text'),
          }
        : {
            kind,
            clientId,
            target: parseElementTarget(record.target),
            paragraphs: parseTextParagraphs(record.paragraphs),
          }
    case 'set_geometry':
      exactKeys(record, ['kind', 'clientId', 'target', 'geometry'], 'set_geometry')
      return {
        kind,
        clientId,
        target: parseElementTarget(record.target),
        geometry: parseGeometry(record.geometry),
      }
    case 'set_fill':
      exactKeys(record, ['kind', 'clientId', 'target', 'fill'], 'set_fill')
      return {
        kind,
        clientId,
        target: parseElementTarget(record.target),
        fill: parseFill(record.fill),
      }
    case 'set_stroke':
      exactKeys(record, ['kind', 'clientId', 'target', 'stroke'], 'set_stroke')
      return {
        kind,
        clientId,
        target: parseElementTarget(record.target),
        stroke: record.stroke === null ? null : parseStroke(record.stroke),
      }
    case 'add_text_box':
      exactKeys(record, ['kind', 'clientId', 'slideId', 'text', 'geometry'], 'add_text_box')
      return {
        kind,
        clientId,
        slideId: slideIdentifier(record.slideId, 'slideId'),
        text: requiredString(record.text, 'text'),
        geometry: parseGeometry(record.geometry),
      }
    case 'delete_element':
      exactKeys(record, ['kind', 'clientId', 'target'], 'delete_element')
      return { kind, clientId, target: parseElementTarget(record.target) }
    case 'set_speaker_notes':
      exactKeys(record, ['kind', 'clientId', 'target', 'notes'], 'set_speaker_notes')
      return {
        kind,
        clientId,
        target: parseTarget(record.target, false),
        notes: requiredString(record.notes, 'notes'),
      }
    case 'set_slide_background': {
      exactKeys(record, ['kind', 'clientId', 'target', 'color'], 'set_slide_background')
      if (typeof record.color !== 'string' || !colorPattern.test(record.color))
        fail('background.color is invalid')
      return {
        kind,
        clientId,
        target: parseTarget(record.target, false),
        color: record.color.toUpperCase(),
      }
    }
    default:
      return fail('operation.kind is unknown')
  }
}

export const parsePresentationTransaction = (value: unknown): PresentationTransaction => {
  const record = object(value, 'transaction')
  exactKeys(record, ['transactionId', 'expectedDeckRevision', 'operations', 'mode'], 'transaction')
  if (record.mode !== 'atomic') fail('transaction.mode must be atomic')
  const operationValues = readStrictArray(record.operations, 'transaction.operations', {
    minLength: 1,
    maxLength: PRESENTATION_OPS_LIMITS.maxOperations,
  })
  const operations: PresentationOperation[] = []
  for (const operation of operationValues) operations.push(parsePresentationOperation(operation))
  const clientIds = new Set<string>()
  for (const operation of operations) {
    if (clientIds.has(operation.clientId)) fail('operation.clientId must be unique')
    if ('target' in operation && 'createdByClientId' in operation.target) {
      const createdByClientId = operation.target.createdByClientId
      if (!clientIds.has(createdByClientId))
        fail('generated target must reference an earlier add_text_box')
      const creator = operations.find((candidate) => candidate.clientId === createdByClientId)
      if (creator?.kind !== 'add_text_box')
        fail('generated target must reference an earlier add_text_box')
    }
    clientIds.add(operation.clientId)
  }
  return {
    transactionId: identifier(record.transactionId, 'transactionId'),
    expectedDeckRevision: fingerprint(record.expectedDeckRevision, 'expectedDeckRevision'),
    operations,
    mode: 'atomic',
  }
}

export const parsePresentationReceipt = (value: unknown): PresentationReceipt => {
  const record = object(value, 'receipt')
  const transactionId = identifier(record.transactionId, 'receipt.transactionId')
  switch (record.status) {
    case 'applied': {
      exactKeys(
        record,
        [
          'status',
          'transactionId',
          'resultingDeckRevision',
          'operationCount',
          'createdIds',
          'createdTargets',
        ],
        'applied receipt',
      )
      const createdIds = optional(record, 'createdIds', (item) => {
        const values = readStrictArray(item, 'receipt.createdIds', {
          maxLength: PRESENTATION_OPS_LIMITS.maxReceiptIds,
        })
        const ids: string[] = []
        for (const id of values) ids.push(elementIdentifier(id, 'receipt.createdId'))
        if (new Set(ids).size !== ids.length) fail('receipt.createdIds must be unique')
        return ids
      })
      const createdTargets = optional(record, 'createdTargets', (item) => {
        const values = readStrictArray(item, 'receipt.createdTargets', {
          maxLength: PRESENTATION_OPS_LIMITS.maxReceiptIds,
        })
        const mappings = values.map((value) => {
          const mapping = object(value, 'receipt.createdTarget')
          exactKeys(mapping, ['clientId', 'elementId'], 'receipt.createdTarget')
          return {
            clientId: identifier(mapping.clientId, 'receipt.createdTarget.clientId'),
            elementId: elementIdentifier(mapping.elementId, 'receipt.createdTarget.elementId'),
          }
        })
        if (new Set(mappings.map((item) => item.clientId)).size !== mappings.length)
          fail('receipt.createdTargets clientIds must be unique')
        if (new Set(mappings.map((item) => item.elementId)).size !== mappings.length)
          fail('receipt.createdTargets elementIds must be unique')
        return mappings
      })
      const operationCount = boundedInteger(
        record.operationCount,
        'receipt.operationCount',
        PRESENTATION_OPS_LIMITS.maxOperations,
      )
      if (
        (createdIds?.length ?? 0) > operationCount ||
        (createdTargets?.length ?? 0) > operationCount
      )
        fail('receipt created targets cannot exceed operationCount')
      if (
        createdIds !== undefined &&
        createdTargets !== undefined &&
        (createdIds.length !== createdTargets.length ||
          createdIds.some((id, index) => id !== createdTargets[index]!.elementId))
      )
        fail('receipt createdIds and createdTargets must identify the same elements')
      return {
        status: 'applied',
        transactionId,
        resultingDeckRevision: fingerprint(
          record.resultingDeckRevision,
          'receipt.resultingDeckRevision',
        ),
        operationCount,
        ...(createdIds === undefined ? {} : { createdIds }),
        ...(createdTargets === undefined ? {} : { createdTargets }),
      }
    }
    case 'unchanged': {
      exactKeys(record, ['status', 'transactionId', 'code', 'operationCount'], 'unchanged receipt')
      if (record.code !== 'operation_noop' && record.code !== 'write_not_applied')
        fail('unchanged receipt code is unknown')
      return {
        status: 'unchanged',
        transactionId,
        code: record.code as 'operation_noop' | 'write_not_applied',
        operationCount: boundedInteger(
          record.operationCount,
          'receipt.operationCount',
          PRESENTATION_OPS_LIMITS.maxOperations,
        ),
      }
    }
    case 'conflict': {
      exactKeys(
        record,
        ['status', 'transactionId', 'code', 'operationIndex', 'targetId'],
        'conflict receipt',
      )
      if (
        record.code !== 'target_stale' &&
        record.code !== 'target_missing' &&
        record.code !== 'target_ambiguous'
      )
        fail('conflict receipt code is unknown')
      const operationIndex = optional(record, 'operationIndex', (item) =>
        boundedInteger(item, 'receipt.operationIndex', PRESENTATION_OPS_LIMITS.maxOperations - 1),
      )
      const targetId = optional(record, 'targetId', (item) =>
        elementIdentifier(item, 'receipt.targetId'),
      )
      return {
        status: 'conflict',
        transactionId,
        code: record.code as 'target_stale' | 'target_missing' | 'target_ambiguous',
        ...(operationIndex === undefined ? {} : { operationIndex }),
        ...(targetId === undefined ? {} : { targetId }),
      }
    }
    case 'uncertain': {
      exactKeys(record, ['status', 'transactionId', 'code', 'operationIndex'], 'uncertain receipt')
      if (record.code !== 'write_state_uncertain') fail('uncertain receipt code is unknown')
      const operationIndex = optional(record, 'operationIndex', (item) =>
        boundedInteger(item, 'receipt.operationIndex', PRESENTATION_OPS_LIMITS.maxOperations - 1),
      )
      return {
        status: 'uncertain',
        transactionId,
        code: 'write_state_uncertain',
        ...(operationIndex === undefined ? {} : { operationIndex }),
      }
    }
    default:
      return fail('receipt.status is unknown')
  }
}
