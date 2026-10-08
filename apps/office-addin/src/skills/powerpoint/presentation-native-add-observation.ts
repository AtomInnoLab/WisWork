import { equivalentPowerPointText } from './powerpoint-text.js'
import { XMLParser, XMLValidator } from 'fast-xml-parser'
import type { PowerPointDeclarativeOperation } from './browser-powerpoint-adapter.js'
import { loadBoundedZip, MAX_PPTX_XML_BYTES } from './powerpoint-package.js'
import {
  inspectNativeTableCellStyles,
  inspectNativeTableStructure,
} from './presentation-structure-comparison.js'

export type NativeAddOperation = Extract<
  PowerPointDeclarativeOperation,
  { op: 'add_text_box' | 'add_geometric_shape' | 'add_native_table' }
>
export interface NativeAddObservation {
  status: 'none' | 'prefix_partial' | 'complete' | 'conflict'
  completedCount: number
  /** IDs read from cNvPr in the package, never Office SDK shape IDs. */
  observed: { operationIndex: number; packageShapeId: string }[]
}
type Xml = Record<string, any>
const ordered = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  ignoreDeclaration: true,
})
const parser = new XMLParser({
  ignoreAttributes: false,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  ignoreDeclaration: true,
})
const many = (value: any): Xml[] =>
  value === undefined ? [] : Array.isArray(value) ? value : [value]
const conflict = (): NativeAddObservation => ({
  status: 'conflict',
  completedCount: 0,
  observed: [],
})
const requireProof = (condition: unknown): void => {
  if (!condition) throw Error('native_add_unproven')
}
const checkAbort = (signal?: AbortSignal): void => {
  if (signal?.aborted) throw Error('cancelled')
}
const close = (actual: unknown, expected: number, scale = 1): boolean => {
  const value = typeof actual === 'string' && /^[0-9]+$/.test(actual) ? Number(actual) : NaN
  return Number.isSafeInteger(value) && Math.abs(value / scale - expected) <= 0.01
}
const color = (fill: Xml | undefined, expected: string): boolean => {
  const rgb = fill?.['a:srgbClr']
  return (
    !!rgb &&
    Object.keys(fill!).join(',') === 'a:srgbClr' &&
    Object.keys(rgb).join(',') === '@_val' &&
    typeof rgb['@_val'] === 'string' &&
    rgb['@_val'].toUpperCase() === expected.toUpperCase()
  )
}

/** Preserve child order and exact text whitespace; normalize attribute ordering and XML indentation only. */
function canonical(value: any, text = false): any {
  if (Array.isArray(value))
    return value
      .filter(
        (node) =>
          text ||
          !(
            Object.keys(node).length === 1 &&
            typeof node['#text'] === 'string' &&
            /^\s*$/.test(node['#text'])
          ),
      )
      .map((node) => canonical(node, text))
  if (!value || typeof value !== 'object') return value
  return Object.fromEntries(
    Object.entries(value)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([key, child]) => [
        key,
        canonical(
          child,
          text ||
            key === 'a:t' ||
            key === 'a:instrText' ||
            value[':@']?.['@_xml:space'] === 'preserve',
        ),
      ]),
  )
}
function parse(xml: string): { tree: Xml[]; root: Xml } {
  requireProof(
    new TextEncoder().encode(xml).byteLength <= MAX_PPTX_XML_BYTES &&
      !/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml) &&
      XMLValidator.validate(xml) === true,
  )
  const clean = (value: any, text = false): any => {
    if (Array.isArray(value)) return value.map((node) => clean(node, text))
    if (!value || typeof value !== 'object') return value
    return Object.fromEntries(
      Object.entries(value)
        .filter(
          ([key, child]) =>
            text || key !== '#text' || typeof child !== 'string' || !/^\s*$/.test(child),
        )
        .map(([key, child]) => [
          key,
          clean(child, text || key === 'a:t' || value['@_xml:space'] === 'preserve'),
        ]),
    )
  }
  return { tree: ordered.parse(xml), root: clean(parser.parse(xml)) }
}
function validateIds(nodes: Xml[]): Set<string> {
  const ids = new Set<string>()
  const names = new Set<string>()
  const visit = (items: Xml[]): void => {
    for (const node of items) {
      if ('p:cNvPr' in node) {
        const id = node[':@']?.['@_id']
        requireProof(
          typeof id === 'string' &&
            /^(?:0|[1-9][0-9]{0,9})$/.test(id) &&
            Number(id) <= 4294967295 &&
            !ids.has(id),
        )
        ids.add(id)
        const name = node[':@']?.['@_name']
        if (name !== undefined && name !== '') {
          requireProof(typeof name === 'string' && name.length <= 256 && !names.has(name))
          names.add(name)
        }
      }
      for (const [key, child] of Object.entries(node))
        if (key !== ':@' && Array.isArray(child)) visit(child)
    }
  }
  visit(nodes)
  return names
}
function findTree(nodes: Xml[], tag: string): Xml[] | undefined {
  for (const node of nodes) {
    if (Array.isArray(node[tag])) return node[tag]
    for (const [key, child] of Object.entries(node))
      if (key !== ':@' && Array.isArray(child)) {
        const found = findTree(child, tag)
        if (found) return found
      }
  }
  return undefined
}
function identity(node: Xml): { name: string; id: string } | undefined {
  if (
    !['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp', 'p:contentPart'].some(
      (tag) => tag in node,
    )
  )
    return undefined
  const search = (nodes: Xml[]): Xml | undefined => {
    for (const value of nodes) {
      if ('p:cNvPr' in value) return value[':@']
      for (const [key, child] of Object.entries(value))
        if (key !== ':@' && Array.isArray(child)) {
          const result = search(child)
          if (result) return result
        }
    }
    return undefined
  }
  const props = search([node])
  if (!props) return undefined
  requireProof(
    typeof props['@_name'] === 'string' &&
      props['@_name'].length > 0 &&
      props['@_name'].length <= 256 &&
      typeof props['@_id'] === 'string' &&
      /^(?:0|[1-9][0-9]{0,9})$/.test(props['@_id']),
  )
  return { name: props['@_name'], id: props['@_id'] }
}
function onlyKeys(value: Xml | string | undefined, keys: string[]): boolean {
  return (
    value === undefined ||
    (typeof value === 'string' && !value.trim()) ||
    (typeof value === 'object' && Object.keys(value).every((key) => keys.includes(key)))
  )
}
function textMatches(
  body: Xml,
  expected: string,
  operation: Pick<
    Extract<NativeAddOperation, { op: 'add_text_box' }>,
    'fontFace' | 'fontSize' | 'color' | 'bold' | 'italic' | 'align'
  >,
): void {
  requireProof(
    body && Object.keys(body).every((key) => ['a:bodyPr', 'a:lstStyle', 'a:p'].includes(key)),
  )
  const bodyProperties = body['a:bodyPr']
  requireProof(
    onlyKeys(bodyProperties, [
      '@_wrap',
      '@_lIns',
      '@_rIns',
      '@_tIns',
      '@_bIns',
      '@_rtlCol',
      '@_anchor',
      '@_vert',
    ]) &&
      [undefined, 'square'].includes(bodyProperties?.['@_wrap']) &&
      [undefined, '0', 'false'].includes(bodyProperties?.['@_rtlCol']) &&
      [undefined, 'horz'].includes(bodyProperties?.['@_vert']) &&
      onlyKeys(body['a:lstStyle'], []),
  )
  const paragraphs = many(body['a:p'])
  const strings: string[] = []
  for (const paragraph of paragraphs) {
    requireProof(
      Object.keys(paragraph).every((key) => ['a:pPr', 'a:r', 'a:endParaRPr'].includes(key)),
    )
    requireProof(
      onlyKeys(paragraph['a:pPr'], ['@_algn', '@_marL', '@_marR', '@_indent', 'a:buNone']) &&
        onlyKeys(paragraph['a:pPr']?.['a:buNone'], []),
    )
    const end = paragraph['a:endParaRPr']
    requireProof(
      onlyKeys(end, ['@_lang', '@_dirty', '@_sz', 'a:latin', 'a:ea', 'a:cs']) &&
        (end?.['@_sz'] === undefined ||
          (operation.fontSize !== undefined && close(end['@_sz'], operation.fontSize, 100))),
    )
    for (const font of ['a:latin', 'a:ea', 'a:cs'])
      requireProof(
        onlyKeys(end?.[font], ['@_typeface', '@_pitchFamily', '@_charset']) &&
          (end?.[font] === undefined || end[font]?.['@_typeface'] === operation.fontFace),
      )
    if (operation.align)
      requireProof(
        (paragraph['a:pPr']?.['@_algn'] ?? 'l') ===
          { left: 'l', center: 'ctr', right: 'r' }[operation.align],
      )
    for (const field of ['@_marL', '@_marR', '@_indent'])
      requireProof(paragraph['a:pPr']?.[field] === undefined || paragraph['a:pPr'][field] === '0')
    const runs = many(paragraph['a:r'])
    strings.push(
      runs
        .map((run) => {
          requireProof(Object.keys(run).every((key) => ['a:rPr', 'a:t'].includes(key)))
          const properties = run['a:rPr']
          requireProof(
            onlyKeys(properties, [
              '@_lang',
              '@_sz',
              '@_b',
              '@_dirty',
              '@_i',
              '@_u',
              '@_strike',
              'a:solidFill',
              'a:latin',
              'a:ea',
              'a:cs',
            ]) &&
              (operation.italic === true ? ['1', 'true'] : [undefined, '0', 'false']).includes(
                properties?.['@_i'],
              ) &&
              [undefined, 'none'].includes(properties?.['@_u']) &&
              [undefined, 'noStrike'].includes(properties?.['@_strike']),
          )
          for (const font of ['a:latin', 'a:ea', 'a:cs'])
            requireProof(
              onlyKeys(properties?.[font], ['@_typeface', '@_pitchFamily', '@_charset']) &&
                (properties?.[font] === undefined ||
                  properties[font]?.['@_typeface'] === operation.fontFace),
            )
          if (operation.fontFace !== undefined)
            requireProof(properties?.['a:latin']?.['@_typeface'] === operation.fontFace)
          if (operation.fontSize !== undefined)
            requireProof(close(properties?.['@_sz'], operation.fontSize, 100))
          if (operation.color !== undefined)
            requireProof(color(properties?.['a:solidFill'], operation.color))
          if (operation.bold !== undefined)
            requireProof(
              (properties?.['@_b'] === '1' || properties?.['@_b'] === 'true') === operation.bold &&
                [undefined, '0', 'false', '1', 'true'].includes(properties?.['@_b']),
            )
          requireProof(typeof run['a:t'] === 'string' || typeof run['a:t']?.['#text'] === 'string')
          return typeof run['a:t'] === 'string' ? run['a:t'] : run['a:t']['#text']
        })
        .join(''),
    )
  }
  requireProof(equivalentPowerPointText(strings.join('\n'), expected))
}
function matches(element: Xml, operation: NativeAddOperation): void {
  requireProof(
    !/"(?:a:extLst|mc:AlternateContent|a:effectLst|a:effectDag|a:hlinkClick|a:hlinkHover|p:ph)":/.test(
      JSON.stringify(element),
    ),
  )
  requireProof(
    Object.keys(element).every((key) =>
      (operation.op === 'add_native_table'
        ? ['p:nvGraphicFramePr', 'p:xfrm', 'a:graphic']
        : ['p:nvSpPr', 'p:spPr', 'p:txBody']
      ).includes(key),
    ),
  )
  const table = operation.op === 'add_native_table'
  const properties = (element['p:nvSpPr'] ?? element['p:nvGraphicFramePr'])?.['p:cNvPr']
  requireProof(properties && [undefined, '0', 'false'].includes(properties['@_hidden']))
  const transform = table ? element['p:xfrm'] : element['p:spPr']?.['a:xfrm']
  requireProof(
    transform &&
      [undefined, '0'].includes(transform['@_rot']) &&
      ['@_flipH', '@_flipV'].every((key) => [undefined, '0', 'false'].includes(transform[key])),
  )
  for (const [group, field, expected] of [
    ['a:off', '@_x', operation.left],
    ['a:off', '@_y', operation.top],
    ['a:ext', '@_cx', operation.width],
    ['a:ext', '@_cy', operation.height],
  ] as const)
    requireProof(close(transform[group]?.[field], expected, 12700))
  if (operation.op === 'add_text_box') {
    const shapeProperties = element['p:spPr']
    requireProof(
      onlyKeys(shapeProperties, ['a:xfrm', 'a:prstGeom', 'a:noFill', 'a:ln']) &&
        onlyKeys(shapeProperties?.['a:prstGeom'], ['@_prst', 'a:avLst']) &&
        [undefined, 'rect'].includes(shapeProperties?.['a:prstGeom']?.['@_prst']) &&
        onlyKeys(shapeProperties?.['a:prstGeom']?.['a:avLst'], []) &&
        onlyKeys(shapeProperties?.['a:noFill'], []) &&
        onlyKeys(shapeProperties?.['a:ln'], []),
    )
    requireProof(
      ['1', 'true'].includes(element['p:nvSpPr']?.['p:cNvSpPr']?.['@_txBox']) &&
        element['p:txBody'],
    )
    textMatches(element['p:txBody'], operation.text, operation)
    const body = element['p:txBody']?.['a:bodyPr']
    if (operation.margin !== undefined)
      for (const side of ['lIns', 'rIns', 'tIns', 'bIns'])
        requireProof(close(body?.[`@_${side}`], operation.margin, 12700))
    if (operation.verticalAlignment !== undefined)
      requireProof(
        (body?.['@_anchor'] ?? 't') ===
          { top: 't', middle: 'ctr', bottom: 'b' }[operation.verticalAlignment],
      )
  } else if (operation.op === 'add_geometric_shape') {
    requireProof(
      [undefined, '0', 'false'].includes(element['p:nvSpPr']?.['p:cNvSpPr']?.['@_txBox']),
    )
    const properties = element['p:spPr']
    requireProof(
      Object.keys(properties).every((key) =>
        ['a:xfrm', 'a:prstGeom', 'a:solidFill', 'a:ln'].includes(key),
      ) && !properties['a:prstGeom']?.['a:avLst']?.['a:gd'],
    )
    requireProof(
      properties?.['a:prstGeom']?.['@_prst'] === operation.shape &&
        color(properties['a:solidFill'], operation.fill) &&
        color(properties['a:ln']?.['a:solidFill'], operation.lineColor),
    )
    requireProof(
      !element['p:txBody'] ||
        many(element['p:txBody']['a:p']).every((p) => !p['a:r'] && !p['a:fld']),
    )
  } else {
    const content = element['a:graphic']?.['a:graphicData']?.['a:tbl']
    requireProof(content)
    const structure = inspectNativeTableStructure(element)
    requireProof(
      !structure.merges.length &&
        structure.rows.length === operation.rows.length &&
        structure.columns.length === operation.rows[0]!.length,
    )
    requireProof(
      structure.rows.every(
        (height) => Math.abs(height / 12700 - operation.height / operation.rows.length) <= 0.01,
      ) &&
        structure.columns.every(
          (width) => Math.abs(width / 12700 - operation.width / operation.rows[0]!.length) <= 0.01,
        ),
    )
    const cells = many(content['a:tr']).map((row) => many(row['a:tc']))
    for (const [r, row] of cells.entries())
      for (const [c, cell] of row.entries()) {
        textMatches(cell['a:txBody'], operation.rows[r]![c]!, {
          fontFace: operation.fontFace,
          fontSize: operation.fontSize,
          color: operation.color,
        })
        if (operation.cellMargin !== undefined)
          for (const side of ['marL', 'marR', 'marT', 'marB'])
            requireProof(close(cell['a:tcPr']?.[`@_${side}`], operation.cellMargin, 12700))
      }
    if (operation.borderColor !== undefined)
      requireProof(
        inspectNativeTableCellStyles(element)!.every((cell) =>
          cell.borders.every(
            ([width, rgb]) =>
              close(width, 1, 12700) && rgb === operation.borderColor!.toUpperCase(),
          ),
        ),
      )
  }
}

/** Conservative proof for additions to an exported single-page package. Unknown changes never imply success. */
export async function observePowerPointNativeAdd(
  beforeBase64: string,
  currentBase64: string,
  operations: readonly NativeAddOperation[],
  signal?: AbortSignal,
): Promise<NativeAddObservation> {
  checkAbort(signal)
  try {
    requireProof(
      Array.isArray(operations) &&
        operations.length > 0 &&
        operations.length <= 32 &&
        new TextEncoder().encode(JSON.stringify(operations)).byteLength <= 128 * 1024 &&
        new Set(operations.map((operation) => operation.name)).size === operations.length,
    )
    for (const operation of operations) {
      const keys = [
        'op',
        'slide_index',
        'name',
        'left',
        'top',
        'width',
        'height',
        ...(operation.op === 'add_text_box'
          ? [
              'text',
              'fontFace',
              'fontSize',
              'color',
              'bold',
              'italic',
              'align',
              'margin',
              'verticalAlignment',
            ]
          : operation.op === 'add_native_table'
            ? ['rows', 'fontFace', 'fontSize', 'color', 'borderColor', 'cellMargin']
            : ['shape', 'fill', 'lineColor']),
      ]
      requireProof(
        Object.keys(operation).every((key) => keys.includes(key)) &&
          ['add_text_box', 'add_geometric_shape', 'add_native_table'].includes(operation.op) &&
          Number.isSafeInteger(operation.slide_index) &&
          operation.slide_index >= 0 &&
          operation.slide_index <= 31 &&
          operation.slide_index === operations[0]!.slide_index &&
          typeof operation.name === 'string' &&
          operation.name.length > 0 &&
          operation.name.length <= 256 &&
          !Array.from(operation.name).some((char) => {
            const code = char.charCodeAt(0)
            return code < 32 || (code >= 127 && code <= 159)
          }) &&
          [operation.left, operation.top, operation.width, operation.height].every(
            Number.isFinite,
          ) &&
          operation.width > 0 &&
          operation.height > 0,
      )
      const fields = operation as unknown as Xml
      for (const field of ['fontSize', 'margin', 'cellMargin'])
        if (fields[field] !== undefined)
          requireProof(
            typeof fields[field] === 'number' &&
              Number.isFinite(fields[field]) &&
              fields[field] >= 0 &&
              fields[field] <= 1000,
          )
      for (const field of ['color', 'fill', 'lineColor', 'borderColor'])
        if (fields[field] !== undefined)
          requireProof(typeof fields[field] === 'string' && /^[0-9A-Fa-f]{6}$/.test(fields[field]))
      if (fields.fontFace !== undefined)
        requireProof(
          typeof fields.fontFace === 'string' &&
            fields.fontFace.length > 0 &&
            fields.fontFace.length <= 128,
        )
      if (fields.fontSize !== undefined) requireProof(fields.fontSize > 0 && fields.fontSize <= 400)
      if (operation.op === 'add_native_table')
        requireProof(
          typeof operation.fontFace === 'string' &&
            typeof operation.fontSize === 'number' &&
            typeof operation.color === 'string',
        )
      if (operation.op === 'add_text_box')
        requireProof(
          (operation.bold === undefined || typeof operation.bold === 'boolean') &&
            (operation.italic === undefined || typeof operation.italic === 'boolean') &&
            (operation.align === undefined ||
              ['left', 'center', 'right'].includes(operation.align)) &&
            (operation.verticalAlignment === undefined ||
              ['top', 'middle', 'bottom'].includes(operation.verticalAlignment)),
        )
      if (operation.op === 'add_native_table')
        requireProof(
          Array.isArray(operation.rows) &&
            operation.rows.length > 0 &&
            operation.rows.length <= 20 &&
            operation.rows[0]!.length > 0 &&
            operation.rows[0]!.length <= 12 &&
            operation.rows.length * operation.rows[0]!.length <= 128 &&
            operation.rows.every(
              (row) =>
                Array.isArray(row) &&
                row.length === operation.rows[0]!.length &&
                row.every((cell) => typeof cell === 'string' && cell.length <= 256),
            ),
        )
      if (operation.op === 'add_text_box')
        requireProof(typeof operation.text === 'string' && operation.text.length <= 12000)
    }
    const [before, current] = await Promise.all([
      loadBoundedZip(beforeBase64, signal),
      loadBoundedZip(currentBase64, signal),
    ])
    const paths = Object.keys(before.files)
      .filter((path) => !before.files[path]!.dir)
      .sort()
    requireProof(
      [
        '[Content_Types].xml',
        '_rels/.rels',
        'ppt/presentation.xml',
        'ppt/_rels/presentation.xml.rels',
      ].every((path) => paths.includes(path)),
    )
    requireProof(
      JSON.stringify(paths) ===
        JSON.stringify(
          Object.keys(current.files)
            .filter((path) => !current.files[path]!.dir)
            .sort(),
        ),
    )
    const slides = paths.filter((path) => /^ppt\/slides\/slide[0-9]+\.xml$/.test(path))
    requireProof(slides.length === 1)
    const slidePath = slides[0]!
    const source = parse(await before.file(slidePath)!.async('string'))
    const actual = parse(await current.file(slidePath)!.async('string'))
    const allOldNames = validateIds(source.tree)
    validateIds(actual.tree)
    requireProof(operations.every((operation) => !allOldNames.has(operation.name)))
    const oldTree = findTree(source.tree, 'p:spTree'),
      newTree = findTree(actual.tree, 'p:spTree')
    requireProof(oldTree && newTree && oldTree.length <= 1004 && newTree.length <= 1036)
    const names = new Set(operations.map((operation) => operation.name))
    const oldIds = oldTree!.map(identity).filter((value) => value !== undefined)
    requireProof(
      new Set(oldIds.map((value) => value.name)).size === oldIds.length &&
        new Set(oldIds.map((value) => value.id)).size === oldIds.length &&
        oldIds.every((value) => !names.has(value.name)),
    )
    const added: { node: Xml; name: string; id: string }[] = []
    const seenNames = new Set<string>(),
      seenIds = new Set<string>()
    for (const node of newTree!) {
      const value = identity(node)
      if (!value) continue
      requireProof(!seenNames.has(value.name) && !seenIds.has(value.id))
      seenNames.add(value.name)
      seenIds.add(value.id)
      if (names.has(value.name)) added.push({ node, ...value })
    }
    requireProof(
      added.every(
        (value, index) =>
          value.name === operations[index]?.name && !oldIds.some((old) => old.id === value.id),
      ),
    )
    const retained = newTree!.filter((node) => !added.some((value) => value.node === node))
    newTree!.splice(0, newTree!.length, ...retained)
    requireProof(JSON.stringify(canonical(source.tree)) === JSON.stringify(canonical(actual.tree)))
    const body = actual.root['p:sld']?.['p:cSld']?.['p:spTree']
    for (const [index, addition] of added.entries()) {
      const operation = operations[index]!
      const tag = operation.op === 'add_native_table' ? 'p:graphicFrame' : 'p:sp'
      requireProof(tag in addition.node)
      const element = many(body?.[tag]).find(
        (value) =>
          (value['p:nvSpPr'] ?? value['p:nvGraphicFramePr'])?.['p:cNvPr']?.['@_name'] ===
          addition.name,
      )
      requireProof(element)
      matches(element!, operation)
    }
    for (const path of paths) {
      checkAbort(signal)
      if (path === slidePath) continue
      const [a, b] = await Promise.all([
        before.file(path)!.async('uint8array'),
        current.file(path)!.async('uint8array'),
      ])
      if (/\.(xml|rels)$/.test(path)) {
        const original = parse(new TextDecoder('utf-8', { fatal: true }).decode(a))
        const exported = parse(new TextDecoder('utf-8', { fatal: true }).decode(b))
        requireProof(
          JSON.stringify(canonical(original.tree)) === JSON.stringify(canonical(exported.tree)),
        )
        if (path.endsWith('.rels')) {
          const relations = many(original.root.Relationships?.Relationship)
          requireProof(new Set(relations.map((rel) => rel['@_Id'])).size === relations.length)
          for (const relation of relations) {
            const target = relation['@_Target']
            requireProof(
              typeof relation['@_Id'] === 'string' &&
                relation['@_Id'] &&
                typeof relation['@_Type'] === 'string' &&
                typeof target === 'string' &&
                target.length > 0,
            )
            if (relation['@_TargetMode'] === 'External') {
              requireProof(relation['@_Type'].endsWith('/hyperlink'))
              continue
            }
            requireProof(relation['@_TargetMode'] === undefined && !/[\\%?#]/.test(target))
            const base = path === '_rels/.rels' ? '' : path.replace(/_rels\/[^/]+\.rels$/, '')
            const parts: string[] = []
            for (const segment of (target.startsWith('/') ? target.slice(1) : base + target).split(
              '/',
            )) {
              if (segment === '..') {
                requireProof(parts.length)
                parts.pop()
              } else if (segment !== '.') {
                requireProof(segment)
                parts.push(segment)
              }
            }
            requireProof(paths.includes(parts.join('/')))
          }
        }
      } else requireProof(a.length === b.length && a.every((byte, index) => byte === b[index]))
    }
    checkAbort(signal)
    return {
      status: !added.length
        ? 'none'
        : added.length === operations.length
          ? 'complete'
          : 'prefix_partial',
      completedCount: added.length,
      observed: added.map((value, operationIndex) => ({
        operationIndex,
        packageShapeId: value.id,
      })),
    }
  } catch {
    checkAbort(signal)
    return conflict()
  }
}
