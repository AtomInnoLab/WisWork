import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { loadBoundedZip, MAX_PPTX_IMPORT_PAGE_BYTES } from './powerpoint-package.js'
import { inspectPowerPointChartSourcesBatch } from './presentation-chart-source-package.js'

const parser = new XMLParser({ ignoreAttributes: false, parseAttributeValue: false })
const HYPERLINK_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const INTERNAL_RELATIONSHIPS = new Set(
  [
    'slide',
    'slideLayout',
    'slideMaster',
    'theme',
    'notesSlide',
    'notesMaster',
    'image',
    'chart',
    'package',
    'presProps',
    'viewProps',
    'tableStyles',
  ].map((name) => `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${name}`),
)
export const validHttpLink = (target: unknown): boolean => {
  if (
    typeof target !== 'string' ||
    !target ||
    target.length > 2048 ||
    target.includes('\\') ||
    [...target].some((character) => {
      const code = character.charCodeAt(0)
      return code < 32 || code === 127
    })
  )
    return false
  try {
    const url = new URL(target)
    return (
      ['http:', 'https:'].includes(url.protocol) && !!url.hostname && !url.username && !url.password
    )
  } catch {
    return false
  }
}
const validXml = (xml: string): boolean =>
  new TextEncoder().encode(xml).byteLength <= 512 * 1024 &&
  !/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml) &&
  XMLValidator.validate(xml) === true
const validPngStructure = (bytes: Uint8Array): boolean => {
  if (
    bytes.length < 57 ||
    [137, 80, 78, 71, 13, 10, 26, 10].some((value, index) => bytes[index] !== value)
  )
    return false
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let offset = 8
  let chunks = 0
  let hasImageData = false
  while (offset + 12 <= bytes.length) {
    const length = view.getUint32(offset)
    if (length > bytes.length - offset - 12) return false
    const type = String.fromCharCode(...bytes.subarray(offset + 4, offset + 8))
    if (chunks++ === 0) {
      if (
        type !== 'IHDR' ||
        length !== 13 ||
        view.getUint32(offset + 8) === 0 ||
        view.getUint32(offset + 12) === 0
      )
        return false
    } else if (type === 'IDAT') {
      if (length > 0) hasImageData = true
    } else if (type === 'IEND') {
      return length === 0 && hasImageData && offset + 12 === bytes.length
    }
    offset += length + 12
  }
  return false
}
const validJpegStructure = (bytes: Uint8Array): boolean => {
  if (
    bytes.length < 20 ||
    bytes[0] !== 0xff ||
    bytes[1] !== 0xd8 ||
    bytes.at(-2) !== 0xff ||
    bytes.at(-1) !== 0xd9
  )
    return false
  let offset = 2
  let hasFrame = false
  while (offset + 4 < bytes.length) {
    if (bytes[offset++] !== 0xff) return false
    while (bytes[offset] === 0xff) offset++
    const marker = bytes[offset++]
    if (marker === undefined || marker === 0xd8 || marker === 0xd9) return false
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue
    if (offset + 2 > bytes.length) return false
    const length = (bytes[offset]! << 8) | bytes[offset + 1]!
    if (length < 2 || offset + length > bytes.length - 2) return false
    if (
      (marker >= 0xc0 && marker <= 0xc3) ||
      (marker >= 0xc5 && marker <= 0xc7) ||
      (marker >= 0xc9 && marker <= 0xcb) ||
      (marker >= 0xcd && marker <= 0xcf)
    ) {
      if (
        length < 8 ||
        !((bytes[offset + 3]! << 8) | bytes[offset + 4]!) ||
        !((bytes[offset + 5]! << 8) | bytes[offset + 6]!)
      )
        return false
      hasFrame = true
    }
    if (marker === 0xda) return hasFrame && offset + length < bytes.length - 2
    offset += length
  }
  return false
}
const items = (value: unknown): Record<string, unknown>[] =>
  value === undefined ? [] : Array.isArray(value) ? value : [value as Record<string, unknown>]
const record = (value: unknown): Record<string, unknown> | undefined =>
  value && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined
const relationshipRefs = (
  xml: Record<string, unknown>,
): { ids: string[]; hyperlinkIds: string[] } => {
  const ids: string[] = []
  const hyperlinkIds: string[] = []
  const visit = (value: unknown, tag?: string): void => {
    if (Array.isArray(value)) {
      value.forEach((item) => visit(item, tag))
      return
    }
    const node = record(value)
    if (!node) return
    for (const [key, child] of Object.entries(node)) {
      if (['@_r:id', '@_r:embed', '@_r:link'].includes(key)) {
        if (typeof child !== 'string' || !child)
          throw new Error('presentation_import_state_invalid')
        ids.push(child)
        if (tag === 'a:hlinkClick' && key === '@_r:id') hyperlinkIds.push(child)
      } else if (!key.startsWith('@_')) visit(child, key)
    }
  }
  visit(xml)
  return { ids, hyperlinkIds }
}
const internalTarget = (slidePath: string, target: string): string | undefined => {
  if (!target || target.includes('\\') || /[?#]/.test(target)) return
  const parts = target.startsWith('/') ? [] : slidePath.split('/').slice(0, -1)
  for (const segment of (target.startsWith('/') ? target.slice(1) : target).split('/')) {
    if (segment === '.') continue
    if (segment === '..') {
      if (parts.length === 0) return
      parts.pop()
    } else if (segment && segment !== '.' && !/%|:/.test(segment)) parts.push(segment)
    else return
  }
  return parts.join('/')
}
const typedTarget = (owner: string, type: string, target: string): string | undefined => {
  const path = internalTarget(owner, target)
  if (!path) return
  if (type.endsWith('/image') && !/^ppt\/media\/[^/]+\.(?:png|jpe?g)$/i.test(path)) return
  if (type.endsWith('/chart') && !/^ppt\/charts\/chart[0-9]+\.xml$/.test(path)) return
  if (type.endsWith('/package') && !/^ppt\/embeddings\/[A-Za-z0-9_.-]+\.xlsx$/.test(path)) return
  return path
}

/** Prove that a bounded production package contains the one slide selected for host import. */
export async function validatePresentationImportSourcePage(
  base64: string,
  sourceSlideId: string,
): Promise<void> {
  const invalid = (): never => {
    throw new Error('presentation_import_state_invalid')
  }
  if (!/^[1-9][0-9]{0,9}#$/.test(sourceSlideId)) invalid()
  const zip = await loadBoundedZip(
    base64,
    undefined,
    false,
    MAX_PPTX_IMPORT_PAGE_BYTES,
    MAX_PPTX_IMPORT_PAGE_BYTES,
  ).catch(() => invalid())
  const slidePaths = Object.keys(zip.files).filter((path) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(path),
  )
  if (slidePaths.length !== 1) invalid()
  const required = [
    '[Content_Types].xml',
    'ppt/presentation.xml',
    'ppt/_rels/presentation.xml.rels',
    slidePaths[0]!,
  ]
  const xmls = await Promise.all(required.map((path) => zip.file(path)?.async('string')))
  if (xmls.some((xml) => !xml || !validXml(xml))) invalid()
  const [typesXml, presentationXml, relsXml, slideXml] = xmls as string[]
  const types = parser.parse(typesXml).Types
  const presentation = parser.parse(presentationXml)['p:presentation']
  const rels = parser.parse(relsXml).Relationships
  const slide = parser.parse(slideXml)['p:sld']
  if (!types || !presentation || !rels || !slide?.['p:cSld']?.['p:spTree']) invalid()
  const defaults = new Map<string, string>()
  const overrides = new Map<string, string>()
  for (const entry of items(types.Default)) {
    const extension = entry['@_Extension'],
      mime = entry['@_ContentType']
    if (typeof extension === 'string' && typeof mime === 'string') {
      if (defaults.has(extension.toLowerCase())) invalid()
      defaults.set(extension.toLowerCase(), mime)
    } else invalid()
  }
  for (const entry of items(types.Override)) {
    const path = entry['@_PartName'],
      mime = entry['@_ContentType']
    if (typeof path === 'string' && typeof mime === 'string') {
      if (overrides.has(path)) invalid()
      overrides.set(path, mime)
    } else invalid()
  }
  const contentTypeFor = (path: string) =>
    overrides.get(`/${path}`) ?? defaults.get(path.slice(path.lastIndexOf('.') + 1).toLowerCase())
  if (
    contentTypeFor('ppt/presentation.xml') !==
      'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml' ||
    contentTypeFor(slidePaths[0]!) !==
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
  )
    invalid()
  const slideIds = items(presentation['p:sldIdLst']?.['p:sldId'])
  const numericId = Number(sourceSlideId.slice(0, -1))
  if (
    numericId < 256 ||
    numericId > 0xffffffff ||
    slideIds.length !== 1 ||
    `${slideIds[0]!['@_id']}#` !== sourceSlideId
  )
    invalid()
  const relId = slideIds[0]!['@_r:id']
  const relationships = items(rels.Relationship)
  if (
    typeof relId !== 'string' ||
    relationships.some((rel) => typeof rel['@_Id'] !== 'string') ||
    new Set(relationships.map((rel) => rel['@_Id'])).size !== relationships.length
  )
    invalid()
  const relationship = relationships.find((rel) => rel['@_Id'] === relId)
  if (
    !relationship ||
    typeof relationship['@_Type'] !== 'string' ||
    !relationship['@_Type'].endsWith('/slide') ||
    relationship['@_TargetMode'] !== undefined ||
    typeof relationship['@_Target'] !== 'string' ||
    !/^slides\/slide\d+\.xml$/.test(relationship['@_Target']) ||
    `ppt/${relationship['@_Target']}` !== slidePaths[0]
  )
    invalid()

  const referencedParts = new Map<string, 'png' | 'jpeg' | 'chart' | 'workbook'>()
  const rememberTypedPart = (owner: string, rel: Record<string, unknown>): void => {
    const type = rel['@_Type'] as string
    const path = typedTarget(owner, type, rel['@_Target'] as string)!
    if (type.endsWith('/image'))
      referencedParts.set(path, path.toLowerCase().endsWith('.png') ? 'png' : 'jpeg')
    else if (type.endsWith('/chart')) referencedParts.set(path, 'chart')
    else if (type.endsWith('/package')) referencedParts.set(path, 'workbook')
  }

  const slideRelsPath = slidePaths[0]!.replace('/slides/', '/slides/_rels/') + '.rels'
  const slideRelsFile = zip.file(slideRelsPath)
  const { ids: referencedIds, hyperlinkIds: hyperlinkRefs } = relationshipRefs({ 'p:sld': slide })
  const hyperlinkIds = new Set(hyperlinkRefs)
  const counts = (ids: string[]) => {
    const result = new Map<string, number>()
    for (const id of ids) result.set(id, (result.get(id) ?? 0) + 1)
    return result
  }
  const referenceCounts = counts(referencedIds),
    hyperlinkCounts = counts(hyperlinkRefs)
  const allowedExternal = (owner: string, rel: Record<string, unknown>) =>
    owner === slidePaths[0] &&
    rel['@_Type'] === HYPERLINK_REL &&
    rel['@_TargetMode'] === 'External' &&
    hyperlinkIds.has(rel['@_Id'] as string) &&
    referenceCounts.get(rel['@_Id'] as string) === hyperlinkCounts.get(rel['@_Id'] as string) &&
    validHttpLink(rel['@_Target'])
  const validPartRelationship = (owner: string, rel: Record<string, unknown>) =>
    typeof rel['@_Id'] === 'string' &&
    typeof rel['@_Type'] === 'string' &&
    typeof rel['@_Target'] === 'string' &&
    (allowedExternal(owner, rel) ||
      (rel['@_TargetMode'] === undefined &&
        INTERNAL_RELATIONSHIPS.has(rel['@_Type']) &&
        !!zip.file(typedTarget(owner, rel['@_Type'], rel['@_Target']) ?? '')))
  if (referencedIds.length && !slideRelsFile) invalid()
  if (slideRelsFile) {
    const slideRelsXml = await slideRelsFile.async('string')
    if (!validXml(slideRelsXml)) invalid()
    const slideRels = parser.parse(slideRelsXml).Relationships
    if (!slideRels) invalid()
    const references = items(slideRels.Relationship)
    if (
      references.some((rel) => !validPartRelationship(slidePaths[0]!, rel)) ||
      new Set(references.map((rel) => rel['@_Id'])).size !== references.length
    )
      invalid()
    references.forEach((rel) => rememberTypedPart(slidePaths[0]!, rel))
    const ids = new Set(references.map((rel) => rel['@_Id']))
    if (referencedIds.some((id) => !ids.has(id))) invalid()
  }

  // Charts and notes may have their own relationships (for example an editable chart workbook).
  // A valid slide relationship alone does not prove those second-level parts are present.
  for (const path of Object.keys(zip.files).filter((name) =>
    /(?:^|\/)_rels\/[^/]+\.rels$/.test(name),
  )) {
    const match = /^(.*?)_rels\/([^/]+)\.rels$/.exec(path)
    if (!match) throw new Error('presentation_import_state_invalid')
    const owner = `${match[1]}${match[2]}`
    if (!zip.file(owner)) invalid()
    const xml = await zip.file(path)!.async('string')
    if (!validXml(xml)) invalid()
    const root = parser.parse(xml).Relationships
    if (!root) invalid()
    const entries = items(root.Relationship)
    if (
      entries.some((entry) => !validPartRelationship(owner, entry)) ||
      new Set(entries.map((entry) => entry['@_Id'])).size !== entries.length
    )
      invalid()
    entries.forEach((entry) => rememberTypedPart(owner, entry))
  }

  for (const [path, kind] of referencedParts) {
    const expectedMime = {
      png: 'image/png',
      jpeg: 'image/jpeg',
      chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }[kind]
    if (contentTypeFor(path) !== expectedMime) invalid()
    const file = zip.file(path) ?? invalid()
    if (kind === 'workbook') {
      const book = await loadBoundedZip(
        await file.async('base64'),
        undefined,
        true,
        MAX_PPTX_IMPORT_PAGE_BYTES,
        MAX_PPTX_IMPORT_PAGE_BYTES,
      ).catch(() => invalid())
      const contentTypes = await book.file('[Content_Types].xml')?.async('string')
      const workbookXml = await book.file('xl/workbook.xml')?.async('string')
      const workbookRelsXml = await book.file('xl/_rels/workbook.xml.rels')?.async('string')
      if (
        !contentTypes ||
        !workbookXml ||
        !workbookRelsXml ||
        !validXml(contentTypes) ||
        !validXml(workbookXml) ||
        !validXml(workbookRelsXml)
      )
        invalid()
      const sheets = items(parser.parse(workbookXml!)['workbook']?.['sheets']?.['sheet'])
      const rels = items(parser.parse(workbookRelsXml!).Relationships?.Relationship)
      if (
        sheets.length === 0 ||
        sheets.some((sheet) => {
          const id = sheet['@_r:id']
          const rel = rels.find((entry) => entry['@_Id'] === id)
          if (
            !rel ||
            typeof rel['@_Type'] !== 'string' ||
            !rel['@_Type'].endsWith('/worksheet') ||
            rel['@_TargetMode'] !== undefined ||
            typeof rel['@_Target'] !== 'string'
          )
            return true
          const sheetPath = internalTarget('xl/workbook.xml', rel['@_Target'])
          return (
            !sheetPath || !/^xl\/worksheets\/[^/]+\.xml$/.test(sheetPath) || !book.file(sheetPath)
          )
        })
      )
        invalid()
      for (const sheet of sheets) {
        const rel = rels.find((entry) => entry['@_Id'] === sheet['@_r:id'])!
        const sheetXml = await book
          .file(internalTarget('xl/workbook.xml', rel['@_Target'] as string)!)!
          .async('string')
        if (!validXml(sheetXml) || !parser.parse(sheetXml)['worksheet']) invalid()
      }
      continue
    }
    if (kind === 'chart') {
      const xml = await file.async('string')
      if (!validXml(xml) || !parser.parse(xml)['c:chartSpace']?.['c:chart']) invalid()
      continue
    }
    const bytes = await file.async('uint8array')
    if (kind === 'png') {
      if (!validPngStructure(bytes)) invalid()
    } else if (!validJpegStructure(bytes)) invalid()
  }

  const tree = record(record(slide['p:cSld'])?.['p:spTree'])
  const chartShapeIds = items(tree?.['p:graphicFrame'])
    .filter((frame) => record(record(frame['a:graphic'])?.['a:graphicData'])?.['c:chart'])
    .map((frame) => record(record(frame['p:nvGraphicFramePr'])?.['p:cNvPr'])?.['@_id'])
  if (chartShapeIds.some((id) => typeof id !== 'string' && typeof id !== 'number')) invalid()
  if (
    new Set(chartShapeIds.map(String)).size !== chartShapeIds.length ||
    ([...referencedParts.values()].includes('chart') && chartShapeIds.length === 0)
  )
    invalid()
  if (chartShapeIds.length) {
    const checked = await inspectPowerPointChartSourcesBatch(
      base64,
      chartShapeIds.map(String),
      undefined,
      { maxBytes: MAX_PPTX_IMPORT_PAGE_BYTES, allowAbsoluteChartTarget: true },
    ).catch(() => invalid())
    if (
      checked.unsupported.length ||
      Object.keys(checked.reports).length !== chartShapeIds.length ||
      Object.values(checked.reports).some(
        (report) => report.sourceKind !== 'embedded_xlsx' || report.verification !== 'matches',
      )
    )
      invalid()
  }

  for (const path of Object.keys(zip.files).filter(
    (name) => name.startsWith('ppt/') && name.endsWith('.xml'),
  )) {
    const ownerXml = await zip.file(path)!.async('string')
    const ids = relationshipRefs(parser.parse(ownerXml)).ids
    if (ids.length === 0) continue
    const slash = path.lastIndexOf('/')
    const relsPath = `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`
    const relsFile = zip.file(relsPath)
    if (!relsFile) throw new Error('presentation_import_state_invalid')
    const relsXml = await relsFile.async('string')
    if (!validXml(relsXml)) invalid()
    const root = parser.parse(relsXml).Relationships
    if (!root) invalid()
    const defined = new Set(items(root.Relationship).map((entry) => entry['@_Id']))
    if (ids.some((id) => !defined.has(id))) invalid()
  }
}
