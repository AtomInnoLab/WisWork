import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { loadBoundedZip, MAX_PPTX_IMPORT_PAGE_BYTES } from './powerpoint-package.js'

const parser = new XMLParser({ ignoreAttributes: false, parseAttributeValue: false })
const validXml = (xml: string): boolean =>
  new TextEncoder().encode(xml).byteLength <= 512 * 1024 &&
  !/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml) &&
  XMLValidator.validate(xml) === true
const items = (value: unknown): Record<string, unknown>[] =>
  value === undefined ? [] : Array.isArray(value) ? value : [value as Record<string, unknown>]
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

  const slideRelsPath = slidePaths[0]!.replace('/slides/', '/slides/_rels/') + '.rels'
  const slideRelsFile = zip.file(slideRelsPath)
  const referencedIds = [...slideXml.matchAll(/\br:(?:id|embed|link)="([^"]+)"/g)].map(
    (match) => match[1]!,
  )
  if (referencedIds.length && !slideRelsFile) invalid()
  if (slideRelsFile) {
    const slideRelsXml = await slideRelsFile.async('string')
    if (!validXml(slideRelsXml)) invalid()
    const slideRels = parser.parse(slideRelsXml).Relationships
    if (!slideRels) invalid()
    const references = items(slideRels.Relationship)
    if (
      references.some(
        (rel) =>
          typeof rel['@_Id'] !== 'string' ||
          typeof rel['@_Type'] !== 'string' ||
          typeof rel['@_Target'] !== 'string' ||
          rel['@_TargetMode'] !== undefined ||
          !zip.file(
            typedTarget(slidePaths[0]!, rel['@_Type'] as string, rel['@_Target'] as string) ?? '',
          ),
      ) ||
      new Set(references.map((rel) => rel['@_Id'])).size !== references.length
    )
      invalid()
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
      entries.some(
        (entry) =>
          typeof entry['@_Id'] !== 'string' ||
          typeof entry['@_Type'] !== 'string' ||
          typeof entry['@_Target'] !== 'string' ||
          entry['@_TargetMode'] !== undefined ||
          !zip.file(
            typedTarget(owner, entry['@_Type'] as string, entry['@_Target'] as string) ?? '',
          ),
      ) ||
      new Set(entries.map((entry) => entry['@_Id'])).size !== entries.length
    )
      invalid()
  }

  for (const path of Object.keys(zip.files).filter(
    (name) => name.startsWith('ppt/') && name.endsWith('.xml'),
  )) {
    const ownerXml = await zip.file(path)!.async('string')
    const ids = [...ownerXml.matchAll(/\br:(?:id|embed|link)="([^"]+)"/g)].map((match) => match[1]!)
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
