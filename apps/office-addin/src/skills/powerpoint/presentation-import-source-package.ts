import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { loadBoundedZip, MAX_PPTX_IMPORT_PAGE_BYTES } from './powerpoint-package.js'

const parser = new XMLParser({ ignoreAttributes: false, parseAttributeValue: false })
const validXml = (xml: string): boolean =>
  new TextEncoder().encode(xml).byteLength <= 512 * 1024 &&
  !/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml) &&
  XMLValidator.validate(xml) === true
const items = (value: unknown): Record<string, unknown>[] =>
  value === undefined ? [] : Array.isArray(value) ? value : [value as Record<string, unknown>]

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
  if (slideIds.length !== 1 || `${slideIds[0]!['@_id']}#` !== sourceSlideId) invalid()
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
}
