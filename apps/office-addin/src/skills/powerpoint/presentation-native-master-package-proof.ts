import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { loadBoundedZip, MAX_PPTX_XML_BYTES } from './powerpoint-package.js'
import { masterOperationKey } from './presentation-master-program.js'
import type { StoredMasterOperation } from './presentation-native-master-change.js'

export interface MasterPackageProtection {
  digest: string
  targetDigests: Record<string, string>
  originalPaths: string[]
  contentTypeDefaultExtensions: string[]
}
type Node = Record<string, any>
type Dependency = { masterId: string; layoutId: string }
const fail = (): never => {
  throw Error('presentation_native_master_package_unproven')
}
const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  ignoreDeclaration: true,
  commentPropName: '#comment',
})
const sha = async (bytes: Uint8Array) =>
  Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes).buffer)),
    (b) => b.toString(16).padStart(2, '0'),
  ).join('')
const textBytes = (v: unknown) => new TextEncoder().encode(JSON.stringify(v))
const one = (nodes: Node[], tag: string): Node => {
  const found = nodes.filter((n) => Object.hasOwn(n, tag))
  if (found.length !== 1 || !Array.isArray(found[0]![tag])) fail()
  return found[0]!
}
const optional = (nodes: Node[], tag: string): Node | undefined => {
  const found = nodes.filter((n) => Object.hasOwn(n, tag))
  if (found.length > 1 || (found[0] && !Array.isArray(found[0][tag]))) fail()
  return found[0]
}
function canonical(v: any, preserveText = false): any {
  if (Array.isArray(v))
    return v
      .filter(
        (n) =>
          preserveText ||
          !(
            Object.keys(n).length === 1 &&
            typeof n['#text'] === 'string' &&
            /^\s*$/.test(n['#text'])
          ),
      )
      .map((n) => canonical(n, preserveText))
  if (!v || typeof v !== 'object') return v
  if (
    Object.hasOwn(v, 'a:srgbClr') &&
    typeof v[':@']?.['@_val'] === 'string' &&
    /^[0-9a-f]{6}$/i.test(v[':@']['@_val'])
  )
    v = { ...v, ':@': { ...v[':@'], '@_val': v[':@']['@_val'].toUpperCase() } }
  if (Array.isArray(v['a:srgbClr']) && v['a:srgbClr'].length === 1) {
    const alpha = v['a:srgbClr'][0],
      value = alpha?.[':@']?.['@_val']
    if (
      Object.keys(alpha ?? {})
        .sort()
        .join(',') === ':@,a:alpha' &&
      Array.isArray(alpha['a:alpha']) &&
      alpha['a:alpha'].length === 0 &&
      Object.keys(alpha[':@']).join(',') === '@_val' &&
      typeof value === 'string' &&
      /^\d+$/.test(value) &&
      Number(value) <= 100000
    )
      v = {
        ...v,
        'a:srgbClr':
          Number(value) === 100000 ? [] : [{ ...alpha, ':@': { '@_val': String(Number(value)) } }],
      }
  }

  return Object.fromEntries(
    Object.entries(v)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([k, child]) => [
        k,
        canonical(
          child,
          preserveText ||
            k === 'a:t' ||
            k === '#comment' ||
            v[':@']?.['@_xml:space'] === 'preserve',
        ),
      ]),
  )
}
function pathFor(source: string, target: string): string {
  if (!target || /[\\:%?#]/.test(target) || target.startsWith('/')) fail()
  const parts = source.slice(0, source.lastIndexOf('/')).split('/')
  for (const p of target.split('/')) {
    if (p === '..') {
      if (!parts.length) fail()
      parts.pop()
    } else if (p !== '.') {
      if (!p) fail()
      parts.push(p)
    }
  }
  return parts.join('/')
}
const relsFor = (path: string) =>
  path.slice(0, path.lastIndexOf('/') + 1) +
  '_rels/' +
  path.slice(path.lastIndexOf('/') + 1) +
  '.rels'
const slot = (name: string) =>
  (
    ({
      Accent1: 'accent1',
      Accent2: 'accent2',
      Accent3: 'accent3',
      Accent4: 'accent4',
      Accent5: 'accent5',
      Accent6: 'accent6',
      Dark1: 'dk1',
      Dark2: 'dk2',
      Light1: 'lt1',
      Light2: 'lt2',
      Hyperlink: 'hlink',
      FollowedHyperlink: 'folHlink',
    }) as Record<string, string>
  )[name]

function inverseFillProven(bg: Node | undefined, inverse?: StoredMasterOperation) {
  const reject = (): never => {
    throw Error('presentation_native_master_inverse_unproven')
  }
  const children = (node: Node, tag: string, attrs: string[] = []): Node[] => {
    if (
      Object.keys(node).some((k) => k !== tag && k !== ':@') ||
      Object.keys(node[':@'] ?? {}).some((k) => !attrs.includes(k)) ||
      !Array.isArray(node[tag])
    )
      reject()
    return node[tag].filter(
      (n: Node) =>
        !(
          Object.keys(n).length === 1 &&
          typeof n['#text'] === 'string' &&
          /^\s*$/.test(n['#text'])
        ),
    )
  }
  if (!bg) reject()
  const bgChildren = children(bg!, 'p:bg')
  if (bgChildren.length !== 1) reject()
  const fills = children(bgChildren[0]!, 'p:bgPr')
  if (fills.length !== 1) reject()
  const fill = fills[0]!
  const rgb = (node: Node, alphaAllowed: boolean): { color: string; transparency: number } => {
    const transforms = children(node, 'a:srgbClr', ['@_val'])
    const color = node[':@']?.['@_val']
    if (
      typeof color !== 'string' ||
      !/^[0-9a-f]{6}$/i.test(color) ||
      transforms.length > (alphaAllowed ? 1 : 0)
    )
      reject()
    let alpha = 100000
    if (transforms.length) {
      const n = transforms[0]!
      if (children(n, 'a:alpha', ['@_val']).length) reject()
      const raw = n[':@']?.['@_val']
      if (typeof raw !== 'string' || !/^\d+$/.test(raw) || Number(raw) > 100000) reject()
      alpha = Number(raw)
    }
    return { color: '#' + color.toUpperCase(), transparency: 1 - alpha / 100000 }
  }
  if (Object.hasOwn(fill, 'a:solidFill')) {
    const colors = children(fill, 'a:solidFill')
    if (colors.length !== 1) reject()
    const value = rgb(colors[0]!, true)
    if (
      inverse &&
      (inverse.op !== 'set_master_background' ||
        inverse.fill.type !== 'solid' ||
        inverse.fill.color.toUpperCase() !== value.color ||
        Math.abs(inverse.fill.transparency - value.transparency) > 1e-10)
    )
      reject()
  } else if (Object.hasOwn(fill, 'a:pattFill')) {
    // Native pattern enum-to-OOXML mapping is not yet proved by the host contract.
    reject()
  } else reject()
}

async function protectedContent(
  base64: string,
  dep: Dependency,
  operations: StoredMasterOperation[],
  applied: StoredMasterOperation[],
  baseline: Omit<MasterPackageProtection, 'digest'> | undefined,
  signal?: AbortSignal,
  guard?: () => Promise<void>,
  inverses?: StoredMasterOperation[],
): Promise<MasterPackageProtection> {
  const io = async <T>(task: () => Promise<T>): Promise<T> => {
    if (signal?.aborted) throw Error('cancelled')
    await guard?.()
    const result = await task()
    if (signal?.aborted) throw Error('cancelled')
    await guard?.()
    return result
  }
  const zip = await io(() => loadBoundedZip(base64, signal))
  const paths = Object.values(zip.files)
    .filter((f) => !f.dir)
    .map((f) => f.name)
    .sort()
  const trees = new Map<string, Node[]>()
  const xml = async (path: string): Promise<Node[]> => {
    const existing = trees.get(path)
    if (existing) return existing
    const file = zip.file(path)
    if (!file) fail()
    const value = await io(() => file!.async('string'))
    if (
      new TextEncoder().encode(value).length > MAX_PPTX_XML_BYTES ||
      /<!\s*(DOCTYPE|ENTITY)\b/i.test(value) ||
      XMLValidator.validate(value) !== true
    )
      fail()
    const parsed = parser.parse(value) as Node[]
    trees.set(path, parsed)
    return parsed
  }
  const rel = async (source: string, type: string, rid?: string) => {
    const path = relsFor(source),
      tree = await xml(path),
      relationships = one(tree, 'Relationships')['Relationships'] as Node[]
    const found = relationships.filter(
      (n) =>
        n.Relationship &&
        n[':@']?.['@_Type'] ===
          `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}` &&
        (rid === undefined || n[':@']?.['@_Id'] === rid),
    )
    if (
      found.length !== 1 ||
      ![undefined, 'Internal'].includes(found[0]![':@']?.['@_TargetMode']) ||
      typeof found[0]![':@']?.['@_Target'] !== 'string'
    )
      fail()
    const target = pathFor(source, found[0]![':@']['@_Target'])
    if (!zip.file(target)) fail()
    return { target, path, tree, node: found[0]!, relationships }
  }
  const slides = paths.filter((p) => /^ppt[/]slides[/]slide[0-9]+[.]xml$/.test(p))
  if (slides.length !== 1) fail()
  const layoutPath = (await rel(slides[0]!, 'slideLayout')).target
  const masterPath = (await rel(layoutPath, 'slideMaster')).target
  const masterNode = one(await xml(masterPath), 'p:sldMaster'),
    layoutNode = one(await xml(layoutPath), 'p:sldLayout')
  const relevant = operations.filter((o) => o.master_id === dep.masterId)
  const masterBg = relevant.find((o) => o.op === 'set_master_background')
  const targetDigests: Record<string, string> = {}
  const target = async (op: StoredMasterOperation, value: unknown) => {
    targetDigests[masterOperationKey(op as any)] = await io(() => sha(textBytes(canonical(value))))
  }
  const removed = new Set<string>()
  const allowedNew = new Set<string>()
  let imageExtension: string | undefined, imageMime: string | undefined
  if (masterBg) {
    const children = one(masterNode['p:sldMaster'], 'p:cSld')['p:cSld'] as Node[]
    const bg = optional(children, 'p:bg')
    if (!baseline)
      inverseFillProven(
        bg,
        inverses?.find((o) => o.op === 'set_master_background' && o.master_id === dep.masterId),
      )
    await target(masterBg, bg ?? null)
    const fill = bg && optional(bg['p:bg'], 'p:bgPr')
    const blipFill = fill && optional(fill['p:bgPr'], 'a:blipFill')
    const blip = blipFill && optional(blipFill['a:blipFill'], 'a:blip')
    const picture = applied.find(
      (o) => o.op === 'set_master_background' && o.master_id === dep.masterId,
    )
    if (picture?.op === 'set_master_background' && picture.fill.type === 'picture_or_texture') {
      const embed = blip?.[':@']?.['@_r:embed']
      if (typeof embed !== 'string' || blip?.[':@']?.['@_r:link'] !== undefined) fail()
      const image = await rel(masterPath, 'image', embed),
        bytes = await io(() => zip.file(image.target)!.async('uint8array'))
      if (
        bytes.length !== picture.fill.imageRef.sizeBytes ||
        (await io(() => sha(bytes))) !== picture.fill.imageRef.sha256
      )
        fail()
      imageExtension = image.target.split('.').at(-1)?.toLowerCase()
      imageMime =
        bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47
          ? 'image/png'
          : bytes[0] === 0xff && bytes[1] === 0xd8
            ? 'image/jpeg'
            : undefined
      if (
        !imageMime ||
        (imageMime === 'image/png'
          ? imageExtension !== 'png'
          : !['jpg', 'jpeg'].includes(imageExtension!))
      )
        fail()
      const position = image.relationships.indexOf(image.node)
      image.relationships.splice(position, 1)
      if (!baseline?.originalPaths.includes(image.target)) {
        removed.add(image.target)
        allowedNew.add(image.target)
      }
    } else if (blip) {
      // The original native inverse supports no picture background. Do not erase unproved media.
      fail()
    }
    const position = bg && children.indexOf(bg)
    if (typeof position === 'number') children.splice(position, 1)
  }
  const colors = relevant.filter((o) => o.op === 'set_master_theme_color')
  if (colors.length) {
    const theme = one(await xml((await rel(masterPath, 'theme')).target), 'a:theme')
    const elements = one(theme['a:theme'], 'a:themeElements'),
      scheme = one(elements['a:themeElements'], 'a:clrScheme')
    for (const op of colors) {
      if (op.op !== 'set_master_theme_color') continue
      const node = one(scheme['a:clrScheme'], `a:${slot(op.theme_color)}`)
      if (!baseline) {
        const children = node[`a:${slot(op.theme_color)}`].filter(
          (n: Node) => !Object.hasOwn(n, '#text'),
        )
        const rgb = children[0]
        if (
          children.length !== 1 ||
          !rgb?.['a:srgbClr'] ||
          rgb['a:srgbClr'].length ||
          Object.keys(rgb[':@'] ?? {}).join(',') !== '@_val' ||
          Object.keys(rgb).sort().join(',') !== ':@,a:srgbClr' ||
          Object.keys(node).some((k) => k !== `a:${slot(op.theme_color)}`) ||
          typeof rgb[':@']?.['@_val'] !== 'string' ||
          !/^[0-9a-f]{6}$/i.test(rgb[':@']['@_val']) ||
          (inverses &&
            !inverses.some(
              (inv) =>
                inv.op === 'set_master_theme_color' &&
                inv.master_id === op.master_id &&
                inv.theme_color === op.theme_color &&
                inv.color.toUpperCase() === '#' + rgb[':@']['@_val'].toUpperCase(),
            ))
        )
          throw Error('presentation_native_master_inverse_unproven')
      }
      await target(op, node)
    }
    const remove = new Set(
      colors.map(
        (o) =>
          `a:${slot((o as Extract<StoredMasterOperation, { op: 'set_master_theme_color' }>).theme_color)}`,
      ),
    )
    scheme['a:clrScheme'] = (scheme['a:clrScheme'] as Node[]).filter(
      (n) => !Object.keys(n).some((k) => remove.has(k)),
    )
  }
  if (
    relevant.some((o) => o.op === 'set_layout_background_following' && o.layout_id === dep.layoutId)
  ) {
    const op = relevant.find(
      (o) => o.op === 'set_layout_background_following' && o.layout_id === dep.layoutId,
    )!
    const originalChildren = one(layoutNode['p:sldLayout'], 'p:cSld')['p:cSld'] as Node[]
    if (
      !baseline &&
      optional(originalChildren, 'p:bg') &&
      op.op === 'set_layout_background_following' &&
      op.follow_master
    )
      throw Error('presentation_native_master_inverse_unproven')
    await target(op, {
      background: optional(originalChildren, 'p:bg') ?? null,
      showMasterSp: !['false', '0'].includes(layoutNode[':@']?.['@_showMasterSp']),
    })
    if (layoutNode[':@']) delete layoutNode[':@']['@_showMasterSp']
    const children = one(layoutNode['p:sldLayout'], 'p:cSld')['p:cSld'] as Node[]
    const bg = optional(children, 'p:bg')
    if (bg) children.splice(children.indexOf(bg), 1)
  }
  let extensions: string[] = []
  if (zip.file('[Content_Types].xml')) {
    const types = one(await xml('[Content_Types].xml'), 'Types')
    extensions = (types.Types as Node[])
      .filter((n) => n.Default)
      .map((n) => n[':@']?.['@_Extension'])
    if (
      extensions.some((v) => typeof v !== 'string') ||
      new Set(extensions).size !== extensions.length
    )
      fail()
    if (imageExtension && !baseline?.contentTypeDefaultExtensions.includes(imageExtension)) {
      const defaults = (types.Types as Node[]).filter(
        (n) => n.Default && n[':@']?.['@_Extension'] === imageExtension,
      )
      if (
        defaults.length !== 1 ||
        defaults[0]![':@']?.['@_ContentType'] !== imageMime ||
        Object.keys(defaults[0]![':@']).sort().join(',') !== '@_ContentType,@_Extension'
      )
        fail()
      types.Types = (types.Types as Node[]).filter((n) => !defaults.includes(n))
    }
  } else if (imageExtension) fail()
  if (
    baseline &&
    (baseline.originalPaths.some((p) => !paths.includes(p)) ||
      paths.some((p) => !baseline.originalPaths.includes(p) && !allowedNew.has(p)))
  )
    fail()
  const entries: Array<[string, string]> = []
  for (const path of paths) {
    if (removed.has(path)) continue
    const bytes =
      trees.has(path) || path.endsWith('.xml') || path.endsWith('.rels')
        ? textBytes(canonical(await xml(path)))
        : await io(() => zip.file(path)!.async('uint8array'))
    entries.push([path, await io(() => sha(bytes))])
  }
  return {
    digest: await io(() => sha(textBytes(entries))),
    targetDigests,
    originalPaths: baseline?.originalPaths ?? paths,
    contentTypeDefaultExtensions: baseline?.contentTypeDefaultExtensions ?? extensions,
  }
}
export function prepareMasterPackageProtection(
  base64: string,
  dependency: Dependency,
  operations: StoredMasterOperation[],
  signal?: AbortSignal,
  guard?: () => Promise<void>,
  inverses?: StoredMasterOperation[],
) {
  return protectedContent(base64, dependency, operations, [], undefined, signal, guard, inverses)
}
export async function verifyMasterPackageProtection(
  base64: string,
  protection: MasterPackageProtection,
  dependency: Dependency,
  operations: StoredMasterOperation[],
  applied: StoredMasterOperation[],
  signal?: AbortSignal,
  guard?: () => Promise<void>,
) {
  if (!validMasterPackageProtectionTargets(protection, dependency, operations))
    throw Error('presentation_master_backup_invalid')
  const actual = await protectedContent(
    base64,
    dependency,
    operations,
    applied,
    protection,
    signal,
    guard,
  )
  const appliedKeys = new Set(applied.map((o) => masterOperationKey(o as any)))
  return (
    actual.digest === protection.digest &&
    Object.entries(protection.targetDigests).every(
      ([key, digest]) => appliedKeys.has(key) || actual.targetDigests[key] === digest,
    )
  )
}
export function validMasterPackageProtection(v: unknown): v is MasterPackageProtection {
  if (!v || typeof v !== 'object' || Array.isArray(v)) return false
  const p = v as MasterPackageProtection
  return (
    Object.keys(p).sort().join(',') ===
      'contentTypeDefaultExtensions,digest,originalPaths,targetDigests' &&
    !!p.targetDigests &&
    typeof p.targetDigests === 'object' &&
    !Array.isArray(p.targetDigests) &&
    Object.entries(p.targetDigests).every(
      ([key, value]) => !!key && /^[a-f0-9]{64}$/.test(value),
    ) &&
    typeof p.digest === 'string' &&
    /^[a-f0-9]{64}$/.test(p.digest) &&
    Array.isArray(p.originalPaths) &&
    !!p.originalPaths.length &&
    p.originalPaths.every((s) => typeof s === 'string' && !!s && s.length <= 4096) &&
    new Set(p.originalPaths).size === p.originalPaths.length &&
    Array.isArray(p.contentTypeDefaultExtensions) &&
    p.contentTypeDefaultExtensions.every((s) => typeof s === 'string' && !!s && s.length <= 256) &&
    new Set(p.contentTypeDefaultExtensions).size === p.contentTypeDefaultExtensions.length
  )
}

export function validMasterPackageProtectionTargets(
  protection: unknown,
  dependency: Dependency,
  operations: StoredMasterOperation[],
): boolean {
  if (!validMasterPackageProtection(protection)) return false
  const keys = operations
    .filter(
      (o) =>
        o.master_id === dependency.masterId &&
        (o.op !== 'set_layout_background_following' || o.layout_id === dependency.layoutId),
    )
    .map((o) => masterOperationKey(o as any))
    .sort()
  return JSON.stringify(Object.keys(protection.targetDigests).sort()) === JSON.stringify(keys)
}
