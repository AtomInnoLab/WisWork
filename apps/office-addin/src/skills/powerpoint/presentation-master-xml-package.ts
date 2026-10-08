import { XMLParser, XMLValidator } from 'fast-xml-parser'
import JSZip from 'jszip'
import {
  loadBoundedZip,
  MAX_PPTX_ENTRY_BYTES,
  MAX_PPTX_PACKAGE_BYTES,
  MAX_PPTX_XML_BYTES,
} from './powerpoint-package.js'
export interface MasterXmlPackageInventory {
  sourceSlideId: string
  sourceLayoutPath: string
  sourceMasterPath: string
  packageDigest: string
  masters: {
    path: string
    contentDigest: string
    orderedLayouts: {
      path: string
      packageLayoutId: string
      relationshipId: string
      contentDigest: string
    }[]
  }[]
}
export interface MasterXmlNativeInventory {
  masters: { masterId: string; layouts: { layoutId: string; representativeBase64: string }[] }[]
}
export interface MasterXmlLayoutMapping {
  masterId: string
  sourceLayoutId: string
  layouts: { packageLayoutPath: string; nativeLayoutId: string }[]
}
type Node = Record<string, any>
type Rel = {
  id: string
  type: string
  target: string
  attributes: Record<string, string>
  external: boolean
}
type Package = {
  inventory: MasterXmlPackageInventory
  bytes: Map<string, Uint8Array>
  projectedPageDigest(): Promise<string>
}
function fail(): never {
  throw Error('presentation_master_xml_package_unproven')
}
const check = (signal?: AbortSignal) => {
  if (signal?.aborted) throw Error('cancelled')
}
const controls = (value: string) =>
  Array.from(value).some(
    (c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159),
  )
const masterPath = /^ppt\/slideMasters\/[A-Za-z0-9._-]+\.xml$/
const layoutPath = /^ppt\/slideLayouts\/[A-Za-z0-9._-]+\.xml$/
const editablePath = (path: string) =>
  masterPath.test(path) || layoutPath.test(path) || /^ppt\/theme\/theme[0-9]+\.xml$/.test(path)
const prefix = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  processEntities: false,
  ignoreDeclaration: true,
  commentPropName: '#comment',
})
const sha = async (bytes: Uint8Array) =>
  Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes).buffer)),
    (b) => b.toString(16).padStart(2, '0'),
  ).join('')
const encoded = (value: unknown) => new TextEncoder().encode(JSON.stringify(value))
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const children = (nodes: Node[]) =>
  nodes.filter(
    (n) =>
      !(Object.keys(n).length === 1 && typeof n['#text'] === 'string' && /^\s*$/.test(n['#text'])),
  )
const one = (nodes: Node[], tag: string): Node => {
  const found = children(nodes).filter((n) => Object.hasOwn(n, tag))
  if (found.length !== 1 || !Array.isArray(found[0]![tag])) fail()
  return found[0]!
}
function targetPath(source: string, target: string): string {
  if (!target || /[\\:%?#]/.test(target) || controls(target) || target.startsWith('/')) fail()
  const parts = source
    .slice(0, source.lastIndexOf('/') + 1)
    .split('/')
    .filter(Boolean)
  for (const part of target.split('/')) {
    if (part === '..') {
      if (!parts.length) fail()
      parts.pop()
    } else if (part !== '.') {
      if (!part) fail()
      parts.push(part)
    }
  }
  const result = parts.join('/')
  if (!result.startsWith('ppt/')) fail()
  return result
}
async function boundedBytes(
  file: JSZip.JSZipObject,
  remaining: number,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  check(signal)
  const stream = (
    file as JSZip.JSZipObject & {
      internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array>
    }
  ).internalStream('uint8array')
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = []
    let size = 0,
      done = false
    const finish = (error?: Error) => {
      if (done) return
      done = true
      signal?.removeEventListener('abort', abort)
      if (error) {
        stream.pause()
        chunks.length = 0
        reject(error)
      } else {
        const result = new Uint8Array(size)
        let offset = 0
        for (const chunk of chunks) {
          result.set(chunk, offset)
          offset += chunk.length
        }
        resolve(result)
      }
    }
    const abort = () => finish(Error('cancelled'))
    signal?.addEventListener('abort', abort, { once: true })
    stream
      .on('data', (chunk) => {
        if (done) return
        size += chunk.length
        if (signal?.aborted) abort()
        else if (size > MAX_PPTX_ENTRY_BYTES || size > remaining)
          finish(Error('presentation_master_xml_package_unproven'))
        else chunks.push(Uint8Array.from(chunk))
      })
      .on('error', () => finish(Error('presentation_master_xml_package_unproven')))
      .on('end', () => finish())
      .resume()
  })
}
function canonical(value: any, relationships: Map<string, string>, preserve = false): any {
  if (Array.isArray(value))
    return value
      .filter(
        (n) =>
          preserve ||
          !(
            Object.keys(n).length === 1 &&
            typeof n['#text'] === 'string' &&
            /^\s*$/.test(n['#text'])
          ),
      )
      .map((n) => canonical(n, relationships, preserve))
  if (!value || typeof value !== 'object') return value
  return Object.fromEntries(
    Object.entries(value)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([key, child]) => {
        if (key === ':@') {
          const attrs = child as Record<string, string>
          return [
            key,
            Object.fromEntries(
              Object.entries(attrs)
                .sort(([a], [b]) => a.localeCompare(b))
                .map(([k, v]) => {
                  if (['@_r:id', '@_r:embed', '@_r:link'].includes(k)) {
                    const resolved = relationships.get(v)
                    if (!resolved) fail()
                    return [k, resolved]
                  }
                  if (k === '@_PartName' && Object.hasOwn(value, 'Override')) {
                    const target = relationships.get('part:' + v.replace(/^\//, ''))
                    if (!target) fail()
                    return [k, target]
                  }
                  if (
                    k === '@_id' &&
                    ['p:sldLayoutId', 'p:sldId', 'p:sldMasterId'].some((tag) =>
                      Object.hasOwn(value, tag),
                    )
                  )
                    return [k, 'generated-layout-id']
                  if (
                    k === '@_val' &&
                    Object.hasOwn(value, 'a:srgbClr') &&
                    /^[a-f0-9]{6}$/i.test(v)
                  )
                    return [k, v.toUpperCase()]
                  return [k, v]
                }),
            ),
          ]
        }
        return [
          key,
          canonical(
            child,
            relationships,
            preserve ||
              key === 'a:t' ||
              key === '#comment' ||
              value[':@']?.['@_xml:space'] === 'preserve',
          ),
        ]
      }),
  )
}
async function parsePackage(base64: string, signal?: AbortSignal): Promise<Package> {
  check(signal)
  const zip = await loadBoundedZip(base64, signal, false),
    bytes = new Map<string, Uint8Array>()
  let total = 0
  for (const path of Object.keys(zip.files).sort()) {
    const file = zip.files[path]!
    if (file.dir) continue
    const value = await boundedBytes(file, MAX_PPTX_PACKAGE_BYTES - total, signal)
    total += value.length
    bytes.set(path, value)
  }
  const xml = (path: string): Node[] => {
    const value = bytes.get(path)
    if (!value || value.length > MAX_PPTX_XML_BYTES) fail()
    let text: string
    try {
      text = new TextDecoder('utf-8', { fatal: true }).decode(value)
    } catch {
      fail()
    }
    if (/<!\s*(DOCTYPE|ENTITY)\b/i.test(text!) || XMLValidator.validate(text!) !== true) fail()
    const nodes = parser.parse(text!)
    if (!Array.isArray(nodes)) fail()
    return nodes
  }
  const relations = new Map<string, Rel[]>()
  for (const path of bytes.keys())
    if (path.endsWith('.rels')) {
      const source =
        path === '_rels/.rels' ? '' : path.replace(/(^|\/)_rels\/([^/]+)\.rels$/, '$1$2')
      if (source && !bytes.has(source)) fail()
      const root = one(xml(path), 'Relationships'),
        nodes = children(root.Relationships),
        ids = new Set<string>(),
        entries: Rel[] = []
      for (const n of nodes) {
        if (!Object.hasOwn(n, 'Relationship') || children(n.Relationship).length) fail()
        const attrs = n[':@'] ?? {},
          id = attrs['@_Id'],
          type = attrs['@_Type'],
          target = attrs['@_Target']
        if (
          typeof id !== 'string' ||
          !id ||
          ids.has(id) ||
          typeof type !== 'string' ||
          !type ||
          typeof target !== 'string' ||
          (attrs['@_TargetMode'] !== undefined &&
            !['Internal', 'External'].includes(attrs['@_TargetMode']))
        )
          fail()
        ids.add(id)
        if (attrs['@_TargetMode'] === 'External') {
          // Preserve exact stored leaves without fetching or certifying remote content.
          if (
            ![
              'hyperlink',
              'image',
              'audio',
              'video',
              'package',
              'oleObject',
              'externalLink',
              'externalLinkPath',
            ].some((name) => type === prefix + name) ||
            !target ||
            target.length > 8192 ||
            controls(target)
          )
            fail()
          entries.push({ id, type, target, attributes: attrs, external: true })
          continue
        }
        // Root package relationships can point to docProps; their bytes are protected by preparation.
        let resolved: string
        if (!source) {
          if (
            /[\\:%?#]/.test(target) ||
            target.startsWith('/') ||
            target.split('/').some((p: string) => !p || p === '.' || p === '..')
          )
            fail()
          resolved = target
        } else resolved = targetPath(source, target)
        if (!bytes.has(resolved)) fail()
        entries.push({ id, type, target: resolved, attributes: attrs, external: false })
      }
      relations.set(source, entries)
    }
  const rels = (path: string) => relations.get(path) ?? []
  const typed = (path: string, type: string) => rels(path).filter((r) => r.type === prefix + type)
  const resolveRel = (path: string, id: unknown, type: string) => {
    const matches = typed(path, type).filter((r) => r.id === id)
    if (matches.length !== 1) fail()
    return matches[0]!
  }
  const presentation = 'ppt/presentation.xml',
    p = one(xml(presentation), 'p:presentation'),
    pchildren = p['p:presentation']
  const slides = children(one(pchildren, 'p:sldIdLst')['p:sldIdLst'])
  if (slides.length !== 1 || !Object.hasOwn(slides[0]!, 'p:sldId')) fail()
  const slideNode = slides[0]!,
    slideId = slideNode[':@']?.['@_id']
  if (
    typeof slideId !== 'string' ||
    !/^[1-9][0-9]{0,9}$/.test(slideId) ||
    Number(slideId) < 256 ||
    Number(slideId) > 4294967295
  )
    fail()
  const slide = resolveRel(presentation, slideNode[':@']?.['@_r:id'], 'slide').target
  if (
    !/^ppt\/slides\/slide[0-9]+\.xml$/.test(slide) ||
    [...bytes.keys()].filter((p) => /^ppt\/slides\/slide[0-9]+\.xml$/.test(p)).length !== 1 ||
    typed(presentation, 'slide').length !== 1
  )
    fail()
  one(xml(slide), 'p:sld')
  const selectedLayouts = typed(slide, 'slideLayout')
  if (selectedLayouts.length !== 1) fail()
  const sourceLayoutPath = selectedLayouts[0]!.target
  const masterNodes = children(one(pchildren, 'p:sldMasterIdLst')['p:sldMasterIdLst']),
    masterPaths: string[] = []
  const masterIds = new Set<string>()
  for (const n of masterNodes) {
    const id = n[':@']?.['@_id']
    if (
      !Object.hasOwn(n, 'p:sldMasterId') ||
      typeof id !== 'string' ||
      !/^[1-9][0-9]{0,9}$/.test(id) ||
      masterIds.has(id) ||
      Number(id) < 2147483648 ||
      Number(id) > 4294967295
    )
      fail()
    masterIds.add(id)
    masterPaths.push(resolveRel(presentation, n[':@']?.['@_r:id'], 'slideMaster').target)
  }
  if (
    !masterPaths.length ||
    new Set(masterPaths).size !== masterPaths.length ||
    typed(presentation, 'slideMaster').length !== masterPaths.length ||
    !same([...masterPaths].sort(), [...bytes.keys()].filter((p) => masterPath.test(p)).sort())
  )
    fail()
  const masters: MasterXmlPackageInventory['masters'] = [],
    allLayouts = new Set<string>()
  const owners = new Map<string, string>()
  for (const path of masterPaths) {
    const node = one(xml(path), 'p:sldMaster'),
      list = children(one(node['p:sldMaster'], 'p:sldLayoutIdLst')['p:sldLayoutIdLst']),
      ids = new Set<string>(),
      layouts: MasterXmlPackageInventory['masters'][number]['orderedLayouts'] = []
    for (const n of list) {
      const id = n[':@']?.['@_id'],
        rid = n[':@']?.['@_r:id']
      if (
        !Object.hasOwn(n, 'p:sldLayoutId') ||
        typeof id !== 'string' ||
        !/^[1-9][0-9]{0,9}$/.test(id) ||
        ids.has(id) ||
        Number(id) > 4294967295 ||
        Number(id) < 1
      )
        fail()
      ids.add(id)
      const target = resolveRel(path, rid, 'slideLayout').target
      if (!layoutPath.test(target) || allLayouts.has(target)) fail()
      allLayouts.add(target)
      owners.set(target, path)
      const back = typed(target, 'slideMaster')
      if (back.length !== 1 || back[0]!.target !== path) fail()
      one(xml(target), 'p:sldLayout')
      layouts.push({ path: target, packageLayoutId: id, relationshipId: rid, contentDigest: '' })
    }
    if (!layouts.length || typed(path, 'slideLayout').length !== layouts.length) fail()
    masters.push({ path, contentDigest: '', orderedLayouts: layouts })
  }
  if (
    !same([...allLayouts].sort(), [...bytes.keys()].filter((p) => layoutPath.test(p)).sort()) ||
    !owners.has(sourceLayoutPath)
  )
    fail()
  const resourceDigests = new Map<string, string>()
  const partDigest = async (
    path: string,
    labels: Map<string, string>,
    visiting: Set<string>,
  ): Promise<string> => {
    check(signal)
    if (visiting.has(path)) fail()
    const chain = new Set(visiting).add(path),
      rmap = new Map<string, string>(),
      edges: unknown[] = []
    for (const r of rels(path)) {
      const target = r.external
        ? 'external:' + r.target
        : (labels.get(r.target) ?? (await resource(r.target, chain)))
      rmap.set(r.id, target)
      edges.push({
        ...r.attributes,
        '@_Id': undefined,
        '@_Target': target,
        '@_TargetMode': r.external ? 'External' : undefined,
      })
    }
    const value = bytes.get(path)!
    const content = path.endsWith('.xml')
      ? canonical(xml(path), rmap)
      : { binarySha256: await sha(value) }
    return sha(
      encoded({ content, edges: edges.map((e) => JSON.stringify(canonical(e, new Map()))).sort() }),
    )
  }
  const resource = async (path: string, visiting: Set<string>): Promise<string> => {
    if (owners.has(path) || masterPaths.includes(path)) fail()
    let value = resourceDigests.get(path)
    if (!value) {
      value = await partDigest(path, new Map(), visiting)
      resourceDigests.set(path, value)
    }
    return value
  }
  for (const master of masters) {
    const labels = new Map<string, string>([
      [master.path, 'master'],
      ...master.orderedLayouts.map((l, i) => [l.path, `layout:${i}`] as [string, string]),
    ])
    for (const layout of master.orderedLayouts)
      layout.contentDigest = await partDigest(layout.path, labels, new Set())
    master.contentDigest = await sha(
      encoded({
        xml: await partDigest(master.path, labels, new Set()),
        layouts: master.orderedLayouts.map((l) => l.contentDigest),
      }),
    )
  }
  const projectedPageDigest = async (): Promise<string> => {
    const walk = (start: string[], pageOnly: boolean) => {
      const found = new Set<string>(),
        visit = (path: string) => {
          if (found.has(path)) return
          found.add(path)
          for (const r of rels(path))
            if (
              !r.external &&
              !(pageOnly && (owners.has(r.target) || masterPaths.includes(r.target)))
            )
              visit(r.target)
        }
      start.forEach(visit)
      return found
    }
    const pageOwned = walk([slide], true),
      masterOwned = walk(masterPaths, false),
      relationshipOwned = walk([''], false)
    const ignored = new Set([...masterOwned].filter((path) => !pageOwned.has(path)))
    const labels = new Map<string, string>([
      [slide, 'slide-root'],
      [presentation, 'presentation-root'],
      ['[Content_Types].xml', 'content-types'],
    ])
    for (const m of masters) {
      labels.set(m.path, 'master:' + m.contentDigest)
      m.orderedLayouts.forEach((l, i) => labels.set(l.path, 'layout:' + m.contentDigest + ':' + i))
    }
    const cache = new Map<string, string>()
    const resolve = async (path: string, visiting: Set<string>): Promise<string> => {
      const label = labels.get(path)
      if (label) return label
      if (visiting.has(path)) fail()
      let value = cache.get(path)
      if (!value) {
        value = await part(path, visiting)
        cache.set(path, value)
      }
      return value
    }
    const edge = async (r: Rel, path: string, visiting: Set<string>) => ({
      ...r.attributes,
      '@_Id': undefined,
      '@_Target': r.external
        ? 'external:' + r.target
        : path === slide && r.type === prefix + 'slideLayout'
          ? 'selected-layout'
          : await resolve(r.target, visiting),
      '@_TargetMode': r.external ? 'External' : undefined,
    })
    const part = async (path: string, visiting: Set<string>): Promise<string> => {
      check(signal)
      if (visiting.has(path)) fail()
      const next = new Set(visiting).add(path),
        rmap = new Map<string, string>(),
        edges: unknown[] = []
      for (const r of rels(path)) {
        if (path === presentation && r.type === prefix + 'slideMaster') continue
        const e = await edge(r, path, next)
        rmap.set(r.id, e['@_Target'])
        edges.push(e)
      }
      let content: unknown
      if (path.endsWith('.xml')) {
        let nodes = xml(path)
        if (path === presentation) {
          const root = one(nodes, 'p:presentation')
          nodes = [
            {
              ...root,
              'p:presentation': root['p:presentation'].filter(
                (n: Node) => !Object.hasOwn(n, 'p:sldMasterIdLst'),
              ),
            },
          ]
        }
        if (path === '[Content_Types].xml') {
          const root = one(nodes, 'Types'),
            entries: Node[] = [],
            seen = new Set<string>()
          for (const node of children(root.Types)) {
            if (Object.hasOwn(node, 'Override')) {
              const name = node[':@']?.['@_PartName']
              if (
                typeof name !== 'string' ||
                !name.startsWith('/') ||
                !bytes.has(name.slice(1)) ||
                seen.has(name)
              )
                fail()
              seen.add(name)
              if (ignored.has(name.slice(1))) continue
              rmap.set('part:' + name.slice(1), await resolve(name.slice(1), new Set()))
            }
            entries.push(node)
          }
          nodes = [{ ...root, Types: entries }]
        }
        content = canonical(nodes, rmap, true)
      } else content = { binarySha256: await sha(bytes.get(path)!) }
      return sha(
        encoded({
          content,
          edges: edges.map((e) => JSON.stringify(canonical(e, new Map()))).sort(),
        }),
      )
    }
    const entries: unknown[] = []
    for (const path of bytes.keys())
      if (!path.endsWith('.rels') && !ignored.has(path))
        entries.push({
          ...(relationshipOwned.has(path) ? {} : { exactPath: path }),
          folder: path.slice(0, path.lastIndexOf('/')),
          extension: path.split('.').at(-1),
          digest: await part(path, new Set()),
        })
    const rootEdges = await Promise.all(rels('').map((r) => edge(r, '', new Set())))
    check(signal)
    return sha(
      encoded({
        entries: entries.map((e) => JSON.stringify(e)).sort(),
        rootEdges: rootEdges.map((e) => JSON.stringify(canonical(e, new Map()))).sort(),
      }),
    )
  }
  const packageDigest = await sha(
    encoded(await Promise.all([...bytes].map(async ([path, value]) => [path, await sha(value)]))),
  )
  check(signal)
  return {
    bytes,
    projectedPageDigest,
    inventory: {
      sourceSlideId: `${slideId}#`,
      sourceLayoutPath,
      sourceMasterPath: owners.get(sourceLayoutPath)!,
      packageDigest,
      masters,
    },
  }
}
async function guarded<T>(run: () => Promise<T>, signal?: AbortSignal): Promise<T> {
  try {
    return await run()
  } catch (error) {
    if (signal?.aborted || (error instanceof Error && error.message === 'cancelled'))
      throw Error('cancelled', { cause: error })
    throw Error('presentation_master_xml_package_unproven', { cause: error })
  }
}
export function inspectMasterXmlPackage(
  base64: string,
  signal?: AbortSignal,
): Promise<MasterXmlPackageInventory> {
  return guarded(async () => (await parsePackage(base64, signal)).inventory, signal)
}
export function assertMasterXmlPreparation(
  originalBase64: string,
  preparedBase64: string,
  changedPaths: string[],
  signal?: AbortSignal,
): Promise<{
  original: MasterXmlPackageInventory
  prepared: MasterXmlPackageInventory
  affectedMasterPaths: string[]
}> {
  return guarded(async () => {
    const paths = structuredClone(changedPaths)
    if (
      paths.length < 1 ||
      paths.length > 32 ||
      new Set(paths).size !== paths.length ||
      paths.some((p) => !editablePath(p))
    )
      fail()
    const original = await parsePackage(originalBase64, signal),
      prepared = await parsePackage(preparedBase64, signal)
    if (
      !same([...original.bytes.keys()], [...prepared.bytes.keys()]) ||
      paths.some((p) => !original.bytes.has(p))
    )
      fail()
    for (const [path, value] of original.bytes) {
      if (paths.includes(path)) continue
      const next = prepared.bytes.get(path)!
      if (value.length !== next.length || value.some((b, i) => b !== next[i])) fail()
    }
    const identity = (v: MasterXmlPackageInventory) => ({
      ...v,
      packageDigest: undefined,
      masters: v.masters.map((m) => ({
        ...m,
        contentDigest: undefined,
        orderedLayouts: m.orderedLayouts.map((l) => ({ ...l, contentDigest: undefined })),
      })),
    })
    if (!same(identity(original.inventory), identity(prepared.inventory))) fail()
    return {
      original: original.inventory,
      prepared: prepared.inventory,
      affectedMasterPaths: original.inventory.masters
        .filter((m, i) => m.contentDigest !== prepared.inventory.masters[i]!.contentDigest)
        .map((m) => m.path),
    }
  }, signal)
}
/** Caller supplies actual native membership evidence. Every candidate layout needs its own representative.
 * An explicit secondary target returns its first OOXML layout as sourceLayoutId; layouts contains the complete mapping. */
export function proveMasterXmlLayoutMapping(
  expectedBase64: string,
  nativeInventory: MasterXmlNativeInventory,
  signal?: AbortSignal,
  targetMasterPath?: string,
): Promise<MasterXmlLayoutMapping> {
  return guarded(async () => {
    const owned = structuredClone(nativeInventory)
    const expected = (await parsePackage(expectedBase64, signal)).inventory,
      master = expected.masters.find(
        (m) => m.path === (targetMasterPath ?? expected.sourceMasterPath),
      ),
      candidates: MasterXmlLayoutMapping[] = []
    if (!master) fail()
    const ids = new Set<string>(),
      masterIds = new Set<string>(),
      validId = (v: unknown) =>
        typeof v === 'string' && !!v.trim() && v.length <= 256 && !controls(v)
    if (
      !owned ||
      !Array.isArray(owned.masters) ||
      !owned.masters.length ||
      encoded(
        owned.masters.map((m) => ({
          masterId: m.masterId,
          layouts: m.layouts.map((l) => ({ layoutId: l.layoutId })),
        })),
      ).length >
        4 * 1024 * 1024
    )
      fail()
    for (const native of owned.masters) {
      if (
        !validId(native.masterId) ||
        masterIds.has(native.masterId) ||
        !Array.isArray(native.layouts) ||
        !native.layouts.length
      )
        fail()
      masterIds.add(native.masterId)
      const map = new Map<number, string>()
      let matches = true
      for (const layout of native.layouts) {
        if (
          !validId(layout.layoutId) ||
          ids.has(layout.layoutId) ||
          typeof layout.representativeBase64 !== 'string'
        )
          fail()
        ids.add(layout.layoutId)
        const actual = (await parsePackage(layout.representativeBase64, signal)).inventory,
          selected = actual.masters.find((m) => m.path === actual.sourceMasterPath)!
        if (selected.contentDigest !== master.contentDigest) {
          matches = false
          continue
        }
        const slot = selected.orderedLayouts.findIndex((l) => l.path === actual.sourceLayoutPath)
        if (slot < 0 || map.has(slot)) fail()
        map.set(slot, layout.layoutId)
      }
      if (
        matches &&
        native.layouts.length === master.orderedLayouts.length &&
        map.size === master.orderedLayouts.length
      ) {
        const sourceSlot =
          master.path === expected.sourceMasterPath
            ? master.orderedLayouts.findIndex((l) => l.path === expected.sourceLayoutPath)
            : 0
        candidates.push({
          masterId: native.masterId,
          sourceLayoutId: map.get(sourceSlot)!,
          layouts: master.orderedLayouts.map((l, i) => ({
            packageLayoutPath: l.path,
            nativeLayoutId: map.get(i)!,
          })),
        })
      }
    }
    if (candidates.length !== 1) fail()
    check(signal)
    return candidates[0]!
  }, signal)
}

export interface MasterXmlPagePreservationOptions {
  expectedMasterBase64: string
  targetMasterPath: string
  packageLayoutPath: string
}
/** Proves local OOXML bytes and immutable external-reference leaves, never remote content or visual QA. */
export function assertMasterXmlPagePreserved(
  originalPageBase64: string,
  actualPageBase64: string,
  options: MasterXmlPagePreservationOptions,
  signal?: AbortSignal,
): Promise<void> {
  return guarded(async () => {
    const owned = structuredClone(options),
      expected = await parsePackage(owned.expectedMasterBase64, signal),
      original = await parsePackage(originalPageBase64, signal),
      actual = await parsePackage(actualPageBase64, signal)
    const wanted = expected.inventory.masters.find((m) => m.path === owned.targetMasterPath),
      old = original.inventory.masters.find((m) => m.path === original.inventory.sourceMasterPath)!,
      now = actual.inventory.masters.find((m) => m.path === actual.inventory.sourceMasterPath)!
    if (!wanted || wanted.contentDigest !== now.contentDigest) fail()
    const slot = wanted.orderedLayouts.findIndex((l) => l.path === owned.packageLayoutPath),
      oldSlot = old.orderedLayouts.findIndex((l) => l.path === original.inventory.sourceLayoutPath),
      nowSlot = now.orderedLayouts.findIndex((l) => l.path === actual.inventory.sourceLayoutPath)
    if (
      slot < 0 ||
      slot !== oldSlot ||
      slot !== nowSlot ||
      old.orderedLayouts.length !== wanted.orderedLayouts.length
    )
      fail()
    if ((await original.projectedPageDigest()) !== (await actual.projectedPageDigest())) fail()
    check(signal)
  }, signal)
}

/** Redirects only the package source's layout edge; it proves no native import or visual acceptance. */
export function deriveMasterXmlCarrier(
  base64: string,
  targetMasterPath: string,
  packageLayoutPath?: string,
  signal?: AbortSignal,
): Promise<{
  base64: string
  inventory: MasterXmlPackageInventory
  sourceSlideId: string
  sourceMasterPath: string
  sourceLayoutPath: string
}> {
  return guarded(async () => {
    check(signal)
    if (
      typeof base64 !== 'string' ||
      typeof targetMasterPath !== 'string' ||
      (packageLayoutPath !== undefined && typeof packageLayoutPath !== 'string')
    )
      fail()
    const original = await parsePackage(base64, signal),
      master = original.inventory.masters.find((m) => m.path === targetMasterPath)
    if (!master) fail()
    const layout = packageLayoutPath ?? master.orderedLayouts[0]!.path
    if (!master.orderedLayouts.some((l) => l.path === layout)) fail()
    let derivedBase64 = base64
    if (layout !== original.inventory.sourceLayoutPath) {
      const slide = [...original.bytes.keys()].find((p) =>
          /^ppt\/slides\/slide[0-9]+\.xml$/.test(p),
        )!,
        relationsPath = slide.replace(/\/([^/]+)$/, '/_rels/$1.rels'),
        raw = new TextDecoder('utf-8', { fatal: true }).decode(original.bytes.get(relationsPath)!)
      let replacements = 0
      const changed = raw.replace(
        /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<Relationship\b(?:[^"'<>]|"[^"]*"|'[^']*')*\/?>/g,
        (tag) => {
          if (!tag.startsWith('<Relationship')) return tag
          const fragment = tag.endsWith('/>') ? tag : tag.slice(0, -1) + '/>',
            node = one(parser.parse(fragment), 'Relationship')
          if (node[':@']?.['@_Type'] !== prefix + 'slideLayout') return tag
          replacements++
          return tag.replace(
            /(\sTarget\s*=\s*)(["'])([\s\S]*?)\2/,
            (_, lead, quote) => lead + quote + '../' + layout.slice(4) + quote,
          )
        },
      )
      if (replacements !== 1 || changed === raw) fail()
      const zip = new JSZip()
      for (const [path, bytes] of original.bytes)
        zip.file(path, path === relationsPath ? changed : bytes, { createFolders: false })
      derivedBase64 = await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' })
      check(signal)
    }
    const inventory = (await parsePackage(derivedBase64, signal)).inventory
    if (
      inventory.sourceSlideId !== original.inventory.sourceSlideId ||
      inventory.sourceMasterPath !== master.path ||
      inventory.sourceLayoutPath !== layout ||
      !same(inventory.masters, original.inventory.masters)
    )
      fail()
    return {
      base64: derivedBase64,
      inventory,
      sourceSlideId: inventory.sourceSlideId,
      sourceMasterPath: inventory.sourceMasterPath,
      sourceLayoutPath: inventory.sourceLayoutPath,
    }
  }, signal)
}
