export const MAX_PRESENTATION_XML_BYTES = 2 * 1024 * 1024
const MAX_XML_NODES = 20_000
const MAX_XML_DEPTH = 128
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace'
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/'
export const PRESENTATIONML_NAMESPACE = 'http://schemas.openxmlformats.org/presentationml/2006/main'
export const DRAWINGML_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const XML_TOKEN_PATTERN =
  /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<(?:[^< >"']|\s|"[^"]*"|'[^']*')*>|[^<]+/g

const isXmlCodePoint = (code: number) =>
  code === 0x9 ||
  code === 0xa ||
  code === 0xd ||
  (code >= 0x20 && code <= 0xd7ff) ||
  (code >= 0xe000 && code <= 0xfffd) ||
  (code >= 0x10000 && code <= 0x10ffff)

function validateXmlCharacters(value: string): void {
  for (const character of value) {
    if (!isXmlCodePoint(character.codePointAt(0)!))
      throw new TypeError('Presentation XML has an illegal character')
  }
}

function decodeXml(value: string): string {
  validateXmlCharacters(value)
  const decoded = value.replace(/&([^;]*);/g, (_match, entity: string) => {
    const named: Record<string, string> = { lt: '<', gt: '>', quot: '"', apos: "'", amp: '&' }
    if (entity in named) return named[entity]!
    const match = /^#(?:x([0-9a-f]+)|(\d+))$/i.exec(entity)
    if (!match) throw new TypeError('Presentation XML has an unknown entity')
    const code = Number.parseInt(match[1] ?? match[2]!, match[1] ? 16 : 10)
    if (!isXmlCodePoint(code))
      throw new TypeError('Presentation XML has an illegal character reference')
    return String.fromCodePoint(code)
  })
  if (decoded.includes('&')) {
    // Any ampersand left after replacement was not the beginning of a complete entity.
    // A decoded &amp; is valid, so compare against unconsumed source occurrences instead.
    const withoutEntities = value.replace(/&(?:lt|gt|quot|apos|amp|#(?:x[0-9a-f]+|\d+));/gi, '')
    if (withoutEntities.includes('&')) throw new TypeError('Presentation XML has an unknown entity')
  }
  validateXmlCharacters(decoded)
  return decoded
}

const qname = (raw: string, namespaces: ReadonlyMap<string, string>, attribute = false): string => {
  if (!/^(?:[A-Za-z_][A-Za-z0-9_.-]*:)?[A-Za-z_][A-Za-z0-9_.-]*$/.test(raw))
    throw new TypeError('Presentation XML has a malformed qualified name')
  const colon = raw.indexOf(':')
  const prefix = colon >= 0 ? raw.slice(0, colon) : ''
  const local = colon >= 0 ? raw.slice(colon + 1) : raw
  const namespace = prefix
    ? prefix === 'xml'
      ? XML_NAMESPACE
      : namespaces.get(prefix)
    : attribute
      ? ''
      : namespaces.get('')
  if (prefix && namespace === undefined)
    throw new TypeError('Presentation XML has an unbound prefix')
  return `{${namespace ?? ''}}${local}`
}

export interface PresentationBackgroundLocation {
  readonly cSldInsertAt: number
  readonly presentationPrefix: string
  readonly presentationDeclaration?: string
  readonly drawingPrefix: string
  readonly drawingDeclaration?: string
  readonly backgroundRange?: readonly [number, number]
}

interface LocationFrame {
  readonly name: string
  readonly rawName: string
  readonly namespaces: ReadonlyMap<string, string>
  readonly start: number
  readonly directBackground: boolean
  readonly isSlideCSld: boolean
}

/** Strictly scan a slide XML prefix and locate only a cSld-direct PresentationML background. */
export function locatePresentationBackground(xml: string): PresentationBackgroundLocation {
  if (new TextEncoder().encode(xml).byteLength > MAX_PRESENTATION_XML_BYTES)
    throw new TypeError('Presentation XML exceeds bounds')
  validateXmlCharacters(xml)
  const tokens = xml.match(XML_TOKEN_PATTERN)
  if (!tokens || tokens.join('') !== xml) throw new TypeError('Presentation XML is malformed')
  const stack: LocationFrame[] = []
  let offset = 0
  let nodes = 0
  let roots = 0
  let cSld:
    | {
        insertAt: number
        prefix: string
        namespaces: ReadonlyMap<string, string>
      }
    | undefined
  let backgroundRange: readonly [number, number] | undefined
  let backgroundPrefix: string | undefined
  let backgroundDeclaration: string | undefined

  for (const token of tokens) {
    const start = offset
    offset += token.length
    if (token.startsWith('<!--') || token.startsWith('<?')) continue
    if (token.startsWith('<![CDATA[')) {
      if (!stack.length) throw new TypeError('Presentation XML is malformed')
      validateXmlCharacters(token.slice(9, -3))
      continue
    }
    if (!token.startsWith('<')) {
      const text = decodeXml(token)
      if (!stack.length && text.trim()) throw new TypeError('Presentation XML is malformed')
      continue
    }
    if (token.startsWith('<!')) throw new TypeError('Unsupported presentation XML declaration')
    if (token.startsWith('</')) {
      const rawName = /^<\/\s*([^\s>]+)\s*>$/.exec(token)?.[1]
      const frame = stack.pop()
      if (!rawName || !frame || qname(rawName, frame.namespaces) !== frame.name)
        throw new TypeError('Presentation XML is malformed')
      if (frame.directBackground) {
        if (backgroundRange) throw new TypeError('Presentation XML has duplicate slide backgrounds')
        backgroundRange = [frame.start, offset]
      }
      continue
    }
    nodes += 1
    if (nodes > MAX_XML_NODES || stack.length >= MAX_XML_DEPTH)
      throw new TypeError('Presentation XML exceeds bounds')
    const selfClosing = /\/\s*>$/.test(token)
    const open = /^<\s*([^\s/>]+)([\s\S]*?)\/?>$/.exec(token)
    if (!open) throw new TypeError('Presentation XML is malformed')
    const rawName = open[1]!
    const rawAttrs = open[2]!
    const parentNamespaces = stack.at(-1)?.namespaces ?? new Map<string, string>()
    if (stack.length === 0 && ++roots > 1)
      throw new TypeError('Presentation XML must have exactly one root element')
    const namespaces = new Map(parentNamespaces)
    const declarations = new Set<string>()
    const attrs: string[] = []
    const attrPattern = /([^\s=]+)\s*=\s*(["'])([\s\S]*?)\2/g
    for (const match of rawAttrs.matchAll(attrPattern)) {
      const raw = match[1]!
      if (match[3]!.includes('<')) throw new TypeError('Presentation XML has malformed attributes')
      const value = decodeXml(match[3]!)
      if (raw === 'xmlns' || raw.startsWith('xmlns:')) {
        const prefix = raw === 'xmlns' ? '' : raw.slice(6)
        if (
          (prefix && !/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(prefix)) ||
          declarations.has(prefix) ||
          (prefix && value === '')
        )
          throw new TypeError('Presentation XML has invalid namespace declarations')
        declarations.add(prefix)
        if (
          prefix === 'xmlns' ||
          value === XMLNS_NAMESPACE ||
          (prefix === 'xml' && value !== XML_NAMESPACE) ||
          (prefix !== 'xml' && value === XML_NAMESPACE)
        )
          throw new TypeError('Presentation XML has an invalid namespace declaration')
        namespaces.set(prefix, value)
      } else attrs.push(raw)
    }
    if (rawAttrs.replace(attrPattern, '').trim())
      throw new TypeError('Presentation XML has malformed attributes')
    const name = qname(rawName, namespaces)
    const expandedAttrs = attrs.map((raw) => qname(raw, namespaces, true))
    if (new Set(expandedAttrs).size !== expandedAttrs.length)
      throw new TypeError('Presentation XML has duplicate attributes')
    const parent = stack.at(-1)
    const isCSld =
      name === `{${PRESENTATIONML_NAMESPACE}}cSld` &&
      (parent?.name === `{${PRESENTATIONML_NAMESPACE}}sld` ||
        parent?.name === `{${PRESENTATIONML_NAMESPACE}}sldLayout` ||
        parent?.name === `{${PRESENTATIONML_NAMESPACE}}sldMaster`)
    if (isCSld) {
      if (cSld) throw new TypeError('Presentation XML has duplicate cSld elements')
      cSld = {
        insertAt: offset,
        prefix: rawName.includes(':') ? rawName.slice(0, rawName.indexOf(':')) : '',
        namespaces,
      }
    }
    const directBackground =
      name === `{${PRESENTATIONML_NAMESPACE}}bg` && parent?.isSlideCSld === true
    if (directBackground) {
      backgroundPrefix = rawName.includes(':') ? rawName.slice(0, rawName.indexOf(':')) : ''
      if (cSld!.namespaces.get(backgroundPrefix) !== PRESENTATIONML_NAMESPACE) {
        backgroundDeclaration = backgroundPrefix
          ? ` xmlns:${backgroundPrefix}="${PRESENTATIONML_NAMESPACE}"`
          : ` xmlns="${PRESENTATIONML_NAMESPACE}"`
      }
    }
    if (directBackground && selfClosing) {
      if (backgroundRange) throw new TypeError('Presentation XML has duplicate slide backgrounds')
      backgroundRange = [start, offset]
    }
    if (!selfClosing)
      stack.push({ name, rawName, namespaces, start, directBackground, isSlideCSld: isCSld })
  }
  if (roots !== 1 || !cSld) throw new TypeError('Presentation XML has no slide cSld')
  const existingDrawingPrefix = [...cSld.namespaces.entries()].find(
    ([prefix, namespace]) => prefix !== '' && namespace === DRAWINGML_NAMESPACE,
  )?.[0]
  let drawingPrefix = existingDrawingPrefix
  let drawingDeclaration: string | undefined
  if (!drawingPrefix) {
    drawingPrefix = 'a'
    let suffix = 1
    while (cSld.namespaces.has(drawingPrefix)) drawingPrefix = `a${suffix++}`
    drawingDeclaration = ` xmlns:${drawingPrefix}="${DRAWINGML_NAMESPACE}"`
  }
  return {
    cSldInsertAt: cSld.insertAt,
    presentationPrefix: backgroundPrefix ?? cSld.prefix,
    ...(backgroundDeclaration ? { presentationDeclaration: backgroundDeclaration } : {}),
    drawingPrefix,
    ...(drawingDeclaration ? { drawingDeclaration } : {}),
    ...(backgroundRange ? { backgroundRange } : {}),
  }
}

interface Frame {
  readonly name: string
  readonly namespaces: ReadonlyMap<string, string>
  readonly preserveSpace: boolean
  readonly textBearing: boolean
  text: string
}

/** Namespace-aware semantic XML canonicalization for bounded OOXML parts. */
export function canonicalizePresentationXml(xml: string): Uint8Array {
  const inputBytes = new TextEncoder().encode(xml)
  if (inputBytes.byteLength > MAX_PRESENTATION_XML_BYTES)
    throw new TypeError('Presentation XML exceeds bounds')

  const tokens = xml.match(XML_TOKEN_PATTERN)
  if (!tokens) throw new TypeError('Presentation XML is malformed')
  if (tokens.join('') !== xml) throw new TypeError('Presentation XML is malformed')
  const stack: Frame[] = []
  const output: string[] = []
  const flushText = (frame: Frame | undefined) => {
    if (!frame) return
    const text = frame.text
    frame.text = ''
    if (text.trim() || frame.preserveSpace || frame.textBearing)
      output.push(`T${JSON.stringify(text)}`)
  }
  let nodes = 0
  let roots = 0

  for (const token of tokens) {
    if (token.startsWith('<!--') || token.startsWith('<?')) continue
    if (token.startsWith('<![CDATA[')) {
      const frame = stack.at(-1)
      if (!frame) throw new TypeError('Presentation XML is malformed')
      const text = token.slice(9, -3)
      validateXmlCharacters(text)
      frame.text += text
      continue
    }
    if (!token.startsWith('<')) {
      const text = decodeXml(token)
      const frame = stack.at(-1)
      if (!frame) {
        if (text.trim()) throw new TypeError('Presentation XML is malformed')
      } else frame.text += text
      continue
    }
    if (token.startsWith('<!')) throw new TypeError('Unsupported presentation XML declaration')
    if (token.startsWith('</')) {
      const rawName = /^<\/\s*([^\s>]+)\s*>$/.exec(token)?.[1]
      const frame = stack.pop()
      if (!rawName || !frame || qname(rawName, frame.namespaces) !== frame.name)
        throw new TypeError('Presentation XML is malformed')
      flushText(frame)
      output.push(`E${JSON.stringify(frame.name)}`)
      continue
    }

    nodes += 1
    if (nodes > MAX_XML_NODES || stack.length >= MAX_XML_DEPTH)
      throw new TypeError('Presentation XML exceeds bounds')
    const selfClosing = /\/\s*>$/.test(token)
    const open = /^<\s*([^\s/>]+)([\s\S]*?)\/?>$/.exec(token)
    if (!open) throw new TypeError('Presentation XML is malformed')
    const rawName = open[1]!
    const rawAttrs = open[2]!
    flushText(stack.at(-1))
    const parentNamespaces = stack.at(-1)?.namespaces ?? new Map<string, string>()
    if (stack.length === 0 && ++roots > 1)
      throw new TypeError('Presentation XML must have exactly one root element')
    const namespaces = new Map(parentNamespaces)
    const attrs: Array<{ raw: string; value: string }> = []
    const declarations = new Set<string>()
    const attrPattern = /([^\s=]+)\s*=\s*(["'])([\s\S]*?)\2/g
    for (const match of rawAttrs.matchAll(attrPattern)) {
      const raw = match[1]!
      if (match[3]!.includes('<')) throw new TypeError('Presentation XML has malformed attributes')
      const value = decodeXml(match[3]!)
      if (raw === 'xmlns' || raw.startsWith('xmlns:')) {
        const prefix = raw === 'xmlns' ? '' : raw.slice(6)
        if (declarations.has(prefix))
          throw new TypeError('Presentation XML has duplicate namespace declarations')
        declarations.add(prefix)
        if (
          prefix === 'xmlns' ||
          value === XMLNS_NAMESPACE ||
          (prefix === 'xml' && value !== XML_NAMESPACE) ||
          (prefix !== 'xml' && value === XML_NAMESPACE)
        )
          throw new TypeError('Presentation XML has an invalid namespace declaration')
        namespaces.set(prefix, value)
      } else attrs.push({ raw, value })
    }
    if (rawAttrs.replace(attrPattern, '').trim())
      throw new TypeError('Presentation XML has malformed attributes')

    const name = qname(rawName, namespaces)
    const canonicalAttrs = attrs
      .map(({ raw, value }) => ({ name: qname(raw, namespaces, true), value }))
      .sort((left, right) =>
        left.name === right.name
          ? left.value.localeCompare(right.value)
          : left.name.localeCompare(right.name),
      )
    if (
      canonicalAttrs.some(
        (attr, index) => index > 0 && attr.name === canonicalAttrs[index - 1]!.name,
      )
    )
      throw new TypeError('Presentation XML has duplicate attributes')
    const preserveAttr = canonicalAttrs.find((attr) => attr.name === `{${XML_NAMESPACE}}space`)
    const preserveSpace = preserveAttr
      ? preserveAttr.value === 'preserve'
      : (stack.at(-1)?.preserveSpace ?? false)
    output.push(
      `S${JSON.stringify(name)}[${canonicalAttrs
        .map((attr) => `${JSON.stringify(attr.name)}:${JSON.stringify(attr.value)}`)
        .join(',')}]`,
    )
    if (selfClosing) output.push(`E${JSON.stringify(name)}`)
    else
      stack.push({
        name,
        namespaces,
        preserveSpace,
        textBearing: /}(?:t|v|f)$/.test(name),
        text: '',
      })
  }
  if (stack.length) throw new TypeError('Presentation XML is malformed')
  if (roots !== 1) throw new TypeError('Presentation XML must have exactly one root element')
  return new TextEncoder().encode(output.join(''))
}
