const MAX_XML_BYTES = 2 * 1024 * 1024
const MAX_XML_NODES = 20_000
const MAX_XML_DEPTH = 128
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace'

const decodeXml = (value: string): string =>
  value
    .replace(/&#x([0-9a-f]+);/gi, (_match, hex: string) =>
      String.fromCodePoint(Number.parseInt(hex, 16)),
    )
    .replace(/&#(\d+);/g, (_match, decimal: string) =>
      String.fromCodePoint(Number.parseInt(decimal, 10)),
    )
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')

const qname = (raw: string, namespaces: ReadonlyMap<string, string>, attribute = false): string => {
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

interface Frame {
  readonly name: string
  readonly namespaces: ReadonlyMap<string, string>
  readonly preserveSpace: boolean
  readonly textBearing: boolean
}

/** Namespace-aware semantic XML canonicalization for bounded OOXML parts. */
export function canonicalizePresentationXml(xml: string): Uint8Array {
  const inputBytes = new TextEncoder().encode(xml)
  if (inputBytes.byteLength > MAX_XML_BYTES) throw new TypeError('Presentation XML exceeds bounds')

  const tokens = xml.match(
    /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<(?:[^< >"']|\s|"[^"]*"|'[^']*')*>|[^<]+/g,
  )
  if (!tokens) throw new TypeError('Presentation XML is malformed')
  if (tokens.join('') !== xml) throw new TypeError('Presentation XML is malformed')
  const stack: Frame[] = []
  const output: string[] = []
  let nodes = 0

  for (const token of tokens) {
    if (token.startsWith('<!--') || token.startsWith('<?')) continue
    if (token.startsWith('<![CDATA[')) {
      const text = token.slice(9, -3)
      if (text || stack.at(-1)?.preserveSpace || stack.at(-1)?.textBearing)
        output.push(`T${JSON.stringify(text)}`)
      continue
    }
    if (!token.startsWith('<')) {
      const text = decodeXml(token)
      if (text.trim() || stack.at(-1)?.preserveSpace || stack.at(-1)?.textBearing)
        output.push(`T${JSON.stringify(text)}`)
      continue
    }
    if (token.startsWith('<!')) throw new TypeError('Unsupported presentation XML declaration')
    if (token.startsWith('</')) {
      const rawName = /^<\/\s*([^\s>]+)\s*>$/.exec(token)?.[1]
      const frame = stack.pop()
      if (!rawName || !frame || qname(rawName, frame.namespaces) !== frame.name)
        throw new TypeError('Presentation XML is malformed')
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
    const parentNamespaces = stack.at(-1)?.namespaces ?? new Map<string, string>()
    const namespaces = new Map(parentNamespaces)
    const attrs: Array<{ raw: string; value: string }> = []
    const attrPattern = /([^\s=]+)\s*=\s*(["'])([\s\S]*?)\2/g
    for (const match of rawAttrs.matchAll(attrPattern)) {
      const raw = match[1]!
      const value = decodeXml(match[3]!)
      if (raw === 'xmlns') namespaces.set('', value)
      else if (raw.startsWith('xmlns:')) namespaces.set(raw.slice(6), value)
      else attrs.push({ raw, value })
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
      })
  }
  if (stack.length) throw new TypeError('Presentation XML is malformed')
  return new TextEncoder().encode(output.join(''))
}
