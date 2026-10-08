import { parse } from 'parse5'
import type { DefaultTreeAdapterMap } from 'parse5'

type Node = DefaultTreeAdapterMap['node']

/** Decode saved page bytes using an HTTP charset, HTML meta declaration, or UTF-8. */
export function decodeHtmlBytes(bytes: Uint8Array, contentType?: string): string {
  const httpCharset = /(?:^|;)\s*charset\s*=\s*(?:"([^"]+)"|'([^']+)'|([^;\s]+))/i.exec(
    contentType ?? '',
  )
  const prefix = Buffer.from(bytes.subarray(0, 1024)).toString('latin1')
  const metaCharset = /<meta\b[^>]*\bcharset\s*=\s*(?:"([^"]+)"|'([^']+)'|([^\s/>;]+))/i.exec(
    prefix,
  )
  const encoding =
    bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf
      ? 'utf-8'
      : bytes[0] === 0xff && bytes[1] === 0xfe
        ? 'utf-16le'
        : bytes[0] === 0xfe && bytes[1] === 0xff
          ? 'utf-16be'
          : (httpCharset?.[1] ??
            httpCharset?.[2] ??
            httpCharset?.[3] ??
            metaCharset?.[1] ??
            metaCharset?.[2] ??
            metaCharset?.[3] ??
            'utf-8')
  return new TextDecoder(encoding, { fatal: true }).decode(bytes)
}

const SKIP = new Set([
  'head',
  'script',
  'style',
  'template',
  'noscript',
  'svg',
  'canvas',
  'iframe',
  'form',
])
const BREAK = new Set([
  'address',
  'article',
  'blockquote',
  'br',
  'dd',
  'div',
  'dl',
  'dt',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'header',
  'footer',
  'li',
  'main',
  'nav',
  'ol',
  'p',
  'pre',
  'section',
  'table',
  'tr',
  'ul',
])

/** Extract page body text from saved HTML. No remote resources are loaded. */
export function htmlToText(html: string): string {
  const parts: string[] = []
  const visit = (node: Node): void => {
    if ('value' in node) {
      parts.push(node.value)
      return
    }
    if (
      'tagName' in node &&
      (SKIP.has(node.tagName) ||
        node.attrs.some(
          (attr) =>
            attr.name === 'hidden' || (attr.name === 'aria-hidden' && attr.value === 'true'),
        ))
    )
      return
    const boundary = 'tagName' in node && BREAK.has(node.tagName)
    if (boundary) parts.push('\n')
    if ('childNodes' in node) for (const child of node.childNodes) visit(child)
    if (boundary) parts.push('\n')
  }
  visit(parse(html))
  return parts
    .join('')
    .replace(/\u00a0/g, ' ')
    .replace(/[\t\r ]+/g, ' ')
    .replace(/ *\n */g, '\n')
    .replace(/\n+/g, '\n')
    .trim()
}
