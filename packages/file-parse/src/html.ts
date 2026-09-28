import { parse } from 'parse5'
import type { DefaultTreeAdapterMap } from 'parse5'

type Node = DefaultTreeAdapterMap['node']

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
