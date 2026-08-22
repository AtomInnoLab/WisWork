import React from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { Markdown } from '@wiswork/ui'

function markup(text: string): string {
  return renderToStaticMarkup(React.createElement(Markdown, { text }))
}

describe('safe shared Markdown', () => {
  it('renders the supported chat subset', () => {
    const html = markup(
      '# Heading\n\nParagraph with **bold**, *italic*, and `inline code`.\n\n- first\n- second\n\n1. one\n2. two',
    )
    expect(html).toContain('<p class="ai-md-h">Heading</p>')
    expect(html).toContain('<strong>bold</strong>')
    expect(html).toContain('<em>italic</em>')
    expect(html).toContain('<code>inline code</code>')
    expect(html).toContain('<ul><li>first</li><li>second</li></ul>')
    expect(html).toContain('<ol><li>one</li><li>two</li></ol>')
  })

  it('escapes raw HTML and never creates link or image authority', () => {
    const text =
      '<img src="https://example.invalid/raw.png" onerror="alert(1)"> **raw**\n' +
      '[**link**](https://example.invalid/page)\n' +
      '![**image**](https://example.invalid/image.png)'
    const html = markup(text)
    expect(html).toContain('&lt;img src=&quot;https://example.invalid/raw.png&quot;')
    expect(html).toContain('**raw**')
    expect(html).toContain('[**link**](https://example.invalid/page)')
    expect(html).toContain('![**image**](https://example.invalid/image.png)')
    expect(html).not.toContain('<a')
    expect(html).not.toContain('<img')
    expect(html).not.toContain('<strong>')
  })

  it('keeps fenced HTML and incomplete streaming syntax literal', () => {
    const html = markup('```html\n<script>bad()</script>\n```\n\n**partial and *still partial')
    expect(html).toContain('```html')
    expect(html).toContain('&lt;script&gt;bad()&lt;/script&gt;')
    expect(html).toContain('**partial and *still partial')
    expect(html).not.toContain('<script>')
    expect(html).not.toContain('<strong>')
    expect(html).not.toContain('<em>')
    expect(html).not.toContain('<code>')
  })
})
