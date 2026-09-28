import { describe, expect, it } from 'vitest'
import { parseFileToText } from '../src/index'
import { writeFixture } from './helpers/fixtures'

describe('parseFileToText: plain-text formats', () => {
  const cases: Array<[string, string]> = [
    ['sample.txt', 'plain text hello'],
    ['sample.md', '# Title\n\nBody paragraph'],
    ['sample.csv', 'a,b,c\n1,2,3'],
    ['sample.tsv', 'a\tb\tc\n1\t2\t3'],
    ['sample.json', '{"key":"value"}'],
    ['sample.xml', '<root><item>value</item></root>'],
  ]

  for (const [name, content] of cases) {
    it(`reads ${name} verbatim`, async () => {
      const path = writeFixture(name, content)
      const result = await parseFileToText(path)
      expect(result).toEqual({ ok: true, kind: 'text', text: content })
    })
  }

  it.each(['html', 'htm'])('extracts visible webpage text from .%s', async (ext) => {
    const path = writeFixture(
      `page.${ext}`,
      '<!doctype html><html><head><title>Quarterly report</title><style>.hidden{display:none}</style></head><body><h1>Results &amp; outlook</h1><p>Revenue <strong>grew</strong>.</p><script>secretToken()</script><p hidden>Hidden claim</p><p aria-hidden="true">Decorative label</p><nav>Menu</nav><main><p>Operating profit rose.</p></main></body></html>',
    )
    const result = await parseFileToText(path)
    expect(result).toEqual({
      ok: true,
      kind: 'text',
      text: 'Results & outlook\nRevenue grew.\nMenu\nOperating profit rose.',
    })
  })

  it('reads a saved GBK webpage using its declared charset', async () => {
    const bytes = Buffer.concat([
      Buffer.from('<html><head><meta charset="gbk"></head><body><p>'),
      Buffer.from('d6d0cec4', 'hex'),
      Buffer.from('</p></body></html>'),
    ])
    const result = await parseFileToText(writeFixture('chinese.html', bytes))
    expect(result).toEqual({ ok: true, kind: 'text', text: '中文' })
  })

  it('is case-insensitive on the extension', async () => {
    const path = writeFixture('UPPER.TXT', 'upper')
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.text).toBe('upper')
  })

  it('fails gracefully on a missing file', async () => {
    const result = await parseFileToText('/nonexistent/nowhere.txt')
    expect(result.ok).toBe(false)
    expect(result.error).toBeTruthy()
  })
})

describe('parseFileToText: images and unsupported', () => {
  it('flags png as image without extracting text', async () => {
    const path = writeFixture('pic.png', Buffer.from('89504e47', 'hex'))
    const result = await parseFileToText(path)
    expect(result).toEqual({ ok: true, kind: 'image', mime: 'image/png' })
  })

  it.each([
    ['pic.jpg', 'image/jpeg'],
    ['pic.jpeg', 'image/jpeg'],
    ['pic.gif', 'image/gif'],
    ['pic.webp', 'image/webp'],
  ])('maps %s to mime %s', async (name, mime) => {
    const path = writeFixture(name, Buffer.from([0]))
    const result = await parseFileToText(path)
    expect(result.kind).toBe('image')
    expect(result.mime).toBe(mime)
  })

  it('rejects unknown extensions as unsupported', async () => {
    const path = writeFixture('archive.zip', Buffer.from([0x50, 0x4b]))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(false)
    expect(result.kind).toBe('unsupported')
    expect(result.error).toContain('.zip')
  })
})
