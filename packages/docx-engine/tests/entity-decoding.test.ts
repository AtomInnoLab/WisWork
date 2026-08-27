import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

describe('entity decoding in parsed run text', () => {
  const STORED = 'Escape &amp;amp; as &amp;lt;b&amp;gt; ok'
  const DISPLAYED = 'Escape &amp; as &lt;b&gt; ok'

  it('keeps text that is literally an entity reference', async () => {
    const bytes = await buildDocx({
      bodyXml: `<w:p><w:r><w:t xml:space="preserve">${STORED}</w:t></w:r></w:p>`,
    })

    const doc = await parseDocx(bytes)

    expect(doc.blocks[0].runs!.map((run) => run.text).join('')).toBe(DISPLAYED)
  })

  it('still resolves numeric character references beside literal entity text', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t xml:space="preserve">&amp;lt; then &#x2713;</w:t></w:r></w:p>',
    })

    const doc = await parseDocx(bytes)

    expect(doc.blocks[0].runs!.map((run) => run.text).join('')).toBe('&lt; then ✓')
  })

  it('does not rewrite an untouched literal entity when a sibling run is edited', async () => {
    const bytes = await buildDocx({
      bodyXml:
        `<w:p><w:r><w:t xml:space="preserve">${STORED} </w:t></w:r>` +
        '<w:r><w:rPr><w:b/></w:rPr><w:t>tail</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    const runs = doc.blocks[0].runs!.map((run) => (run.bold ? { ...run, text: 'edited' } : run))

    const saved = await saveDocx(doc, [{ kind: 'generated', block: { type: 'paragraph', runs } }])
    const zip = await JSZip.loadAsync(saved)
    const xml = await zip.file('word/document.xml')!.async('string')

    expect(xml).toContain(`${STORED} `)
    expect(xml).toContain('edited')
  })
})
