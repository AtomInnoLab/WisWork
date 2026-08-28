import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  openPptx,
  savePptx,
  addElement,
  deleteElement,
  deleteElementWithCleanup,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const OFF = { x: 914400, y: 914400, cx: 3657600, cy: 914400 }

describe('add/delete element', () => {
  it('added textbox survives save → reopen with text and geometry', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const before = slide.elements.length
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { ...OFF },
      paragraphs: [{ runs: [{ text: 'Newly inserted textbox', bold: true }] }],
    })
    expect(el.type).toBe('text')
    expect(slide.elements.length).toBe(before + 1)

    const reopened = await openPptx(await savePptx(opened))
    const slide2 = reopened.deck.slides[0]!
    expect(slide2.elements.length).toBe(before + 1)
    const el2: any = slide2.elements[slide2.elements.length - 1]
    expect(el2.type).toBe('text')
    expect(el2.transform.offset).toEqual(OFF)
    const text = el2.text.paragraphs
      .flatMap((p: any) => p.runs)
      .map((r: any) => r.text)
      .join('')
    expect(text).toBe('Newly inserted textbox')
    expect(el2.text.paragraphs[0].runs[0].bold).toBe(true)
  })

  it('added shape round-trips preset geometry and solid fill', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    addElement(slide, { kind: 'ellipse', offset: { ...OFF }, fillColor: '#C43E1C' })

    const reopened = await openPptx(await savePptx(opened))
    const el2: any = reopened.deck.slides[0]!.elements.at(-1)
    expect(el2.type).toBe('shape')
    expect(el2.presetGeometry).toBe('ellipse')
    expect(el2.fill).toEqual({ type: 'solid', color: '#C43E1C' })
  })

  it('run fontFamily writes both latin and ea slots and round-trips', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { ...OFF },
      paragraphs: [{ runs: [{ text: '中文 Latin', fontFamily: '微软雅黑' }] }],
    })
    expect(el.anchor.originalXml).toContain('<a:latin typeface="微软雅黑"/>')
    expect(el.anchor.originalXml).toContain('<a:ea typeface="微软雅黑"/>')

    const reopened = await openPptx(await savePptx(opened))
    const el2: any = reopened.deck.slides[0]!.elements.at(-1)
    expect(el2.text.paragraphs[0].runs[0].fontFamily).toBe('微软雅黑')
  })

  it('cNvPr ids stay unique after two inserts', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const a = addElement(slide, { kind: 'rect', offset: { ...OFF } })
    const b = addElement(slide, { kind: 'rect', offset: { ...OFF } })
    const idOf = (xml: string) => /<p:cNvPr\s[^>]*\bid="(\d+)"/.exec(xml)![1]
    expect(idOf(a.anchor.originalXml)).not.toBe(idOf(b.anchor.originalXml))
  })

  it('delete element persists through save → reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const before = slide.elements.length
    expect(before).toBeGreaterThan(1)
    const victim = slide.elements[0]!
    expect(deleteElement(slide, victim.id)).toBe(true)

    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.slides[0]!.elements.length).toBe(before - 1)
  })

  it('delete-only edit still marks the deck dirty for save', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    deleteElement(slide, slide.elements[0]!.id)
    expect(slide.structureDirty).toBe(true)
  })

  it('authoritative delete detaches connectors and removes only exclusive slide relationships', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const mediaBefore = new Map(
      [...opened.archive.entries].filter(([path]) => path.startsWith('ppt/media/')),
    )
    expect(mediaBefore.size).toBeGreaterThan(0)
    const victim = addElement(slide, { kind: 'rect', offset: { ...OFF } })
    const survivor = addElement(slide, { kind: 'rect', offset: { ...OFF } })
    const connector = addElement(slide, { kind: 'line', offset: { ...OFF } })
    const victimNvId = /<p:cNvPr\s[^>]*\bid="(\d+)"/.exec(victim.anchor.originalXml)![1]
    victim.anchor.originalXml = victim.anchor.originalXml.replace(
      '</p:cNvPr>',
      '<a:hlinkClick r:id="rIdExclusive"/></p:cNvPr>',
    )
    survivor.anchor.originalXml = survivor.anchor.originalXml.replace(
      '</p:cNvPr>',
      '<a:hlinkClick r:id="rIdShared"/></p:cNvPr>',
    )
    connector.anchor.originalXml = connector.anchor.originalXml.replace(
      '<p:cNvCxnSpPr/>',
      `<p:cNvCxnSpPr><a:stCxn id="${victimNvId}" idx="0"/><a:endCxn id="${victimNvId}" idx="2"/></p:cNvCxnSpPr>`,
    )
    const relsPath = 'ppt/slides/_rels/slide1.xml.rels'
    const rels = opened.archive.readText(relsPath)!
    opened.archive.entries.set(
      relsPath,
      Buffer.from(
        rels.replace(
          '</Relationships>',
          '<Relationship Id="rIdExclusive" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://exclusive.invalid" TargetMode="External"/>' +
            '<Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://shared.invalid" TargetMode="External"/>' +
            '</Relationships>',
        ),
      ),
    )

    expect(deleteElementWithCleanup(opened, slide, victim.id)).toBe(true)
    const reopened = await openPptx(await savePptx(opened))
    const xml = reopened.archive.readText(slide.path)!
    const reopenedRels = reopened.archive.readText(relsPath)!
    expect(xml).not.toMatch(new RegExp(`<a:(?:stCxn|endCxn)[^>]*id="${victimNvId}"`))
    expect(reopenedRels).not.toContain('Id="rIdExclusive"')
    expect(reopenedRels).toContain('Id="rIdShared"')
    for (const [path, bytes] of mediaBefore) expect(reopened.archive.readBytes(path)).toEqual(bytes)
    expect(reopened.deck.slides[0]!.elements).toHaveLength(slide.elements.length)
  })

  it('fails closed when unsupported timing XML still targets the deleted shape', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const victim = addElement(slide, { kind: 'rect', offset: { ...OFF } })
    const victimNvId = /<p:cNvPr\s[^>]*\bid="(\d+)"/.exec(victim.anchor.originalXml)![1]
    slide.bodySuffix += `<p:timing><p:spTgt spid="${victimNvId}"/></p:timing>`
    const count = slide.elements.length
    expect(deleteElementWithCleanup(opened, slide, victim.id)).toBe(false)
    expect(slide.elements).toHaveLength(count)
  })
})
