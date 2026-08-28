import { describe, expect, it } from 'vitest'
import { canonicalizePresentationXml } from '../src/presentation-xml'

describe('presentation XML semantic canonicalization', () => {
  it('ignores namespace prefixes, attribute ordering, and formatting-only whitespace', () => {
    const left =
      '<a:root xmlns:a="urn:test" a:z="2" x="1"><a:child foo="bar">value</a:child></a:root>'
    const right = `
      <q:root x="1" q:z="2" xmlns:q="urn:test">
        <q:child foo="bar">value</q:child>
      </q:root>`
    expect(canonicalizePresentationXml(left)).toEqual(canonicalizePresentationXml(right))
  })

  it('preserves text and xml:space semantics', () => {
    const plain = '<a:t xmlns:a="urn:test">value</a:t>'
    const spaced = '<x:t xmlns:x="urn:test" xml:space="preserve"> value </x:t>'
    expect(canonicalizePresentationXml(plain)).not.toEqual(canonicalizePresentationXml(spaced))
  })

  it('merges adjacent character content across comments, CDATA, and entity tokens', () => {
    const split = '<a:t xmlns:a="urn:test">a<!-- ignored --><![CDATA[b]]>&#99;</a:t>'
    const plain = '<q:t xmlns:q="urn:test">abc</q:t>'
    expect(canonicalizePresentationXml(split)).toEqual(canonicalizePresentationXml(plain))
  })

  it('preserves mixed-content child order and xml:space content', () => {
    expect(canonicalizePresentationXml('<r>a<c/>b</r>')).not.toEqual(
      canonicalizePresentationXml('<r>ab<c/></r>'),
    )
    expect(canonicalizePresentationXml('<t xml:space="preserve">a <!--x--> b</t>')).toEqual(
      canonicalizePresentationXml('<t xml:space="preserve"><![CDATA[a  b]]></t>'),
    )
  })

  it.each([
    '<r xmlns="a" xmlns="b"/>',
    '<r xmlns:x="a" xmlns:x="b"/>',
    '<r>&unknown;</r>',
    '<r>&#0;</r>',
    '<r>&#xD800;</r>',
    '<r>\u0001</r>',
    '<!DOCTYPE r [<!ENTITY x "value">]><r>&x;</r>',
    '<r xmlns:x="http://www.w3.org/XML/1998/namespace"/>',
    '<r xmlns="http://www.w3.org/XML/1998/namespace"/>',
    '<r xmlns:x="http://www.w3.org/2000/xmlns/"/>',
    '<r xmlns="http://www.w3.org/2000/xmlns/"/>',
    '<r xmlns:xml="urn:not-xml"/>',
    '',
    '  <!-- no root --> <?pi ok?>',
    '<a/><b/>',
  ])('rejects invalid or unsafe XML: %s', (xml) => {
    expect(() => canonicalizePresentationXml(xml)).toThrow()
  })

  it('accepts one root surrounded only by whitespace, comments, and processing instructions', () => {
    expect(() =>
      canonicalizePresentationXml(' <!--before--><?pi ok?><r/><!--after--> '),
    ).not.toThrow()
  })

  it('fails closed when XML exceeds structural bounds', () => {
    const oversized = `<r>${'<n/>'.repeat(20_001)}</r>`
    expect(() => canonicalizePresentationXml(oversized)).toThrow('bounds')
  })
})
