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

  it('fails closed when XML exceeds structural bounds', () => {
    const oversized = `<r>${'<n/>'.repeat(20_001)}</r>`
    expect(() => canonicalizePresentationXml(oversized)).toThrow('bounds')
  })
})
