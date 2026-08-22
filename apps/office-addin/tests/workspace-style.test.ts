import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const styles = readFileSync(new URL('../src/styles.css', import.meta.url), 'utf8')
const main = readFileSync(new URL('../src/main.tsx', import.meta.url), 'utf8')

describe('Office workspace styling', () => {
  it('imports shared WisWork tokens and uses no raw UI colors', () => {
    expect(main).toContain('@wiswork/ui/tokens.css')
    expect(styles).not.toMatch(/#[0-9a-f]{3,8}\b/i)
    expect(styles).not.toMatch(/\brgba?\(/i)
  })

  it('supports narrow task panes, high contrast, dark tokens, and reduced motion', () => {
    expect(styles).toContain('min-width: 280px')
    expect(styles).toMatch(/@media\s*\(max-width:\s*500px\)/)
    expect(styles).toMatch(/forced-colors:\s*active/)
    expect(styles).toContain("[data-theme='dark']")
    expect(styles).toMatch(/prefers-reduced-motion:\s*reduce/)
    expect(styles).toMatch(/html,[\s\S]*#root\s*\{[\s\S]*height:\s*100%/)
    expect(styles).toMatch(/\.agent-workspace\s*\{[\s\S]*height:\s*100%/)
    expect(styles).toMatch(/\.agent-workspace\s*\{[\s\S]*overflow:\s*hidden/)
    expect(styles).toMatch(/\.agent-timeline\s*\{[\s\S]*overflow-y:\s*auto/)
    expect(styles).toMatch(/\.agent-timeline\s*\{[\s\S]*overscroll-behavior:\s*contain/)
    expect(styles).toMatch(/\.agent-timeline\s*\{[\s\S]*touch-action:\s*pan-y/)
  })
})
