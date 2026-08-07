import { readFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import { registerLatexProtocolScheme } from '../src/main/latex-protocol-scheme'

describe('LaTeX PDF privileged scheme', () => {
  it('registers the secure standard fetchable scheme once per protocol owner', () => {
    const protocol = { registerSchemesAsPrivileged: vi.fn() }
    registerLatexProtocolScheme(protocol)
    registerLatexProtocolScheme(protocol)
    expect(protocol.registerSchemesAsPrivileged).toHaveBeenCalledOnce()
    expect(protocol.registerSchemesAsPrivileged).toHaveBeenCalledWith([
      {
        scheme: 'wiswork-latex-pdf',
        privileges: {
          secure: true,
          standard: true,
          supportFetchAPI: true,
          corsEnabled: true,
        },
      },
    ])
  })

  it('is called at module top level before app readiness while ready only installs the handler', () => {
    const source = readFileSync(new URL('../src/main/index.ts', import.meta.url), 'utf8')
    expect(source.indexOf('registerLatexProtocolScheme(protocol)')).toBeGreaterThan(-1)
    expect(source.indexOf('registerLatexProtocolScheme(protocol)')).toBeLessThan(
      source.indexOf('app.whenReady()'),
    )
    expect(source.indexOf('registerLatexPdfProtocol(protocol)')).toBeGreaterThan(
      source.indexOf('app.whenReady()'),
    )
  })
})
