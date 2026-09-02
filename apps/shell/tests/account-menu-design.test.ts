import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { officeConnectionCopy } from '../src/renderer/src/Home'

const root = resolve(__dirname, '../../..')
const read = (path: string) => readFileSync(resolve(root, path), 'utf8')

describe('account menu information hierarchy', () => {
  it('uses localized, concise Office connection copy', () => {
    expect(officeConnectionCopy('zh', 'disconnected')).toEqual({
      label: '连接 Office',
      detail: '输入 Office 中显示的 6 位代码',
      action: '连接',
      connected: false,
    })
    expect(officeConnectionCopy('zh', 'paired').detail).toBe('已连接')
    expect(officeConnectionCopy('en', 'disconnected').label).toBe('Connect Office')
  })

  it('keeps implementation diagnostics out of the primary account menu', () => {
    const home = read('apps/shell/src/renderer/src/Home.tsx')
    expect(home).not.toContain('enhanced-mode-diagnostics')
    expect(home).not.toContain('OpenAI GitHub Release')
    expect(home).not.toContain('Office relay:')
    expect(home).toContain('office-connection-row')
    expect(home).toContain('office-connection-panel')
  })
})
