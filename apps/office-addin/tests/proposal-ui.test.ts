import { describe, expect, it, vi } from 'vitest'
import { proposalPresentation, safeUploadError, uploadSessionFile } from '../src/App.js'
import type { OfficeHostRuntime } from '../src/agent/host-runtime.js'

describe('generic proposal presentation', () => {
  it('presents structured impact as a readable comparison without protocol JSON or code', () => {
    expect(
      proposalPresentation({
        id: 'p1',
        operation: 'edit_slide_xml',
        title: 'Edit XML',
        impact: { host: 'powerpoint', targets: ['slide-1'], count: 1 },
        preview: { nodes: 2 },
        fingerprint: 'fp',
        before: '<old/>',
        after: '<new/>',
        code: 'sync()',
      }),
    ).toEqual({
      title: 'Edit XML',
      host: 'PowerPoint',
      count: 1,
      targets: ['Slide 1'],
      before: '<old/>',
      after: '<new/>',
      preview: '',
      code: undefined,
    })
  })

  it('turns a preview-only structured proposal into readable copy', () => {
    expect(
      proposalPresentation({
        id: 'p3',
        operation: 'write_document',
        title: 'Append text',
        impact: { host: 'word', targets: ['document:end'], count: 1 },
        preview: { mode: 'append', text: 'New paragraph' },
        fingerprint: 'fp',
        code: '{"version":1}',
      }),
    ).toMatchObject({
      host: 'Word',
      targets: ['End of document'],
      preview: 'Mode: append\nText: New paragraph',
      code: undefined,
    })
  })

  it('keeps the actionable preview when only an internal before snapshot exists', () => {
    expect(
      proposalPresentation({
        id: 'p4',
        operation: 'duplicate_slide',
        title: 'Duplicate slide',
        impact: { host: 'powerpoint', targets: ['slide-1'], count: 1 },
        preview: { slideIndex: 3, slideId: 'slide-1' },
        fingerprint: 'fp',
        before: { fingerprint: 'internal-hash', slideId: 'slide-1' },
      }),
    ).toMatchObject({
      preview: 'Slide index: 3\nSlide ID: Slide 1',
      before: '',
      after: '',
    })
  })

  it('does not expose structured snapshots, fingerprints, hashes, or raw operations', () => {
    const presentation = proposalPresentation({
      id: 'p5',
      operation: 'execute_office_program',
      title: 'Update cells',
      impact: { host: 'Excel', targets: ['/worksheets/Q1/range/A1:B2'], count: 4 },
      preview: {
        rangeAddress: 'A1:B2',
        rowCount: 2,
        fingerprint: 'secret-fingerprint',
        payloadHash: 'secret-hash',
        operation: 'raw-operation',
        nested: { private: 'value' },
      },
      fingerprint: 'top-secret',
      before: { fingerprint: 'before-secret', values: [['old']] },
      after: { hash: 'after-secret', values: [['new']] },
    })

    expect(presentation).toMatchObject({
      host: 'Excel',
      targets: ['Worksheets · Q1 · Range · A1:B2'],
      before: '',
      after: '',
      preview: 'Range address: A1:B2\nRow count: 2',
    })
    expect(JSON.stringify(presentation)).not.toMatch(
      /secret|raw-operation|nested|fingerprint|hash/i,
    )
  })

  it.each(['WORD', 'Microsoft Word', 'excel', 'Excel', 'POWERPOINT', 'Microsoft PowerPoint'])(
    'normalizes the real Office host spelling %s',
    (host) => {
      const presentation = proposalPresentation({
        id: 'host',
        operation: 'write',
        title: 'Write',
        impact: { host, targets: ['selection'], count: 1 },
        preview: { text: 'Hello' },
        fingerprint: 'fp',
      })
      expect(['Word', 'Excel', 'PowerPoint']).toContain(presentation.host)
    },
  )

  it('keeps the legacy append preview compatible', () => {
    expect(
      proposalPresentation({
        id: 'p2',
        operation: 'append',
        before: 'old',
        value: ' new',
        fingerprint: 'fp',
      }),
    ).toMatchObject({ title: 'Append to selection', before: 'old', after: 'old new' })
  })

  it('maps upload failures to stable UI-safe errors', () => {
    expect(safeUploadError(new Error('invalid_skill_package'))).toBe('invalid_skill_package')
    expect(safeUploadError(new Error('/Users/alice/private'))).toBe('upload_failed')
  })

  it.each([
    ['large.bin', 2 * 1024 * 1024 + 1, 'vfs_limit'],
    ['SKILL.md', 64 * 1024 + 1, 'invalid_skill_package'],
  ] as const)('rejects %s by File.size before loading its content', async (name, size, code) => {
    const arrayBuffer = vi.fn()
    const text = vi.fn()
    const runtime = { uploadFile: vi.fn(), installSkill: vi.fn() } as unknown as OfficeHostRuntime
    await expect(uploadSessionFile(runtime, { name, size, arrayBuffer, text })).rejects.toThrow(
      code,
    )
    expect(arrayBuffer).not.toHaveBeenCalled()
    expect(text).not.toHaveBeenCalled()
    expect(runtime.uploadFile).not.toHaveBeenCalled()
    expect(runtime.installSkill).not.toHaveBeenCalled()
  })
})
