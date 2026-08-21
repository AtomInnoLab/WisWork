import { afterEach, describe, expect, it, vi } from 'vitest'
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { createExcelImportMediaSkill } from '../src/skills/excel/excel-import-media.js'
import { createPowerPointImportMediaSkill } from '../src/skills/powerpoint/powerpoint-import-media.js'
import {
  exportSafeCsv,
  readBoundedCsv,
  readBoundedImage,
} from '../src/skills/shared/import-media.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'

const png = (width = 1, height = 1) => {
  const bytes = new Uint8Array(24)
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10])
  new DataView(bytes.buffer).setUint32(16, width)
  new DataView(bytes.buffer).setUint32(20, height)
  return bytes
}
const call = (name: string, input: Record<string, unknown>) => ({ id: 'c1', name, input })

afterEach(() => {
  delete (globalThis as Record<string, unknown>).Office
  delete (globalThis as Record<string, unknown>).Excel
  delete (globalThis as Record<string, unknown>).PowerPoint
})

describe('host capability advertisement', () => {
  it('keeps import/media tools hidden without the complete API set', () => {
    expect(createOfficeHostRuntime('excel').skill.tools.map((item) => item.name)).not.toContain(
      'csv-to-sheet',
    )
    expect(
      createOfficeHostRuntime('powerpoint').skill.tools.map((item) => item.name),
    ).not.toContain('insert-image')
  })

  it('advertises only the matching host tools after API-set negotiation', () => {
    ;(globalThis as Record<string, unknown>).Office = {
      context: {
        host: 'Excel',
        requirements: {
          isSetSupported: (name: string, version: string) =>
            name === 'ExcelApi' && version === '1.9',
        },
      },
    }
    ;(globalThis as Record<string, unknown>).Excel = { run: vi.fn() }
    const excel = createOfficeHostRuntime('excel').skill.tools.map((item) => item.name)
    expect(excel).toEqual(
      expect.arrayContaining(['csv-to-sheet', 'sheet-to-csv', 'image-to-sheet']),
    )
    expect(excel).not.toContain('insert-image')

    ;(globalThis as Record<string, any>).Office.context.host = 'PowerPoint'
    ;(globalThis as Record<string, any>).Office.context.requirements.isSetSupported = (
      name: string,
      version: string,
    ) => name === 'PowerPointApi' && version === '1.8'
    ;(globalThis as Record<string, unknown>).PowerPoint = { run: vi.fn() }
    const powerpoint = createOfficeHostRuntime('powerpoint').skill.tools.map((item) => item.name)
    expect(powerpoint).toContain('insert-image')
    expect(powerpoint).not.toContain('csv-to-sheet')
  })
})

describe('bounded CSV and image contracts', () => {
  it('parses quoted CSV and rejects hostile dimensions and malformed quotes', () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/input.csv', 'a,"b,b"\r\n"c\nline",d')
    expect(readBoundedCsv(vfs, '/home/user/input.csv')).toEqual([
      ['a', 'b,b'],
      ['c\nline', 'd'],
    ])
    vfs.writeFile('/home/user/bad.csv', '"unterminated')
    expect(() => readBoundedCsv(vfs, '/home/user/bad.csv')).toThrow('invalid_csv')
    vfs.writeFile('/home/user/wide.csv', Array.from({ length: 101 }, () => 'x').join(','))
    expect(() => readBoundedCsv(vfs, '/home/user/wide.csv')).toThrow('import_limit')
  })

  it('neutralizes spreadsheet formulas on export', () => {
    expect(exportSafeCsv([['=cmd()', '+1', '-2', '@x', 'ok', 'a,b']])).toBe(
      "'=cmd(),'+1,'-2,'@x,ok,\"a,b\"",
    )
  })

  it('checks image magic, dimensions, pixels and bytes rather than file extension', () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/not-really.jpg', png(10, 20))
    expect(readBoundedImage(vfs, '/home/user/not-really.jpg')).toMatchObject({
      mime: 'image/png',
      width: 10,
      height: 20,
    })
    vfs.writeFile('/home/user/huge.png', png(8192, 8192))
    expect(() => readBoundedImage(vfs, '/home/user/huge.png')).toThrow('image_limit')
    vfs.writeFile('/home/user/fake.png', new Uint8Array([1, 2, 3]))
    expect(() => readBoundedImage(vfs, '/home/user/fake.png')).toThrow('image_mime_unsupported')
  })
})

describe('Excel import/export proposals', () => {
  const adapter = () => ({
    fingerprintRange: vi.fn().mockResolvedValue('before'),
    readRangeValues: vi.fn().mockResolvedValue([['=danger', 'safe']]),
    writeRangeValues: vi.fn().mockResolvedValue(undefined),
    verifyRangeValues: vi.fn().mockResolvedValue(true),
    insertImage: vi.fn().mockResolvedValue({ id: 'image-1' }),
    verifyImage: vi.fn().mockResolvedValue(true),
  })

  it('publishes exact schemas, exports safely, and gates CSV writes behind immutable proposals', async () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/in.csv', 'a,b\n1,2')
    const fake = adapter()
    const proposals = createStructuredProposalController()
    const skill = createExcelImportMediaSkill({ adapter: fake, proposals, vfs })
    expect(skill.tools.map((tool) => tool.name)).toEqual([
      'csv-to-sheet',
      'sheet-to-csv',
      'image-to-sheet',
    ])
    expect(skill.tools.every((tool) => tool.inputSchema.additionalProperties === false)).toBe(true)
    await expect(
      skill.executeTool(
        call('csv-to-sheet', {
          path: '/home/user/in.csv',
          sheetId: 1,
          startCell: 'A1',
          extra: true,
        }),
      ),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true })
    const result = await skill.executeTool(
      call('csv-to-sheet', { path: '/home/user/in.csv', sheetId: 1, startCell: 'A1' }),
    )
    expect(result).toMatchObject({ mutated: false, output: expect.stringContaining('proposalId') })
    expect(fake.writeRangeValues).not.toHaveBeenCalled()
    const pending = proposals.pending()!
    expect(Object.isFrozen(pending.preview)).toBe(true)
    await proposals.confirm(pending.id)
    expect(fake.writeRangeValues).toHaveBeenCalledOnce()
    expect(fake.verifyRangeValues).toHaveBeenCalledOnce()

    await skill.executeTool(
      call('sheet-to-csv', { sheetId: 1, range: 'A1:B1', path: '/home/user/out.csv' }),
    )
    expect(vfs.readText('/home/user/out.csv')).toBe("'=danger,safe")
  })

  it('rejects stale/cancelled proposals, executes exactly once, and fails post-write mismatch', async () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/in.csv', 'a')
    const fake = adapter()
    fake.fingerprintRange.mockResolvedValueOnce('before').mockResolvedValueOnce('changed')
    const proposals = createStructuredProposalController()
    const skill = createExcelImportMediaSkill({ adapter: fake, proposals, vfs })
    await skill.executeTool(
      call('csv-to-sheet', { path: '/home/user/in.csv', sheetId: 1, startCell: 'A1' }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(fake.writeRangeValues).not.toHaveBeenCalled()

    fake.fingerprintRange.mockResolvedValue('before')
    fake.verifyRangeValues.mockResolvedValue(false)
    await skill.executeTool(
      call('csv-to-sheet', { path: '/home/user/in.csv', sheetId: 1, startCell: 'A1' }),
    )
    const id = proposals.pending()!.id
    await expect(proposals.confirm(id)).rejects.toThrow('office_verify_failed')
    expect(fake.writeRangeValues).toHaveBeenCalledTimes(1)
    await expect(proposals.confirm(id)).rejects.toThrow('proposal_missing')
  })

  it('does not verify or claim success after a partial host failure and rejection cancels the write', async () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/in.csv', 'a')
    const fake = adapter()
    fake.writeRangeValues.mockRejectedValue(new Error('partial Office failure with document data'))
    const proposals = createStructuredProposalController()
    const skill = createExcelImportMediaSkill({ adapter: fake, proposals, vfs })
    await skill.executeTool(
      call('csv-to-sheet', { path: '/home/user/in.csv', sheetId: 1, startCell: 'A1' }),
    )
    proposals.reject()
    expect(fake.writeRangeValues).not.toHaveBeenCalled()
    await skill.executeTool(
      call('csv-to-sheet', { path: '/home/user/in.csv', sheetId: 1, startCell: 'A1' }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow()
    expect(fake.verifyRangeValues).not.toHaveBeenCalled()
  })
})

describe('PowerPoint image proposal', () => {
  it('revalidates the slide, inserts once, and semantically verifies the created shape', async () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/image.png', png(10, 10))
    const adapter = {
      snapshotSlide: vi.fn().mockResolvedValue({ slideId: 's1', fingerprint: 'fp' }),
      insertImage: vi.fn().mockResolvedValue({ id: 'pic1' }),
      verifyImage: vi.fn().mockResolvedValue(true),
    }
    const proposals = createStructuredProposalController()
    const skill = createPowerPointImportMediaSkill({ adapter, proposals, vfs })
    expect(skill.tools.map((tool) => tool.name)).toEqual(['insert-image'])
    await skill.executeTool(
      call('insert-image', {
        path: '/home/user/image.png',
        slide_index: 0,
        left: 1,
        top: 2,
        width: 30,
        height: 40,
      }),
    )
    await proposals.confirm(proposals.pending()!.id)
    expect(adapter.insertImage).toHaveBeenCalledOnce()
    expect(adapter.verifyImage).toHaveBeenCalledWith(
      0,
      'pic1',
      { left: 1, top: 2, width: 30, height: 40 },
      expect.any(AbortSignal),
    )
  })
})
