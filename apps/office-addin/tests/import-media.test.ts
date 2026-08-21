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
  const crc32 = (bytes: Uint8Array) => {
    let crc = 0xffffffff
    for (const byte of bytes) {
      crc ^= byte
      for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0)
    }
    return (crc ^ 0xffffffff) >>> 0
  }
  const chunk = (name: string, data: Uint8Array) => {
    const result = new Uint8Array(12 + data.length)
    const view = new DataView(result.buffer)
    view.setUint32(0, data.length)
    result.set(new TextEncoder().encode(name), 4)
    result.set(data, 8)
    view.setUint32(8 + data.length, crc32(result.subarray(4, 8 + data.length)))
    return result
  }
  const ihdr = new Uint8Array(13)
  const view = new DataView(ihdr.buffer)
  view.setUint32(0, width)
  view.setUint32(4, height)
  ihdr.set([8, 6, 0, 0, 0], 8)
  const parts = [
    new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    chunk('IHDR', ihdr),
    chunk('IDAT', new Uint8Array([0])),
    chunk('IEND', new Uint8Array()),
  ]
  const result = new Uint8Array(parts.reduce((sum, item) => sum + item.length, 0))
  let offset = 0
  for (const part of parts) {
    result.set(part, offset)
    offset += part.length
  }
  return result
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
    vfs.writeFile('/home/user/comma-stream.csv', ','.repeat(20_000))
    expect(() => readBoundedCsv(vfs, '/home/user/comma-stream.csv')).toThrow('import_limit')
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
    const corrupt = png(10, 10)
    corrupt[corrupt.length - 1] ^= 1
    vfs.writeFile('/home/user/corrupt.png', corrupt)
    expect(() => readBoundedImage(vfs, '/home/user/corrupt.png')).toThrow('invalid_image')
    vfs.writeFile(
      '/home/user/truncated.jpg',
      new Uint8Array([0xff, 0xd8, 0xff, 0xc0, 0, 17, 8, 0, 1, 0, 1]),
    )
    expect(() => readBoundedImage(vfs, '/home/user/truncated.jpg')).toThrow('invalid_image')
    const jpeg = new Uint8Array([
      0xff, 0xd8, 0xff, 0xc0, 0x00, 0x0b, 0x08, 0x00, 0x01, 0x00, 0x01, 0x01, 0x01, 0x11, 0x00,
      0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3f, 0x00, 0x00, 0xff, 0xd9,
    ])
    vfs.writeFile('/home/user/valid.jpg', jpeg)
    expect(readBoundedImage(vfs, '/home/user/valid.jpg')).toMatchObject({
      mime: 'image/jpeg',
      width: 1,
      height: 1,
    })
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
    captureRange: vi.fn().mockResolvedValue({ formulas: [['old']] }),
    restoreRange: vi.fn().mockResolvedValue(undefined),
    verifyRangeSnapshot: vi.fn().mockResolvedValue(true),
    removeImage: vi.fn().mockResolvedValue(undefined),
    verifyImageAbsent: vi.fn().mockResolvedValue(true),
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
    expect(fake.restoreRange).toHaveBeenCalledOnce()
    expect(fake.verifyRangeSnapshot).toHaveBeenCalledOnce()
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
    expect(fake.restoreRange).toHaveBeenCalledOnce()
    expect(fake.verifyRangeSnapshot).toHaveBeenCalledOnce()
  })

  it('restores the captured range when logout cancels an in-flight host write', async () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/in.csv', 'a')
    const fake = adapter()
    fake.writeRangeValues.mockImplementation(
      (_sheet, _cell, _values, signal?: AbortSignal) =>
        new Promise((_resolve, reject) =>
          signal?.addEventListener('abort', () => reject(new Error('cancelled')), { once: true }),
        ),
    )
    const proposals = createStructuredProposalController()
    const skill = createExcelImportMediaSkill({ adapter: fake, proposals, vfs })
    await skill.executeTool(
      call('csv-to-sheet', { path: '/home/user/in.csv', sheetId: 1, startCell: 'A1' }),
    )
    const confirmation = proposals.confirm(proposals.pending()!.id)
    await vi.waitFor(() => expect(fake.writeRangeValues).toHaveBeenCalledOnce())
    proposals.logout()
    await expect(confirmation).rejects.toThrow('cancelled')
    expect(fake.restoreRange).toHaveBeenCalledOnce()
    expect(fake.verifyRangeSnapshot).toHaveBeenCalledOnce()
  })

  it('returns terminal recovery failure when the restored range cannot be proven', async () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/in.csv', 'a')
    const fake = adapter()
    fake.writeRangeValues.mockRejectedValue(new Error('partial write'))
    fake.verifyRangeSnapshot.mockResolvedValue(false)
    const proposals = createStructuredProposalController()
    const skill = createExcelImportMediaSkill({ adapter: fake, proposals, vfs })
    await skill.executeTool(
      call('csv-to-sheet', { path: '/home/user/in.csv', sheetId: 1, startCell: 'A1' }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow(
      'office_recovery_failed',
    )
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
      removeImage: vi.fn().mockResolvedValue(undefined),
      verifyImageAbsent: vi.fn().mockResolvedValue(true),
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

  it('deletes and proves absence when post-write image verification fails', async () => {
    const vfs = new InMemoryVfs()
    vfs.writeFile('/home/user/image.png', png(10, 10))
    const adapter = {
      snapshotSlide: vi.fn().mockResolvedValue({ slideId: 's1', fingerprint: 'fp' }),
      insertImage: vi.fn().mockResolvedValue({ id: 'pic1' }),
      verifyImage: vi.fn().mockResolvedValue(false),
      removeImage: vi.fn().mockResolvedValue(undefined),
      verifyImageAbsent: vi.fn().mockResolvedValue(true),
    }
    const proposals = createStructuredProposalController()
    const skill = createPowerPointImportMediaSkill({ adapter, proposals, vfs })
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
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('office_verify_failed')
    expect(adapter.removeImage).toHaveBeenCalledWith(0, 'pic1')
    expect(adapter.verifyImageAbsent).toHaveBeenCalledWith(0, 'pic1')
  })
})
