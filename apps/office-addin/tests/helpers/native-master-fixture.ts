import { expect, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { createPresentationMasterBackupService } from '../../../shell/src/main/presentation-master-backups.js'
import { createPresentationNativeMasterSkill } from '../../src/skills/powerpoint/presentation-native-master.js'
import {
  validatePresentationNativeMasterChange,
  validNativeMasterTransition,
  type PresentationNativeMasterChange,
} from '../../src/skills/powerpoint/presentation-native-master-change.js'
import { projectedMasterState } from '../../src/skills/powerpoint/presentation-master-program.js'
import type {
  PowerPointAdapter,
  PowerPointMasterOperation,
  PowerPointMasterState,
} from '../../src/skills/powerpoint/browser-powerpoint-adapter.js'
import { createStructuredProposalController } from '../../src/agent/proposal-controller.js'
const roots: string[] = []
export function cleanupNativeMasterFixtures() {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
}
export async function nativeMasterFixture(count = 2) {
  const root = mkdtempSync(join(tmpdir(), 'native-master-'))
  roots.push(root)
  const service = createPresentationMasterBackupService({ userDataPath: root })
  const masters = Array.from({ length: 2 }, (_, i) => ({
    id: `m${i}`,
    name: `Master${i}`,
    background: { type: 'Solid', color: '#FFFFFF', transparency: 0 },
    themeColors: Object.fromEntries(
      [
        'Accent1',
        'Accent2',
        'Accent3',
        'Accent4',
        'Accent5',
        'Accent6',
        'Dark1',
        'Dark2',
        'Light1',
        'Light2',
        'Hyperlink',
        'FollowedHyperlink',
      ].map((key) => [key, '#FFFFFF']),
    ),
    layouts: [
      {
        id: `l${i}`,
        name: 'Layout',
        isMasterBackgroundFollowed: true,
        areBackgroundGraphicsHidden: false,
        background: { type: 'Solid' },
      },
    ],
  }))
  let native: PowerPointMasterState = { masters },
    documentId = 'doc',
    available = true
  let hostMode = 'normal',
    receiptMode = 'normal',
    beforeWrite: () => void = () => {},
    sdkRead: () => void = () => {}
  const order = Array.from({ length: count }, (_, i) => `s${i}`)
  const dependencies = {
    slides: order.map((slideId, i) => ({ slideId, masterId: `m${i % 2}`, layoutId: `l${i % 2}` })),
  }
  const data = new Map<string, PresentationNativeMasterChange>()
  const drift = new Map<string, string>()
  const images = new Map<string, string>()
  const started = Date.now()
  let exportCount = 0
  let nativeCalls = 0
  const packageFor = async (index: number) => {
    const zip = new JSZip(),
      master = native.masters[index % 2]!,
      layout = master.layouts[0]!
    const relation = (id: string, type: string, target: string) =>
      `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`
    const relationships = (rows: string) =>
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rows}</Relationships>`
    const picture =
      master.background.type === 'PictureOrTexture' ? images.get(master.id) : undefined
    let bg = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${master.background.color?.replace('#', '') ?? 'FFFFFF'}"/></a:solidFill></p:bgPr></p:bg>`
    if (drift.has('original_bg')) bg = drift.get('original_bg')!
    if (picture)
      bg = '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="image"/></a:blipFill></p:bgPr></p:bg>'
    zip.file(
      'ppt/slides/slide1.xml',
      `<p:sld xmlns:p="urn:p"><p:extLst>${index}-${drift.get(order[index]!) ?? ''}</p:extLst></p:sld>`,
    )
    zip.file(
      'ppt/slides/_rels/slide1.xml.rels',
      relationships(relation('layout', 'slideLayout', '../slideLayouts/slideLayout1.xml')),
    )
    zip.file(
      'ppt/slideLayouts/slideLayout1.xml',
      `<p:sldLayout xmlns:p="urn:p" showMasterSp="${!layout.areBackgroundGraphicsHidden}"><p:cSld>${layout.isMasterBackgroundFollowed ? '' : '<p:bg><p:bgPr/></p:bg>'}<p:spTree/></p:cSld></p:sldLayout>`,
    )
    zip.file(
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
      relationships(relation('master', 'slideMaster', '../slideMasters/slideMaster1.xml')),
    )
    zip.file(
      'ppt/slideMasters/slideMaster1.xml',
      `<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r"><p:cSld>${bg}<p:spTree><p:sp><p:txBody><a:p><a:r><a:t>${drift.get('master_placeholder') ?? 'Original placeholder'}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sldMaster>`,
    )
    const slots: Record<string, string> = {
      Accent1: 'accent1',
      Accent2: 'accent2',
      Accent3: 'accent3',
      Accent4: 'accent4',
      Accent5: 'accent5',
      Accent6: 'accent6',
      Dark1: 'dk1',
      Dark2: 'dk2',
      Light1: 'lt1',
      Light2: 'lt2',
      Hyperlink: 'hlink',
      FollowedHyperlink: 'folHlink',
    }
    zip.file(
      'ppt/theme/theme1.xml',
      `<a:theme xmlns:a="urn:a"><a:themeElements><a:clrScheme name="Office">${Object.entries(
        master.themeColors,
      )
        .map(
          ([key, color]) =>
            `<a:${slots[key]}>${key === 'Accent1' && drift.has('original_slot') ? drift.get('original_slot') : `<a:srgbClr val="${color.replace('#', '')}"/>`}</a:${slots[key]}>`,
        )
        .join(
          '',
        )}</a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="${drift.get('theme_font') ?? 'Original font'}"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>`,
    )
    zip.file(
      'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      relationships(
        relation('theme', 'theme', '../theme/theme1.xml') +
          (picture ? relation('image', 'image', '../media/background.png') : ''),
      ),
    )
    zip.file(
      'ppt/notesSlides/notesSlide1.xml',
      `<p:notes xmlns:p="urn:p"><p:extLst>${drift.get('note') ?? 'Original note'}</p:extLst></p:notes>`,
    )
    if (picture)
      zip.file(
        'ppt/media/background.png',
        Buffer.from(
          hostMode === 'wrong_picture' ? picture.slice(0, -4) + 'AAAA' : picture,
          'base64',
        ),
      )
    zip.file(
      '[Content_Types].xml',
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>${picture ? '<Default Extension="png" ContentType="image/png"/>' : ''}</Types>`,
    )
    return zip.generateAsync({ type: 'base64' })
  }
  const adapter = {
    inspectSlideMasters: vi.fn(async () => {
      sdkRead()
      return native
    }),
    inspectStyleDependencies: vi.fn(async () => dependencies),
    readSlideOrder: vi.fn(async () => order),
    exportSlidePackage: vi.fn(async (index: number) => {
      exportCount++
      if (count === 600 && exportCount % 600 === 0) {
        const requestOperations: Record<string, number> = {}
        for (const [body] of request.mock.calls) {
          const op = String((body as any).operation)
          requestOperations[op] = (requestOperations[op] ?? 0) + 1
        }
        console.info('native-master600 export phase', {
          elapsedMs: Date.now() - started,
          exportCount,
          nativeCalls,
          requestOperations,
        })
      }
      return { slideId: order[index]!, base64: await packageFor(index), fingerprint: 'volatile' }
    }),
    screenshotSlide: vi.fn(async (index: number) => ({
      slideId: order[index]!,
      mime: 'image/png',
      base64:
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII=',
    })),
    executeMasterOperations: vi.fn(async (ops: PowerPointMasterOperation[]) => {
      nativeCalls++
      beforeWrite()
      if (hostMode === 'before_failure') throw Error('sync_failed')
      native = projectedMasterState(native, ops)
      for (const op of ops)
        if (op.op === 'set_master_background' && op.fill.type === 'picture_or_texture')
          images.set(op.master_id, op.fill.image_base64)
      if (hostMode === 'after_failure') throw Error('sync_failed')
    }),
  } as unknown as PowerPointAdapter
  const request = vi.fn(
    async (body: unknown, signal?: AbortSignal) =>
      new Response(
        JSON.stringify(
          await service(body as Record<string, unknown>, signal ?? new AbortController().signal),
        ),
      ),
  )
  const write = vi.fn(
    async (
      next: PresentationNativeMasterChange,
      expected: PresentationNativeMasterChange | undefined,
    ) => {
      expect(validatePresentationNativeMasterChange(next)).toBe(true)
      expect(validNativeMasterTransition(expected, next)).toBe(true)
      expect(data.get(next.changeId)).toEqual(expected)
      if (receiptMode === 'fail_observed' && next.pending?.afterProofRef)
        throw Error('receipt_failed')
      data.set(next.changeId, structuredClone(next))
      if (receiptMode === 'lost_ack' && next.pending?.afterProofRef) throw Error('ack_lost')
    },
  )
  const proposals = createStructuredProposalController()
  const make = () =>
    createPresentationNativeMasterSkill({
      documentId: async () => documentId,
      assertDocumentId: (expected) => {
        if (documentId !== expected) throw Error('presentation_document_changed')
      },
      available: () => available,
      adapter,
      request,
      proposals,
      readNativeMasterChange: (id) => data.get(id),
      writeNativeMasterChange: write,
    })
  let skill = make()
  const op: PowerPointMasterOperation = {
    op: 'set_master_theme_color',
    master_id: 'm0',
    theme_color: 'Accent1',
    color: '#000000',
  }
  return {
    adapter,
    request,
    write,
    proposals,
    data,
    op,
    order,
    dependencies,
    drift,
    proposeWith: (ops: PowerPointMasterOperation[], explanation?: string, signal?: AbortSignal) =>
      skill.propose(ops, explanation, signal),
    skill: () => skill,
    propose: (ops: PowerPointMasterOperation[] = [op], signal?: AbortSignal) =>
      skill.propose(ops, undefined, signal),
    tool: (name: string, input: Record<string, unknown>) =>
      skill.executeTool({ id: 'call', name, input }),
    confirm: () => proposals.confirm(proposals.pending()!.id),
    reopen: () => {
      skill = make()
    },
    setHost: (mode: string) => {
      hostMode = mode
    },
    setReceipt: (mode: string) => {
      receiptMode = mode
    },
    setBeforeWrite: (fn: () => void) => {
      beforeWrite = fn
    },
    setSdkRead: (fn: () => void) => {
      sdkRead = fn
    },
    setDocument: () => {
      documentId = 'other'
    },
    disconnect: () => {
      available = false
    },
    clear: () => skill.clear(),
    native: () => native,
  }
}
