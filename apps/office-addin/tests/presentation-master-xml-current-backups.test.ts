import { afterEach, expect, it } from 'vitest'
import JSZip from 'jszip'
import { masterXmlFixture, cleanupMasterXmlFixtures } from './helpers/master-xml-fixture'
import { readPackageBackup } from '../src/skills/powerpoint/presentation-package-backup'
afterEach(cleanupMasterXmlFixtures)
const hash = async (bytes: Uint8Array) =>
  Buffer.from(await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes).buffer)).toString('hex')
async function changed(base64: string, text: string) {
  const zip = await JSZip.loadAsync(base64, { base64: true })
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><a:t>${text}</a:t></p:spTree></p:cSld></p:sld>`,
  )
  return zip.generateAsync({ type: 'base64' })
}
async function uncertain(unaffected = false) {
  const f = await masterXmlFixture(3),
    originalApply = f.adapter.applyLayout.getMockImplementation()!
  if (unaffected) {
    const zip = await JSZip.loadAsync(f.original, { base64: true })
    zip.file(
      'ppt/slideMasters/slideMaster1.xml',
      f.originalMaster.replace('name="original"', 'name="unrelated master"'),
    )
    const base64 = await zip.generateAsync({ type: 'base64' })
    f.masters.push({
      masterId: 'unrelated',
      name: 'same',
      base64,
      layouts: [1, 2].map((i) => ({
        layoutId: `unrelated-l${i}`,
        name: 'duplicate',
        path: `ppt/slideLayouts/slideLayout${i}.xml`,
      })),
    })
    f.packages.set('s2', base64)
    f.deps.set('s2', { slideId: 's2', masterId: 'unrelated', layoutId: 'unrelated-l1' })
  }
  const current = new Map<string, string>()
  let injected = false
  f.adapter.applyLayout.mockImplementation(async (...args) => {
    await originalApply(...args)
    if (args[0].slideId === 's1' && !injected) {
      injected = true
      for (const id of ['s1', 's2']) {
        const bytes = await changed(f.packages.get(id)!, `current manual edit ${id}`)
        f.packages.set(id, bytes)
        current.set(id, bytes)
      }
      throw Error('lost_ack')
    }
  })
  const proposal = await f.propose(),
    id = String(proposal.preview.changeId)
  await expect(f.confirm()).rejects.toThrow()
  expect(injected).toBe(true)
  expect(f.data.get(id)!.pending).toBeDefined()
  f.reopen()
  return { f, id, current }
}
async function assertSavedCurrent(
  f: Awaited<ReturnType<typeof masterXmlFixture>>,
  id: string,
  current: Map<string, string>,
) {
  for (const [slideId, base64] of current) {
    const expected = Uint8Array.from(Buffer.from(base64, 'base64')),
      sha256 = await hash(expected)
    const begin = f.request.mock.calls
      .map(([body]) => body as Record<string, unknown>)
      .find(
        (body) =>
          body.operation === 'package_backup_begin' &&
          body.changeId === id &&
          body.sha256 === sha256 &&
          body.sizeBytes === expected.length,
      )
    expect(begin, `missing current package backup for ${slideId}`).toBeDefined()
    const bytes = await readPackageBackup({
      request: f.request,
      documentId: 'doc',
      changeId: id,
      backup: { key: String(begin!.key), sha256, sizeBytes: expected.length },
    })
    expect(await hash(bytes)).toBe(sha256)
    expect(bytes).toEqual(expected)
  }
}
it('backs up all changed owned current bytes before recovery acknowledgement and any destructive original restoration', async () => {
  const { f, id, current } = await uncertain()
  const inspected = await f.tool('inspect', { change_id: id })
  expect(inspected.isError).not.toBe(true)
  expect(JSON.parse(inspected.output).observed).not.toBe('ready')
  const write = f.options.writeMasterXmlChange
  let savedBeforeRecovery = false
  f.options.writeMasterXmlChange = async (next, expected) => {
    if (next.state === 'recovery_required') {
      await assertSavedCurrent(f, id, current)
      savedBeforeRecovery = true
    }
    return write(next, expected)
  }
  const reconcile = await f.tool('reconcile', { change_id: id })
  expect(reconcile.isError, JSON.stringify(reconcile)).not.toBe(true)
  const removesBefore = f.adapter.remove.mock.calls.length
  await f.confirm()
  expect(savedBeforeRecovery).toBe(true)
  expect(f.data.get(id)!.state).toBe('recovery_required')
  const proofBytes = await readPackageBackup({
    request: f.request,
    documentId: 'doc',
    changeId: id,
    backup: f.data.get(id)!.currentProofRef,
  })
  const proof = JSON.parse(new TextDecoder().decode(proofBytes))
  for (const [slideId, base64] of current) {
    const bytes = Uint8Array.from(Buffer.from(base64, 'base64'))
    const sha256 = await hash(bytes)
    const row = proof.progress.currentBackups.find(
      (row: { slideId: string; packageRef: { sha256: string } }) =>
        row.slideId === slideId && row.packageRef.sha256 === sha256,
    )
    expect(row, `current bytes must be reachable from durable proof for ${slideId}`).toBeDefined()
    expect(row.packageRef).toMatchObject({ sha256: await hash(bytes), sizeBytes: bytes.length })
    expect(
      await readPackageBackup({
        request: f.request,
        documentId: 'doc',
        changeId: id,
        backup: row.packageRef,
      }),
    ).toEqual(bytes)
  }

  expect(f.adapter.remove.mock.calls).toHaveLength(removesBefore)
  const remove = f.adapter.remove.getMockImplementation()!
  f.adapter.remove.mockImplementation(async (...args) => {
    await assertSavedCurrent(f, id, current)
    return remove(...args)
  })
  expect((await f.tool('undo', { change_id: id })).isError).not.toBe(true)
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('undone')
  expect(f.adapter.remove.mock.calls.length).toBeGreaterThan(removesBefore)
  await assertSavedCurrent(f, id, current)
}, 120000)
it('refuses unknown page identities without recovery acknowledgement or deletion', async () => {
  const { f, id } = await uncertain()
  f.order.push('foreign-user-page')
  f.packages.set('foreign-user-page', f.original)
  f.deps.set('foreign-user-page', { slideId: 'foreign-user-page', masterId: 'm1', layoutId: 'l1' })
  const before = f.adapter.remove.mock.calls.length
  const result = await f.tool('reconcile', { change_id: id })
  expect(result.isError).toBe(true)
  expect(f.proposals.pending()).toBeUndefined()
  expect(f.data.get(id)!.state).not.toBe('recovery_required')
  expect(f.adapter.remove.mock.calls).toHaveLength(before)
}, 120000)

it('refuses changed content outside the proven affected native master scope', async () => {
  const { f, id } = await uncertain(true)
  const before = f.adapter.remove.mock.calls.length
  const result = await f.tool('reconcile', { change_id: id })
  expect(result.isError).toBe(true)
  expect(f.proposals.pending()).toBeUndefined()
  expect(f.data.get(id)!.state).not.toBe('recovery_required')
  expect(f.adapter.remove.mock.calls).toHaveLength(before)
}, 120000)
