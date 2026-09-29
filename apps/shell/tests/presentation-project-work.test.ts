import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { PresentationStore } from '@wiswork/project-store'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { handlePresentationJob, hasPresentationWorker } from '../src/main/presentation-jobs'
import {
  registerPresentationProjectWork,
  hasPresentationProjectWork,
  stopPresentationProjectWork,
} from '../src/main/presentation-project-work'
const roots: string[] = [],
  finishes: (() => void)[] = []
afterEach(() => {
  for (const finish of finishes.splice(0)) finish()
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'project-work-'))
  roots.push(root)
  return { root, projectId: 'p', documentId: 'doc' }
}
function register(scope: ReturnType<typeof fixture>, signal?: AbortSignal) {
  const work = registerPresentationProjectWork({ scope, signal })
  finishes.push(work.finish)
  return work
}
function deferred() {
  let resolve!: () => void
  const promise = new Promise<void>((done) => {
    resolve = done
  })
  return { promise, resolve }
}
it('observes active foreground work only for the exact project and document', () => {
  const scope = fixture(),
    other = fixture()
  expect(hasPresentationProjectWork(scope)).toBe(false)
  const work = register(scope)
  expect(hasPresentationProjectWork(scope)).toBe(true)
  expect(hasPresentationProjectWork({ ...scope, documentId: 'other' })).toBe(false)
  expect(hasPresentationProjectWork({ ...scope, projectId: 'other' })).toBe(false)
  expect(hasPresentationProjectWork(other)).toBe(false)
  work.finish()
  expect(hasPresentationProjectWork(scope)).toBe(false)
})
it('keeps multiple same-scope requests and finishes only their own tokens', async () => {
  const scope = fixture(),
    first = register(scope),
    second = register(scope),
    third = register(scope)
  expect(new Set([first.token, second.token, third.token]).size).toBe(3)
  first.finish()
  first.finish()
  const drain = stopPresentationProjectWork(scope)
  expect(first.signal.aborted).toBe(false)
  expect(second.signal.aborted).toBe(true)
  expect(third.signal.aborted).toBe(true)
  let settled = false
  void drain.then(() => {
    settled = true
  })
  second.finish()
  await Promise.resolve()
  expect(settled).toBe(false)
  third.finish()
  await drain
  expect(settled).toBe(true)
})
it('fixes caller scope and aborts only exact root/project/document entries', async () => {
  const scope = fixture(),
    original = { ...scope },
    own = register(scope),
    otherDoc = register({ ...scope, documentId: 'other' }),
    otherProject = register({ ...scope, projectId: 'other' }),
    otherRoot = register(fixture())
  scope.documentId = 'mutated'
  const stopScope = { ...original },
    drain = stopPresentationProjectWork(stopScope)
  stopScope.projectId = 'mutated'
  expect(own.scope).toEqual(original)
  expect(Object.isFrozen(own.scope)).toBe(true)
  expect(own.signal.aborted).toBe(true)
  for (const other of [otherDoc, otherProject, otherRoot]) expect(other.signal.aborted).toBe(false)
  own.finish()
  await drain
})
it('merges client cancellation into an owned controller without claiming unfinished work stopped', async () => {
  const scope = fixture(),
    client = new AbortController(),
    own = register(scope, client.signal)
  expect(own.signal).not.toBe(client.signal)
  client.abort()
  expect(own.signal.aborted).toBe(true)
  let done = false
  void own.settled.then(() => {
    done = true
  })
  await Promise.resolve()
  expect(done).toBe(false)
  own.finish()
  await own.settled
  expect(done).toBe(true)
  expect(register(scope, client.signal).signal.aborted).toBe(true)
})
it('drains actual background and foreground work while an unsignalled compiler must truly return', async () => {
  const scope = fixture(),
    store = new PresentationStore(scope.root),
    plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  scope.projectId = deck.id
  plan.slides = plan.slides.slice(0, 2)
  deck.slides = deck.slides.slice(0, 2)
  store.beginProduction(scope.projectId, scope.documentId, 'run', deck, { revision: 1, plan })
  const entered = deferred(),
    gate = deferred(),
    key = `${resolve(scope.root)}\0${scope.projectId}`
  handlePresentationJob(
    key,
    {
      operation: 'production_job_start',
      documentId: scope.documentId,
      projectId: scope.projectId,
      requestId: 'run',
    },
    {
      store,
      compile: async (input) => {
        entered.resolve()
        await gate.promise
        return compilePresentationDeck(input)
      },
      attachments: async () => {
        throw Error('unused')
      },
    },
  )
  await entered.promise
  expect(hasPresentationProjectWork(scope)).toBe(true)
  await stopPresentationProjectWork({ ...scope, documentId: 'other' })
  expect(hasPresentationWorker(key)).toBe(true)
  const foreground = register(scope),
    drain = stopPresentationProjectWork(scope)
  let done = false
  void drain.then(() => {
    done = true
  })
  foreground.finish()
  await Promise.resolve()
  expect(done).toBe(false)
  expect(hasPresentationWorker(key)).toBe(true)
  gate.resolve()
  await drain
  expect(done).toBe(true)
  expect(hasPresentationWorker(key)).toBe(false)
  expect(hasPresentationProjectWork(scope)).toBe(false)
})

it('waits for actual foreground compilation even when the compiler has no cancellation signal', async () => {
  const scope = fixture(),
    client = new AbortController(),
    own = register(scope, client.signal),
    gate = deferred()
  const deck = benchmarkPlannedDeck()
  let returned = false
  const task = (async () => {
    try {
      await gate.promise
      await compilePresentationDeck(deck)
      returned = true
    } finally {
      own.finish()
    }
  })()
  const drain = stopPresentationProjectWork(scope)
  expect(own.signal.aborted).toBe(true)
  expect(client.signal.aborted).toBe(false)
  let drained = false
  void drain.then(() => {
    drained = true
  })
  await Promise.resolve()
  expect(returned).toBe(false)
  expect(drained).toBe(false)
  gate.resolve()
  await Promise.all([task, drain])
  expect(returned).toBe(true)
  expect(drained).toBe(true)
})

it('releases only its client abort listener after real completion', () => {
  const scope = fixture(),
    client = new AbortController(),
    own = register(scope, client.signal)
  own.finish()
  client.abort()
  expect(own.signal.aborted).toBe(false)
})
