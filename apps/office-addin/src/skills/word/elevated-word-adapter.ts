import type {
  ElevatedOfficeAdapter,
  ElevatedOfficeAuthority,
  ElevatedOfficeProgram,
  ElevatedOfficeSnapshot,
} from '../shared/elevated-office-program.js'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import type { WordAdapter, WordDeclarativeOperation } from './browser-word-adapter.js'

/** The only reviewed authority boundary allowed to interpret Word elevated programs. */
export interface WordElevatedAuthority {
  captureAuthority(): ElevatedOfficeAuthority
  snapshot(program: ElevatedOfficeProgram, signal?: AbortSignal): Promise<ElevatedOfficeSnapshot>
  validateSnapshot(
    program: ElevatedOfficeProgram,
    snapshot: ElevatedOfficeSnapshot,
    signal?: AbortSignal,
  ): Promise<boolean>
  execute(
    program: ElevatedOfficeProgram,
    snapshot: ElevatedOfficeSnapshot,
    signal?: AbortSignal,
  ): Promise<void>
  readback(
    program: ElevatedOfficeProgram,
    snapshot: ElevatedOfficeSnapshot,
    signal?: AbortSignal,
  ): Promise<{ verified: boolean; output?: unknown }>
  rollback(snapshot: ElevatedOfficeSnapshot, signal?: AbortSignal): Promise<void>
}

export const createWordElevatedAdapter = (
  authority: WordElevatedAuthority,
): ElevatedOfficeAdapter => Object.freeze({ host: 'word', ...authority })

type WordState = { fingerprint: string; xml?: string }
const wordState = (snapshot: ElevatedOfficeSnapshot) => snapshot.state as WordState
const wordOperations = (program: ElevatedOfficeProgram): WordDeclarativeOperation[] => {
  if (program.kind !== 'office_js_ast') throw new Error('office_api_unsupported')
  return program.operations.map(({ call, args }) =>
    call === 'body.insertText'
      ? {
          op: 'insert_text',
          location: args.location as 'start' | 'end' | 'replace',
          text: args.text as string,
        }
      : {
          op: 'replace_all',
          search: args.search as string,
          replacement: args.replacement as string,
          matchCase: true,
        },
  )
}
async function replaceBodyOoxml(xml: string, signal?: AbortSignal) {
  if (signal?.aborted) throw new Error('cancelled')
  const word = (globalThis as any).Word
  if (typeof word?.run !== 'function') throw new Error('office_api_unsupported')
  await word.run(async (context: any) => {
    if (signal?.aborted) throw new Error('cancelled')
    context.document.body.insertOoxml(xml, 'Replace')
    await context.sync()
    if (signal?.aborted) throw new Error('cancelled')
  })
}

export function createBrowserWordElevatedAdapter(options: {
  adapter: WordAdapter
  authority(): ElevatedOfficeAuthority
}): ElevatedOfficeAdapter {
  return createWordElevatedAdapter({
    captureAuthority: options.authority,
    async snapshot(program, signal) {
      const document = await options.adapter.getDocumentSnapshot(signal)
      const xml = (await options.adapter.getOoxml({}, signal)).xml
      return {
        id: `history_${selectionFingerprint(`${document.fingerprint}:${xml ?? ''}`).replace(':', '_')}`,
        state: { fingerprint: document.fingerprint, xml } satisfies WordState,
      }
    },
    async validateSnapshot(_program, snapshot, signal) {
      return (await options.adapter.fingerprint(signal)) === wordState(snapshot).fingerprint
    },
    async execute(program, _snapshot, signal) {
      if (program.kind === 'office_js_ast')
        await options.adapter.executeOperations(wordOperations(program), signal)
      else {
        if (program.patches.length !== 1 || program.patches[0].part !== 'word/document.xml')
          throw new Error('office_api_unsupported')
        await replaceBodyOoxml(program.patches[0].xml, signal)
      }
    },
    async readback(program, snapshot, signal) {
      if (program.kind === 'office_js_ast')
        return { verified: await options.adapter.verifyOperations(wordOperations(program), signal) }
      const current = await options.adapter.getOoxml({}, signal)
      return {
        verified: current.xml === program.patches[0].xml,
        output: { changed: current.xml !== wordState(snapshot).xml },
      }
    },
    async rollback(snapshot, signal) {
      const xml = wordState(snapshot).xml
      if (!xml) throw new Error('office_state_uncertain')
      await replaceBodyOoxml(xml, signal)
    },
  })
}
