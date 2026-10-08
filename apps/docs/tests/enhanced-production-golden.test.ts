// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'
import { runEnhancedGolden } from '../../../packages/agent-runtime/src/production-golden'
import {
  createHostGoldenBridge,
  writeEnhancedGoldenReport,
} from '../../../packages/agent-runtime/tests/host-golden-bridge'
import { createDocsSkill } from '../src/renderer/ai/docs-skill'
import { editorExtensions } from '../src/renderer/editor/extensions'

describe('Docs production Enhanced golden', () => {
  let editor: Editor | undefined
  afterEach(() => editor?.destroy())

  it('runs the production atomic block command with snapshot, readback and undo', async () => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: null },
            content: [{ type: 'text', text: 'before' }],
          },
        ],
      },
    })
    const call = {
      id: 'docs-golden-call',
      name: 'replace_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>after</p>' },
    }
    const result = await runEnhancedGolden('docs', {
      documentId: 'docs-golden-document',
      generation: 1,
      instruction: 'Replace the paragraph',
      bridge: createHostGoldenBridge({
        documentId: 'docs-golden-document',
        generation: 1,
        call,
      }),
      captureSnapshot: () => editor!.getJSON(),
      skill: createDocsSkill(
        () => editor!,
        () => ({ bullet: null, ordered: null }),
      ),
      confirm: async () => ({ mutationReceiptId: 'docs-golden-receipt' }),
      readback: async () => ({
        status: editor!.getText() === 'after' ? 'verified' : 'failed',
      }),
      rollback: async () => ({
        status: editor!.commands.undo() ? ('restored' as const) : ('failed' as const),
      }),
    })
    expect(result.verification).toEqual({ status: 'verified' })
    expect(result.rollback).toEqual({ status: 'restored' })
    expect(editor.getText()).toBe('before')
    writeEnhancedGoldenReport({
      host: 'docs',
      verification: result.verification.status,
      rollback: result.rollback.status,
    })
  })
})
