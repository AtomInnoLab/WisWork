import { afterEach, describe, expect, it, vi } from 'vitest'
import { PNG } from 'pngjs'
import { createOfficeAgentSession } from '../src/agent/use-office-agent.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { createOfficeDiagnostics } from '../src/diagnostics/office-diagnostics.js'
import { createOfficeRelaySession, type RelayWebSocket } from '../src/relay/session.js'
import { createPowerPointImportMediaSkill } from '../src/skills/powerpoint/powerpoint-import-media.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'

afterEach(() => vi.unstubAllGlobals())

describe('Enhanced PowerPoint image insertion', () => {
  it.each([false, true])(
    'keeps one active agent request with prefetched image bytes: %s',
    async (prefetched) => {
      const sent: Record<string, any>[] = []
      const socket: RelayWebSocket = {
        readyState: 1,
        onopen: null,
        onclose: null,
        onerror: null,
        onmessage: null,
        send: (value) => sent.push(JSON.parse(value)),
        close: vi.fn(),
      }
      const receive = (frame: Record<string, unknown>) =>
        socket.onmessage?.({ data: JSON.stringify({ version: 2, ...frame }) })
      const bridge = createOfficeRelaySession({
        createSocket: () => socket,
        persistentPairing: false,
        capabilities: ['agent.v1', 'image-fetch.v1'],
        randomUUID: () => 'request_12345678',
      })
      const connecting = bridge.connect('powerpoint')
      socket.onopen?.()
      receive({
        type: 'office.created',
        pairing_id: 'pair_12345678',
        verification_code: '123456',
        expires_in: 120,
      })
      receive({
        type: 'office.approved',
        session_id: 'session_12345678',
        capability: 'capability_12345678',
        capabilities: ['agent.v1', 'image-fetch.v1'],
        expires_in: 1800,
      })
      await connecting
      receive({
        type: 'relay.session_state',
        session_id: 'session_12345678',
        generation: 1,
        enhanced: {
          version: 1,
          runtime_mode: 'enhanced',
          runtime_instance: 'runtime_0123456789abcdef',
          component_version: '0.147.0',
          host: 'office-powerpoint',
          raw_office: false,
          expires_at: Date.now() + 60_000,
          policy_generation: 1,
          session_generation: 1,
        },
      })
      const image = new PNG({ width: 1, height: 1 })
      image.data.fill(0)
      const base64 = PNG.sync.write(image).toString('base64')
      vi.stubGlobal('createImageBitmap', async () => ({ width: 1, height: 1, close() {} }))
      const adapter = {
        snapshotSlide: vi.fn(async () => ({ slideId: 's1', fingerprint: 'fp' })),
        insertImage: vi.fn(async () => ({ id: 'image-1' })),
        verifyImage: vi.fn(async () => true),
        removeImage: vi.fn(),
        verifyImageAbsent: vi.fn(async () => true),
      }
      const proposals = createStructuredProposalController()
      const proposed = vi.spyOn(proposals, 'propose')
      const fetchImage = vi.fn(async (url: string, signal?: AbortSignal) => {
        const response = await bridge.capabilityFetch('image-fetch.v1', { url }, signal)
        const payload = await response.json()
        return Uint8Array.from(atob(payload.data_base64), (value) => value.charCodeAt(0))
      })
      const skill = createPowerPointImportMediaSkill({
        adapter,
        proposals,
        vfs: new InMemoryVfs(),
        fetchImage,
      })
      const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'test' })
      const session = createOfficeAgentSession({
        transport: { stream: () => ({ cancel() {} }) },
        skill,
        proposals,
        remoteTools: bridge,
        automaticPowerPointMutations: true,
        diagnostics,
      })
      const request = bridge.capabilityFetch('agent.v1', {}).catch(() => undefined)
      try {
        receive({
          type: 'relay.tool_call',
          session_id: 'session_12345678',
          request_id: 'request_12345678',
          turn_id: 'turn_12345678',
          call_id: 'call_12345678',
          generation: 1,
          tool_name: 'insert_web_image',
          input: {
            url: 'https://images.example/approved.png',
            slide_index: 0,
            left: 1,
            top: 2,
            width: 30,
            height: 40,
            ...(prefetched ? { _wiswork_image_base64: base64 } : {}),
          },
        })
        await vi.waitFor(() =>
          expect(sent.some((frame) => frame.type === 'office.tool_result')).toBe(true),
        )
        const result = sent.find((frame) => frame.type === 'office.tool_result')!
        expect(sent.filter((frame) => frame.type === 'office.request')).toHaveLength(1)
        expect(bridge.snapshot().status).toBe('connected')
        if (prefetched) {
          expect(result.is_error).toBe(false)
          expect(adapter.insertImage).toHaveBeenCalledWith(
            0,
            base64,
            { left: 1, top: 2, width: 30, height: 40 },
            expect.any(AbortSignal),
          )
          expect(fetchImage).not.toHaveBeenCalled()
          expect(JSON.stringify(result)).not.toContain(base64)
          expect(JSON.stringify(proposed.mock.calls)).not.toContain(base64)
          expect(JSON.stringify(session.snapshot())).not.toContain(base64)
          expect(diagnostics.exportJson()).not.toContain(base64)
        } else {
          expect(result).toMatchObject({ is_error: true, output: 'image_fetch_unavailable' })
          expect(adapter.snapshotSlide).not.toHaveBeenCalled()
          expect(diagnostics.snapshot().events.at(-1)?.error_code).toBe('image_fetch_unavailable')
        }
      } finally {
        session.dispose()
        bridge.disconnect()
        await request
      }
    },
  )
})
