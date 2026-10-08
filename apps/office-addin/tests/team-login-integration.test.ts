import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it, vi } from 'vitest'
import { createBrowserAuthClient, createMemorySessionStore } from '@wiswork/auth/browser'
import { createOfficeTeamConnection } from '../src/agent/team-connection.js'
import {
  createOfficeTeamLoginDialog,
  type OfficeTeamLoginDialogRuntime,
} from '../src/agent/team-login-dialog.js'
import { officeTeamAuthConfig } from '../src/agent/team-auth-config.js'
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'
import { createOfficeWorkspaceUi } from '../src/App.js'
import type { OfficeRelaySession, OfficeRelaySnapshot } from '../src/relay/session.js'
import type { StructuredProposalController } from '../src/agent/proposal-controller.js'
import { createPresentationService } from '../../shell/src/main/presentation-service.js'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'

it('connects browser auth, an Office-shaped dialog, team transport and actual PC proposals; logout clears a pending shared write', async () => {
  const root = mkdtempSync(join(tmpdir(), 'office-team-login-'))
  const service = createPresentationService({ userDataPath: root })
  const origin = 'https://office.example'
  const config = officeTeamAuthConfig(
    {
      VITE_WISWORK_TEAM_CLIENT_ID: 'registered-fixture-client',
      VITE_WISWORK_TEAM_REDIRECT_URI: `${origin}/team-auth-callback.html`,
    },
    origin,
  )!
  const exchange = vi.fn(async (input: string | URL | Request) => {
    const url = new URL(String(input))
    expect(url.origin).toBe('https://gateway.wispaper.ai')
    expect(url.searchParams.get('redirect_uri')).toBe(config.redirectUri)
    return Response.json({
      success: true,
      data: {
        token: 'synthetic-access-token',
        refresh_token: 'synthetic-refresh-token',
        user_id: 'fixture-account',
      },
    })
  })
  const auth = createBrowserAuthClient({
    config,
    store: createMemorySessionStore(),
    fetch: exchange,
  })
  const close = vi.fn()
  const dialogRuntime: OfficeTeamLoginDialogRuntime = {
    context: {
      requirements: { isSetSupported: () => true },
      ui: {
        displayDialogAsync(url, _options, callback) {
          const authorization = new URL(decodeURIComponent(new URL(url).hash.slice(1)))
          callback({
            status: 'succeeded',
            value: {
              close,
              addEventHandler(name: string, handler: (event: unknown) => void) {
                if (name === 'message')
                  queueMicrotask(() =>
                    handler({
                      origin,
                      message: JSON.stringify({
                        type: 'team.auth.callback',
                        url: `${config.redirectUri}?code=fixture-code&state=${authorization.searchParams.get('state')}&iss=${encodeURIComponent(config.authorizationResponseIssuer)}`,
                      }),
                    }),
                  )
              },
            },
          })
        },
      },
    },
    EventType: { DialogMessageReceived: 'message', DialogEventReceived: 'event' },
    AsyncResultStatus: { Succeeded: 'succeeded' },
  }
  let relayState: OfficeRelaySnapshot = { status: 'offline' }
  const listeners = new Set<() => void>()
  const owner = 'a'.repeat(64)
  const privateRequest = async (body: unknown, signal?: AbortSignal) =>
    new Response(
      Buffer.from(await service(body, signal ?? new AbortController().signal)).toString('utf8'),
    )
  const connection = createOfficeTeamConnection({
    auth,
    loginDialog: createOfficeTeamLoginDialog({
      ...config,
      addinOrigin: origin,
      runtime: dialogRuntime,
    }),
    onUnavailable: () => runtime?.team?.clear(),
    createSession({ capabilities, getTeamAccessToken }) {
      expect(capabilities).toEqual(['presentation-team.v1'])
      return {
        snapshot: () => relayState,
        subscribe(listener) {
          listeners.add(listener)
          return () => listeners.delete(listener)
        },
        async connect() {
          relayState = { status: 'connected', capabilities: ['presentation-team.v1'] }
          listeners.forEach((l) => l())
        },
        disconnect() {
          relayState = { status: 'offline' }
          listeners.forEach((l) => l())
        },
        async capabilityFetch(capability, body, signal) {
          expect(capability).toBe('presentation-team.v1')
          expect(await getTeamAccessToken()).toBe('synthetic-access-token')
          const raw = Buffer.from(
            await service(body, signal ?? new AbortController().signal, {
              version: 1,
              actorSubject: owner,
              pcSubject: owner,
            }),
          ).toString('utf8')
          return new Response(raw, { status: JSON.parse(raw).error ? 400 : 200 })
        },
        authenticatedFetch: vi.fn(),
        sendDiagnostic: vi.fn(),
        diagnosticSessionId: () => undefined,
        forget: vi.fn(async () => undefined),
      } as OfficeRelaySession
    },
  })
  vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } })
  const qa = vi.fn(async () => {})
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      request: privateRequest,
      documentId: async () => 'login-fixture-document',
      teamAvailable: connection.available,
      teamRequest: connection.request,
      invalidateQa: qa,
      lastProject: () => benchmarkPlan().projectId,
      rememberProject: async () => {},
    },
  })
  try {
    const plan = benchmarkPlan()
    await privateRequest({
      operation: 'save_plan',
      documentId: 'login-fixture-document',
      projectId: plan.projectId,
      expectedRevision: 0,
      plan,
    })
    const ui = createOfficeWorkspaceUi(runtime, undefined, undefined, undefined, connection)
    expect(ui.teamConnection).toBe(connection)
    await connection.signIn()
    expect(connection.available()).toBe(true)
    expect(exchange).toHaveBeenCalledTimes(1)
    expect(close).toHaveBeenCalledOnce()
    await ui.team!.refresh()
    expect(ui.team!.snapshot().identity?.actorSubject).toBe(owner)
    const proposal = await ui.team!.create(plan.projectId, 1)
    const proposals = runtime.proposals as StructuredProposalController
    expect(proposals.pending()?.id).toBe(proposal)
    await connection.signOut()
    expect(proposals.pending()).toBeUndefined()
    expect(ui.team!.snapshot().identity).toBeUndefined()
    expect(connection.available()).toBe(false)
    expect(await auth.getAccessToken()).toBeNull()
    expect(JSON.stringify(connection.snapshot())).not.toContain('synthetic-')
    expect(qa).not.toHaveBeenCalled()
  } finally {
    connection.dispose()
    runtime.dispose()
    vi.unstubAllGlobals()
    rmSync(root, { recursive: true, force: true })
  }
})
