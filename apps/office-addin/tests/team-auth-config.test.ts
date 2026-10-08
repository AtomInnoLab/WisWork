import { expect, it } from 'vitest'
import { officeTeamAuthConfig } from '../src/agent/team-auth-config.js'
import { deploymentConnectOrigins } from '../build-config.js'
const origin = 'https://office.example'
const env = {
  VITE_WISWORK_TEAM_CLIENT_ID: 'registered-web-client',
  VITE_WISWORK_TEAM_REDIRECT_URI: `${origin}/team-auth-callback.html`,
  VITE_WISWORK_ADDIN_ORIGIN: origin,
}
it('disables unconfigured team login and requires both registered web settings', () => {
  expect(officeTeamAuthConfig({}, origin)).toBeUndefined()
  expect(() => officeTeamAuthConfig({ VITE_WISWORK_TEAM_CLIENT_ID: 'x' }, origin)).toThrow(
    'invalid_office_team_auth_config',
  )
})
it('requires exact HTTPS same-origin callback and explicit public client', () => {
  expect(officeTeamAuthConfig(env, origin)).toMatchObject({
    clientId: 'registered-web-client',
    redirectUri: env.VITE_WISWORK_TEAM_REDIRECT_URI,
  })
  for (const value of [
    'http://office.example/team-auth-callback.html',
    'https://other.example/team-auth-callback.html',
    `${origin}/wrong.html`,
    `${origin}/team-auth-callback.html?code=x`,
    `${origin}/team-auth-callback.html#x`,
  ])
    expect(() =>
      officeTeamAuthConfig({ ...env, VITE_WISWORK_TEAM_REDIRECT_URI: value }, origin),
    ).toThrow('invalid_office_team_auth_config')
})
it('limits gateway CSP additions to explicitly configured login', () => {
  expect(deploymentConnectOrigins({})).not.toContain('gateway.wispaper.ai')
  expect(deploymentConnectOrigins(env)).toBe(
    'wss://office.8-216-134-194.sslip.io https://gateway.wispaper.ai',
  )
})
