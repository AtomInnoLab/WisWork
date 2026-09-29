import { DEFAULT_AUTH_CONFIG, type AuthConfig } from '../../../../packages/auth/src/config.js'
/** Public registration settings only; credentials remain in the in-memory auth client. */
export function officeTeamAuthConfig(
  env: Record<string, string | undefined>,
  addinOrigin: string,
): AuthConfig | undefined {
  const clientId = env.VITE_WISWORK_TEAM_CLIENT_ID
  const redirectUri = env.VITE_WISWORK_TEAM_REDIRECT_URI
  if (!clientId && !redirectUri) return undefined
  try {
    const origin = new URL(addinOrigin)
    const redirect = new URL(redirectUri ?? '')
    if (
      !clientId ||
      !/^[A-Za-z0-9_.-]{1,256}$/.test(clientId) ||
      origin.protocol !== 'https:' ||
      origin.origin !== addinOrigin ||
      redirect.origin !== addinOrigin ||
      redirect.pathname !== '/team-auth-callback.html' ||
      redirect.search ||
      redirect.hash ||
      redirect.username ||
      redirect.password ||
      redirect.href !== redirectUri
    )
      throw Error('invalid')
    return { ...DEFAULT_AUTH_CONFIG, clientId, redirectUri: redirect.href }
  } catch {
    throw Error('invalid_office_team_auth_config')
  }
}
