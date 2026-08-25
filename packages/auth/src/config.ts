export interface AuthConfig {
  authorizationEndpoint: string
  authorizationResponseIssuer: string
  callbackEndpoint: string
  refreshEndpoint: string
  clientId: string
  redirectUri: string
  refreshFixedCode?: string
  scope: string
  transactionTtlMs: number
}

export const DEFAULT_AUTH_CONFIG: AuthConfig = {
  authorizationEndpoint: 'https://auth.wispaper.ai/oidc/auth',
  authorizationResponseIssuer: 'https://auth.wispaper.ai/oidc',
  callbackEndpoint: 'https://gateway.wispaper.ai/api/v1/auth/user/callback',
  refreshEndpoint: 'https://gateway.wispaper.ai/api/v1/auth/user/refresh',
  clientId: 'i9au2rbqzktme4runr9gy',
  redirectUri: 'wiswork://oauth/callback',
  scope: 'openid profile email offline_access',
  transactionTtlMs: 10 * 60_000,
}

export function authConfigFromEnv(env: NodeJS.ProcessEnv = process.env): AuthConfig {
  const refreshFixedCode = env.WISWORK_OAUTH_REFRESH_FIXED_CODE
  return {
    ...DEFAULT_AUTH_CONFIG,
    authorizationEndpoint:
      env.WISWORK_OAUTH_AUTHORIZATION_URL ?? DEFAULT_AUTH_CONFIG.authorizationEndpoint,
    callbackEndpoint: env.WISWORK_OAUTH_CALLBACK_URL ?? DEFAULT_AUTH_CONFIG.callbackEndpoint,
    refreshEndpoint: env.WISWORK_OAUTH_REFRESH_URL ?? DEFAULT_AUTH_CONFIG.refreshEndpoint,
    clientId: env.WISWORK_OAUTH_CLIENT_ID ?? DEFAULT_AUTH_CONFIG.clientId,
    ...(refreshFixedCode ? { refreshFixedCode } : {}),
  }
}
