export interface AuthConfig {
  authorizationEndpoint: string
  authorizationResponseIssuer: string
  callbackEndpoint: string
  refreshEndpoint: string
  clientId: string
  redirectUri: string
  refreshFixedCode: string
  scope: string
  transactionTtlMs: number
}

export const DEFAULT_AUTH_CONFIG: AuthConfig = {
  authorizationEndpoint: 'https://auth.dev.wispaper.ai/oidc/auth',
  authorizationResponseIssuer: 'https://auth.dev.wispaper.ai/oidc',
  callbackEndpoint: 'https://gateway.dev.wispaper.ai/api/v1/auth/user/callback',
  refreshEndpoint: 'https://gateway.dev.wispaper.ai/api/v1/auth/user/refresh',
  clientId: 'y3xpwx3ytskxf66p0wztm',
  redirectUri: 'wiswork://oauth/callback',
  refreshFixedCode: '1234567890',
  scope: 'openid profile email offline_access',
  transactionTtlMs: 10 * 60_000,
}

export function authConfigFromEnv(env: NodeJS.ProcessEnv = process.env): AuthConfig {
  return {
    ...DEFAULT_AUTH_CONFIG,
    authorizationEndpoint:
      env.WISWORK_OAUTH_AUTHORIZATION_URL ?? DEFAULT_AUTH_CONFIG.authorizationEndpoint,
    callbackEndpoint: env.WISWORK_OAUTH_CALLBACK_URL ?? DEFAULT_AUTH_CONFIG.callbackEndpoint,
    refreshEndpoint: env.WISWORK_OAUTH_REFRESH_URL ?? DEFAULT_AUTH_CONFIG.refreshEndpoint,
    clientId: env.WISWORK_OAUTH_CLIENT_ID ?? DEFAULT_AUTH_CONFIG.clientId,
    refreshFixedCode: env.WISWORK_OAUTH_REFRESH_FIXED_CODE ?? DEFAULT_AUTH_CONFIG.refreshFixedCode,
  }
}
