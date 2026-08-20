const FIXED_MESSAGES_URL = 'https://wisusage.dev.atominnolab.com/v1/messages'

type BuildEnv = Record<string, string | undefined>

export interface DeploymentConfig {
  addinOrigin: string
  authorizationOrigin: string
  tokenOrigin: string
  issuerOrigin: string
}

function secureUrl(value: string | undefined): URL | undefined {
  if (!value) return undefined
  try {
    const url = new URL(value)
    if (url.protocol !== 'https:' || url.username || url.password || url.hostname.includes('*')) {
      return undefined
    }
    return url
  } catch {
    return undefined
  }
}

export function deploymentConfig(env: BuildEnv): DeploymentConfig | undefined {
  const authorization = secureUrl(env.VITE_WISWORK_AUTHORIZATION_URL)
  const token = secureUrl(env.VITE_WISWORK_TOKEN_URL)
  const callback = secureUrl(env.VITE_WISWORK_CALLBACK_URL)
  const issuer = secureUrl(env.VITE_WISWORK_ISSUER)
  const messages = secureUrl(env.VITE_WISWORK_MESSAGES_URL)
  if (
    !authorization ||
    !token ||
    !callback ||
    !issuer ||
    !messages ||
    !env.VITE_WISWORK_CLIENT_ID?.trim() ||
    callback.pathname !== '/oauth/callback' ||
    callback.search ||
    callback.hash ||
    messages.href !== FIXED_MESSAGES_URL
  ) {
    return undefined
  }
  return {
    addinOrigin: callback.origin,
    authorizationOrigin: authorization.origin,
    tokenOrigin: token.origin,
    issuerOrigin: issuer.origin,
  }
}

export function deploymentConnectOrigins(config: DeploymentConfig): string {
  return [
    ...new Set([
      config.authorizationOrigin,
      config.tokenOrigin,
      config.issuerOrigin,
      new URL(FIXED_MESSAGES_URL).origin,
    ]),
  ].join(' ')
}

export function renderDeploymentManifest(template: string, config: DeploymentConfig): string {
  const domains = [
    ...new Set([
      config.addinOrigin,
      config.authorizationOrigin,
      config.tokenOrigin,
      config.issuerOrigin,
      new URL(FIXED_MESSAGES_URL).origin,
    ]),
  ]
    .map((origin) => `    <AppDomain>${origin}</AppDomain>`)
    .join('\n')
  return template
    .replace(
      '<!-- DEVELOPMENT-ONLY MANIFEST: production builds generate origin-specific dist/manifest.xml. -->\n',
      '',
    )
    .replaceAll('https://localhost:3000', config.addinOrigin)
    .replace(
      / {2}<AppDomains>[\s\S]*?<\/AppDomains>/,
      `  <AppDomains>\n${domains}\n  </AppDomains>`,
    )
}

export function rewriteOAuthCallbackRequest(requestUrl: string | undefined): string | undefined {
  if (!requestUrl) return requestUrl
  const queryIndex = requestUrl.indexOf('?')
  const pathname = queryIndex === -1 ? requestUrl : requestUrl.slice(0, queryIndex)
  if (pathname !== '/oauth/callback') return requestUrl
  return `/oauth/callback.html${queryIndex === -1 ? '' : requestUrl.slice(queryIndex)}`
}
